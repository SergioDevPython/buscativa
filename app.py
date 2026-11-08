import datetime
import io
import urllib.parse
import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pymongo import MongoClient

# Configuração da Página
st.set_page_config(
    page_title="Gestão Escolar - Tempo Integral", page_icon="📚", layout="wide"
)

# -----------------------------------------------------------------------------
# 1. CONFIGURAÇÃO DO BANCO DE DADOS (MongoDB via secrets.toml)
# -----------------------------------------------------------------------------
MONGO_URI = st.secrets["mongo"]["uri"]


@st.cache_resource
def get_db():
    client = MongoClient(MONGO_URI)
    return client["gestao_escolar"]


db = get_db()
colecao_usuarios = db["usuarios"]
colecao_alunos = db["alunos"]
colecao_turmas = db["turmas"]
colecao_chamadas = db["chamadas"]
colecao_envios = db["historico_envios"]

# Carga inicial de usuário padrão Master se o banco estiver vazio
if colecao_usuarios.count_documents({}) == 0:
    colecao_usuarios.insert_one(
        {"usuario": "master", "senha": "123", "perfil": "master"}
    )


# -----------------------------------------------------------------------------
# 2. AUTENTICAÇÃO E SESSÃO
# -----------------------------------------------------------------------------
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["perfil"] = None
    st.session_state["usuario"] = None


def login(usuario, senha):
    user = colecao_usuarios.find_one({"usuario": usuario, "senha": senha})
    if user:
        st.session_state["autenticado"] = True
        st.session_state["perfil"] = user["perfil"]
        st.session_state["usuario"] = user["usuario"]
        return True
    return False


def logout():
    st.session_state["autenticado"] = False
    st.session_state["perfil"] = None
    st.session_state["usuario"] = None
    st.rerun()


# -----------------------------------------------------------------------------
# 3. MENSAGEM DINÂMICA DE WHATSAPP
# -----------------------------------------------------------------------------
def gerar_mensagem_whatsapp(nome_aluno, faltas, nome_escola="Nossa Escola"):
    if faltas >= 7:
        texto = (
            f"🚨 *AVISO IMPORTANTE - BUSCA ATIVA*\n\n"
            f"Olá! Entramos em contato da *{nome_escola}* devido ao alto índice de ausências do(a) estudante *{nome_aluno}* "
            f"no programa de Tempo Integral (atualmente com *{faltas} faltas* acumuladas).\n\n"
            f"Solicitamos que o responsável entre em contato conosco com urgência para justificativa ou alinhamento."
        )
    elif 5 <= faltas <= 6:
        texto = (
            f"⚠️ *Aviso de Frequência Escolar*\n\n"
            f"Olá! Entramos em contato da *{nome_escola}* para informar que o(a) aluno(a) *{nome_aluno}* "
            f"já possui *{faltas} faltas* acumuladas no Tempo Integral.\n\n"
            f"Gostaríamos de saber se está tudo bem ou se há algum motivo específico para as ausências recentes."
        )
    else:
        texto = (
            f"Olá! Entramos em contato da *{nome_escola}* referente ao acompanhamento de rotina de frequência "
            f"do(a) estudante *{nome_aluno}* no Tempo Integral. Poderia nos dar um breve retorno quando possível?"
        )

    return urllib.parse.quote(texto)


# -----------------------------------------------------------------------------
# 4. GERAÇÃO DE EXCEL FORMATADO E PROFISSIONAL
# -----------------------------------------------------------------------------
def gerar_excel_formatado(
    nome_escola, turma_sel, periodo_str, df_dados
):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_dados.to_excel(
            writer, index=False, sheet_name="Frequencia", startrow=4
        )

        workbook = writer.book
        worksheet = writer.sheets["Frequencia"]

        # Cabeçalho Principal (Informações da Escola)
        worksheet["A1"] = f"Escola: {nome_escola} (Tempo Integral)"
        worksheet["A1"].font = Font(
            name="Arial", size=14, bold=True, color="1F4E78"
        )

        worksheet["A2"] = f"Relatório da Turma: {turma_sel}"
        worksheet["A2"].font = Font(name="Arial", size=11, bold=True)

        worksheet["A3"] = f"Período: {periodo_str}"
        worksheet["A3"].font = Font(name="Arial", size=10, italic=True)

        # Estilos para o cabeçalho da tabela (Linha 5)
        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        # Formatação dos cabeçalhos das colunas
        for col_num in range(1, len(df_dados.columns) + 1):
            cell = worksheet.cell(row=5, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border

        # Cores para o status de faltas
        fill_verde = PatternFill(
            start_color="D4EDDA", end_color="D4EDDA", fill_type="solid"
        )
        font_verde = Font(name="Arial", size=10, color="155724", bold=True)

        fill_amarelo = PatternFill(
            start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"
        )
        font_amarelo = Font(name="Arial", size=10, color="856404", bold=True)

        fill_vermelho = PatternFill(
            start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"
        )
        font_vermelho = Font(name="Arial", size=10, color="721C24", bold=True)

        # Formatação das linhas de dados
        for row_idx in range(6, len(df_dados) + 6):
            for col_idx in range(1, len(df_dados.columns) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.border = thin_border
                cell.font = Font(name="Arial", size=10)

                # Alinhamentos específicos por coluna
                col_name = df_dados.columns[col_idx - 1]
                if "Nome" in col_name:
                    cell.alignment = align_left
                else:
                    cell.alignment = align_center

                # Destacar a coluna de Situação de Faltas
                if "Situação" in col_name or "Faltas" in col_name:
                    val_str = str(cell.value)
                    if "Normal" in val_str or (
                        isinstance(cell.value, int) and cell.value <= 4
                    ):
                        cell.fill = fill_verde
                        cell.font = font_verde
                    elif "Atenção" in val_str or (
                        isinstance(cell.value, int)
                        and 5 <= cell.value <= 6
                    ):
                        cell.fill = fill_amarelo
                        cell.font = font_amarelo
                    elif "BUSCA ATIVA" in val_str or (
                        isinstance(cell.value, int) and cell.value >= 7
                    ):
                        cell.fill = fill_vermelho
                        cell.font = font_vermelho

        # Ajuste automático da largura das colunas
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row >= 5:  # Considera a partir do cabeçalho
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max(max_len + 5, 15)

    output.seek(0)
    return output


# -----------------------------------------------------------------------------
# 5. FUNÇÃO DE GERAÇÃO DE PDF
# -----------------------------------------------------------------------------
def gerar_pdf(nome_escola, turma_sel, data_inicial, data_atual, df_dados):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    largura, altura = letter

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, altura - 50, f"Escola: {nome_escola} (Tempo Integral)")

    c.setFont("Helvetica", 12)
    c.drawString(
        50,
        altura - 70,
        f"Relatório de Frequência e Merenda - Turma: {turma_sel}",
    )
    c.drawString(
        50,
        altura - 90,
        f"Período: {data_inicial.strftime('%d/%m/%Y')} até {data_atual.strftime('%d/%m/%Y')}",
    )
    c.line(50, altura - 100, largura - 50, altura - 100)

    y = altura - 130
    c.setFont("Helvetica-Bold", 10)
    c.drawString(50, y, "Nome do Aluno")
    c.drawString(260, y, "Faltas Acum.")
    c.drawString(340, y, "Situação")
    c.drawString(440, y, "Alimenta na Escola")

    y -= 20
    c.setFont("Helvetica", 10)

    for index, row in df_dados.iterrows():
        if y < 50:
            c.showPage()
            y = altura - 50
            c.setFont("Helvetica", 10)

        c.drawString(50, y, str(row["Nome do Aluno"])[:30])
        c.drawString(260, y, str(row["Qtd. Faltas"]))
        c.drawString(340, y, str(row["Situação Faltas"]))
        c.drawString(440, y, str(row["Alimenta na Escola"]))
        y -= 18

    c.save()
    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# 6. TELA DE LOGIN
# -----------------------------------------------------------------------------
if not st.session_state["autenticado"]:
    st.title("🔒 Sistema de Controle Escolar - Login")
    with st.form("form_login"):
        user_input = st.text_input("Usuário")
        pass_input = st.text_input("Senha", type="password")
        btn_entrar = st.form_submit_button("Entrar")

        if btn_entrar:
            if login(user_input, pass_input):
                st.success("Login realizado com sucesso!")
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")
    st.stop()

# -----------------------------------------------------------------------------
# 7. PAINEL PRINCIPAL
# -----------------------------------------------------------------------------
st.sidebar.title(f"👤 {st.session_state['usuario']}")
st.sidebar.caption(f"Perfil: **{st.session_state['perfil'].upper()}**")
if st.sidebar.button("Sair (Logout)"):
    logout()

st.title("📋 Controle Escolar - Tempo Integral")

# -----------------------------------------------------------------------------
# ESTRUTURA DE ABAS
# -----------------------------------------------------------------------------
aba0, aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs(
    [
        "📈 Dashboard Geral",
        "📊 Chamada Diária",
        "📄 Relatórios & Impressão",
        "🏫 Turmas & Alunos",
        "⚙️ Ficha do Aluno",
        "📨 Histórico de Zaps",
        "👥 Gerenciar Usuários",
    ]
)

# -----------------------------------------------------------------------------
# ABA 0: DASHBOARD GERAL
# -----------------------------------------------------------------------------
with aba0:
    st.subheader("Visão Geral da Escola")

    total_alunos = colecao_alunos.count_documents({})
    total_turmas = colecao_turmas.count_documents({})
    alunos_busca_ativa = colecao_alunos.count_documents({"faltas": {"$gte": 7}})
    alunos_nao_come = colecao_alunos.count_documents({"come_escola": False})

    col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
    col_kpi1.metric("Total de Alunos", total_alunos)
    col_kpi2.metric("Total de Turmas", total_turmas)
    col_kpi3.metric(
        "🔴 Em Busca Ativa (7+ faltas)",
        alunos_busca_ativa,
        delta_color="inverse",
    )
    col_kpi4.metric("Alunos que não comem na escola", alunos_nao_come)

    st.divider()

    if total_alunos > 0:
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("### 📊 Status de Faltas Geral")
            normal = colecao_alunos.count_documents({"faltas": {"$lte": 4}})
            atencao = colecao_alunos.count_documents(
                {"faltas": {"$gt": 4, "$lt": 7}}
            )
            critico = alunos_busca_ativa

            fig_kpi, ax_kpi = plt.subplots(figsize=(5, 3))
            ax_kpi.bar(
                ["Normal (0-4)", "Atenção (5-6)", "Busca Ativa (7+)"],
                [normal, atencao, critico],
                color=["#28a745", "#ffc107", "#dc3545"],
            )
            ax_kpi.set_ylabel("Qtd de Alunos")
            st.pyplot(fig_kpi)

        with col_g2:
            st.markdown("### 🥗 Adesão à Merenda Escolar")
            come = total_alunos - alunos_nao_come
            fig_m, ax_m = plt.subplots(figsize=(5, 3))
            ax_m.pie(
                [come, alunos_nao_come],
                labels=["Alimentam-se", "Não se alimentam"],
                autopct="%1.1f%%",
                colors=["#17a2b8", "#6c757d"],
                startangle=90,
            )
            ax_m.axis("equal")
            st.pyplot(fig_m)
    else:
        st.info("Cadastre alunos para visualizar os gráficos do dashboard.")

# -----------------------------------------------------------------------------
# ABA 1: CHAMADA DIÁRIA
# -----------------------------------------------------------------------------
with aba1:
    st.subheader("Registro de Frequência Diária")

    col_escola, col_d1, col_d2 = st.columns([2, 1.5, 1.5])
    nome_escola_zap = col_escola.text_input(
        "Nome da Escola (Para a mensagem do WhatsApp)",
        "Escola Municipal de Tempo Integral",
        key="input_nome_escola_zap",
    )
    turmas = [t["nome"] for t in colecao_turmas.find().sort("nome", 1)]
    turma_sel = col_d1.selectbox("Selecione a Turma", turmas if turmas else [""])
    data_chamada = col_d2.date_input(
        "Data da Chamada", datetime.date.today(), key="data_chamada_input"
    )

    if not turmas:
        st.info("Cadastre uma turma na aba 'Turmas & Alunos'.")
    else:
        alunos = list(colecao_alunos.find({"turma": turma_sel}).sort("nome", 1))

        if alunos:
            st.markdown(
                "**Legenda:** 🟢 **0-4 faltas:** Normal | 🟡 **5-6 faltas:** Atenção | 🔴 **7+ faltas:** Busca Ativa"
            )
            st.divider()

            col_n, col_f, col_a, col_m, col_w = st.columns([2.5, 1.2, 1.8, 2, 1.5])
            col_n.markdown("**Nome do Aluno**")
            col_f.markdown("**Acumulado**")
            col_a.markdown("**Lançar Falta**")
            col_m.markdown("**Alimenta na Escola?**")
            col_w.markdown("**Ação WhatsApp**")
            st.divider()

            for aluno in alunos:
                c1, c2, c3, c4, c5 = st.columns([2.5, 1.2, 1.8, 2, 1.5])
                aluno_id = str(aluno["_id"])
                faltas = aluno.get("faltas", 0)
                come_escola = aluno.get("come_escola", True)
                telefone = aluno.get("telefone", "")

                if faltas <= 4:
                    bg_color, txt_color, status = "#d4edda", "#155724", ""
                elif 5 <= faltas <= 6:
                    bg_color, txt_color, status = "#fff3cd", "#856404", "⚠️"
                else:
                    bg_color, txt_color, status = "#f8d7da", "#721c24", "🔴 BUSCA ATIVA"

                c1.markdown(
                    f"""
                    <div style="background-color: {bg_color}; color: {txt_color}; 
                                padding: 6px 10px; border-radius: 6px; font-weight: bold; font-size: 14px;">
                        {aluno['nome']} {status}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                c2.markdown(
                    f"<h4 style='text-align: center; margin: 0;'>{faltas}</h4>",
                    unsafe_allow_html=True,
                )

                col_sub1, col_sub2 = c3.columns(2)
                if col_sub1.button("➕", key=f"inc_{aluno_id}"):
                    colecao_alunos.update_one(
                        {"_id": aluno["_id"]}, {"$inc": {"faltas": 1}}
                    )
                    colecao_chamadas.insert_one(
                        {
                            "aluno_id": aluno_id,
                            "data": str(data_chamada),
                            "falta": True,
                        }
                    )
                    st.rerun()
                if col_sub2.button("➖", key=f"dec_{aluno_id}"):
                    if faltas > 0:
                        colecao_alunos.update_one(
                            {"_id": aluno["_id"]}, {"$inc": {"faltas": -1}}
                        )
                        st.rerun()

                status_merenda = c4.checkbox(
                    "Alimenta-se", value=come_escola, key=f"merenda_{aluno_id}"
                )
                if status_merenda != come_escola:
                    colecao_alunos.update_one(
                        {"_id": aluno["_id"]},
                        {"$set": {"come_escola": status_merenda}},
                    )
                    st.rerun()

                if telefone:
                    tel_limpo = "".join(filter(str.isdigit, str(telefone)))
                    msg_encoded = gerar_mensagem_whatsapp(
                        aluno["nome"], faltas, nome_escola=nome_escola_zap
                    )
                    link_zap = f"https://wa.me/55{tel_limpo}?text={msg_encoded}"

                    if faltas >= 7:
                        label_botao = "🚨 Busca Ativa"
                    elif 5 <= faltas <= 6:
                        label_botao = "⚠️ Alerta"
                    else:
                        label_botao = "💬 Zap"

                    if c5.button(label_botao, key=f"btn_zap_{aluno_id}"):
                        colecao_envios.insert_one(
                            {
                                "aluno_id": aluno_id,
                                "nome_aluno": aluno["nome"],
                                "turma": aluno["turma"],
                                "faltas_momento": faltas,
                                "tipo_mensagem": "Busca Ativa"
                                if faltas >= 7
                                else ("Alerta" if faltas >= 5 else "Rotina"),
                                "usuario_responsavel": st.session_state["usuario"],
                                "data_envio": datetime.datetime.now(),
                            }
                        )
                        components.html(
                            f"""
                            <script>
                                window.open("{link_zap}", "_blank");
                            </script>
                        """,
                            height=0,
                        )
                        st.success(f"Registro salvo para {aluno['nome']}!")
                else:
                    c5.caption("Sem tel.")

                st.markdown(
                    "<hr style='margin: 4px 0; border-color: #eee;'>",
                    unsafe_allow_html=True,
                )
        else:
            st.info(f"Nenhum aluno cadastrado na turma {turma_sel}.")

# -----------------------------------------------------------------------------
# ABA 2: RELATÓRIOS & EXPORTAÇÃO EXCEL FORMATADO
# -----------------------------------------------------------------------------
with aba2:
    st.subheader("Relatórios, Gráficos e Impressão")
    tab_rel1, tab_rel2 = st.tabs(
        ["📄 Visualizar & Exportar", "🖨️ Modo Impressão Direta"]
    )

    with tab_rel1:
        col1, col2 = st.columns(2)
        nome_escola = col1.text_input(
            "Nome da Escola", "Escola Municipal de Tempo Integral"
        )
        turmas_rel = [t["nome"] for t in colecao_turmas.find().sort("nome", 1)]
        turma_rel_sel = col2.selectbox(
            "Turma", turmas_rel if turmas_rel else ["Todas"]
        )

        col3, col4 = st.columns(2)
        data_inicial = col3.date_input(
            "Data Inicial",
            datetime.date.today() - datetime.timedelta(days=30),
            key="di_rel",
        )
        data_atual = datetime.date.today()
        col4.text_input(
            "Data Atual", value=data_atual.strftime("%d/%m/%Y"), disabled=True
        )

        if st.button("Gerar Relatório Analítico"):
            query = {} if turma_rel_sel == "Todas" else {"turma": turma_rel_sel}
            alunos_rel = list(colecao_alunos.find(query, {"_id": 0}).sort("nome", 1))

            if alunos_rel:
                df = pd.DataFrame(alunos_rel)
                df["Alimenta na Escola"] = df["come_escola"].apply(
                    lambda x: "Sim" if x else "Não"
                )
                df["Situação Faltas"] = df["faltas"].apply(
                    lambda x: "🟢 Normal"
                    if x <= 4
                    else ("🟡 Atenção" if x <= 6 else "🔴 BUSCA ATIVA")
                )

                df_exibicao = df.rename(
                    columns={
                        "nome": "Nome do Aluno",
                        "turma": "Turma",
                        "faltas": "Qtd. Faltas",
                    }
                )[
                    [
                        "Nome do Aluno",
                        "Turma",
                        "Qtd. Faltas",
                        "Situação Faltas",
                        "Alimenta na Escola",
                    ]
                ]

                st.session_state["df_relatorio_atual"] = df_exibicao
                st.session_state["meta_escola"] = nome_escola
                st.session_state["meta_turma"] = turma_rel_sel
                periodo_str = f"{data_inicial.strftime('%d/%m/%Y')} até {data_atual.strftime('%d/%m/%Y')}"
                st.session_state["meta_periodo"] = periodo_str

                st.divider()
                st.markdown(f"### 🏫 {nome_escola}")
                st.markdown(
                    f"**Turma:** {turma_rel_sel} | **Período:** {periodo_str}"
                )
                st.dataframe(df_exibicao, use_container_width=True)

                fig_r, (ax_r1, ax_r2) = plt.subplots(1, 2, figsize=(12, 4))
                contagem_sit = df_exibicao["Situação Faltas"].value_counts()
                cores = [
                    "#28a745"
                    if "Normal" in item
                    else ("#ffc107" if "Atenção" in item else "#dc3545")
                    for item in contagem_sit.index
                ]
                contagem_sit.plot(kind="bar", ax=ax_r1, color=cores)
                ax_r1.set_title("Alunos por Situação")
                ax_r1.tick_params(axis="x", rotation=10)

                contagem_m = df["Alimenta na Escola"].value_counts()
                contagem_m.plot(
                    kind="pie",
                    ax=ax_r2,
                    autopct="%1.1f%%",
                    colors=["#17a2b8", "#e83e8c"],
                )
                ax_r2.set_title("Merenda Escolar")
                ax_r2.set_ylabel("")
                st.pyplot(fig_r)

                st.divider()
                col_b1, col_b2, col_b3 = st.columns(3)

                # Botão de Download Excel Formatado (.xlsx)
                excel_buffer = gerar_excel_formatado(
                    nome_escola, turma_rel_sel, periodo_str, df_exibicao
                )
                col_b1.download_button(
                    label="📊 Baixar Excel Formatado",
                    data=excel_buffer,
                    file_name=f"relatorio_{turma_rel_sel}_{data_atual}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                pdf_buffer = gerar_pdf(
                    nome_escola,
                    turma_rel_sel,
                    data_inicial,
                    data_atual,
                    df_exibicao,
                )
                col_b2.download_button(
                    label="📄 Baixar em PDF",
                    data=pdf_buffer,
                    file_name=f"relatorio_{turma_rel_sel}_{data_atual}.pdf",
                    mime="application/pdf",
                )
            else:
                st.warning("Nenhum registro encontrado.")

    with tab_rel2:
        st.subheader("🖨️ Impressão Direta")
        if "df_relatorio_atual" in st.session_state:
            html_tabela = st.session_state["df_relatorio_atual"].to_html(
                index=False, classes="table table-striped"
            )
            html_conteudo = f"""
            <html>
                <head><style>body{{font-family:Arial;padding:20px;}} table{{width:100%;border-collapse:collapse;margin-top:20px;}} th,td{{border:1px solid #ddd;padding:8px;text-align:left;}} th{{background-color:#f2f2f2;}}</style></head>
                <body>
                    <h2>🏫 {st.session_state.get('meta_escola', '')}</h2>
                    <h4>Relatório - Turma: {st.session_state.get('meta_turma', '')}</h4>
                    <p><b>Período:</b> {st.session_state.get('meta_periodo', '')}</p>
                    {html_tabela}
                    <script>window.print();</script>
                </body>
            </html>
            """
            if st.button("🖨️ Executar Impressão", type="primary"):
                components.html(html_conteudo, height=0)
        else:
            st.info(
                "Gere um relatório na aba anterior para habilitar a impressão."
            )

# -----------------------------------------------------------------------------
# ABA 3: TURMAS & ALUNOS
# -----------------------------------------------------------------------------
with aba3:
    st.subheader("Gestão de Turmas e Alunos")
    t_aba1, t_aba2, t_aba3 = st.tabs(
        ["🏫 Cadastrar Turmas", "👨‍🎓 Cadastrar Alunos", "📂 Importação em Massa"]
    )

    with t_aba1:
        with st.form("form_turma"):
            nome_nova_turma = st.text_input(
                "Nome/Código da Turma (Ex: 5º Ano Integral)"
            )
            if st.form_submit_button("Criar Turma") and nome_nova_turma:
                if colecao_turmas.find_one({"nome": nome_nova_turma.strip()}):
                    st.error("Turma já cadastrada.")
                else:
                    colecao_turmas.insert_one(
                        {"nome": nome_nova_turma.strip()}
                    )
                    st.success("Turma criada com sucesso!")
                    st.rerun()

        st.markdown("### Turmas Existentes")
        for t in colecao_turmas.find().sort("nome", 1):
            c_t1, c_t2 = st.columns([4, 1])
            c_t1.write(f"• **{t['nome']}**")
            if (
                st.session_state["perfil"] == "master"
                and c_t2.button("Excluir", key=f"dt_{t['_id']}")
            ):
                colecao_turmas.delete_one({"_id": t["_id"]})
                st.rerun()

    with t_aba2:
        if st.session_state["perfil"] != "master":
            st.warning(
                "Apenas usuários Master podem cadastrar alunos diretamente."
            )
        else:
            lista_t = [t["nome"] for t in colecao_turmas.find().sort("nome", 1)]
            if not lista_t:
                st.info("Cadastre uma turma primeiro.")
            else:
                with st.form("form_aluno"):
                    n_aluno = st.text_input("Nome Completo do Aluno")
                    t_aluno = st.selectbox("Turma", lista_t)
                    f_ini = st.number_input(
                        "Faltas Iniciais", min_value=0, value=0
                    )
                    c_esc = st.checkbox(
                        "Alimenta-se na escola",
                        value=True,
                        key="chk_cad_aluno",
                    )
                    tel = st.text_input(
                        "Telefone Responsável (WhatsApp com DDD)"
                    )
                    obs = st.text_area(
                        "Observações / Restrições Alimentares ou Médicas"
                    )

                    if st.form_submit_button("Cadastrar Aluno") and n_aluno:
                        colecao_alunos.insert_one(
                            {
                                "nome": n_aluno.strip(),
                                "turma": t_aluno,
                                "faltas": f_ini,
                                "come_escola": c_esc,
                                "telefone": tel.strip(),
                                "obs": obs.strip(),
                            }
                        )
                        st.success(f"Aluno {n_aluno} cadastrado!")
                        st.rerun()

    with t_aba3:
        st.markdown(
            "### Importação de Alunos via CSV\nO arquivo deve conter as colunas: `nome`, `turma`, `faltas`, `telefone`"
        )
        arquivo_csv = st.file_uploader("Enviar arquivo CSV", type=["csv"])
        if arquivo_csv is not None:
            df_imp = pd.read_csv(arquivo_csv)
            if st.button("Confirmar Importação em Massa"):
                for _, row in df_imp.iterrows():
                    colecao_alunos.insert_one(
                        {
                            "nome": str(row["nome"]).strip(),
                            "turma": str(row["turma"]).strip(),
                            "faltas": int(row.get("faltas", 0)),
                            "come_escola": True,
                            "telefone": str(row.get("telefone", "")),
                            "obs": "",
                        }
                    )
                st.success("Alunos importados com sucesso!")
                st.rerun()

# -----------------------------------------------------------------------------
# ABA 4: FICHA INDIVIDUAL DO ALUNO
# -----------------------------------------------------------------------------
with aba4:
    st.subheader("Ficha Detalhada do Aluno")
    todos_alunos = list(colecao_alunos.find().sort("nome", 1))

    if todos_alunos:
        aluno_dict = {f"{a['nome']} ({a['turma']})": a for a in todos_alunos}
        aluno_escolhido = st.selectbox(
            "Selecione o Aluno para Ver/Editar Ficha",
            list(aluno_dict.keys()),
        )
        dados_aluno = aluno_dict[aluno_escolhido]

        with st.form("form_edicao_aluno"):
            novo_tel = st.text_input(
                "Telefone WhatsApp", value=dados_aluno.get("telefone", "")
            )
            nova_obs = st.text_area(
                "Observações / Alergias / Restrições Alimentares",
                value=dados_aluno.get("obs", ""),
            )
            btn_salvar = st.form_submit_button("Salvar Alterações na Ficha")

            if btn_salvar:
                colecao_alunos.update_one(
                    {"_id": dados_aluno["_id"]},
                    {"$set": {"telefone": novo_tel, "obs": nova_obs}},
                )
                st.success("Ficha atualizada com sucesso!")
                st.rerun()
    else:
        st.info("Nenhum aluno cadastrado.")

# -----------------------------------------------------------------------------
# ABA 5: HISTÓRICO DE MENSAGENS ENVIADAS (Zaps)
# -----------------------------------------------------------------------------
with aba5:
    st.subheader("Histórico de Contatos via WhatsApp")
    envios = list(colecao_envios.find().sort("data_envio", -1).limit(100))

    if envios:
        df_envios = pd.DataFrame(envios)
        df_envios = df_envios.rename(
            columns={
                "data_envio": "Data e Hora",
                "nome_aluno": "Aluno",
                "turma": "Turma",
                "faltas_momento": "Faltas",
                "tipo_mensagem": "Tipo",
                "usuario_responsavel": "Usuário",
            }
        )
        st.dataframe(
            df_envios[
                [
                    "Data e Hora",
                    "Aluno",
                    "Turma",
                    "Faltas",
                    "Tipo",
                    "Usuário",
                ]
            ],
            use_container_width=True,
        )
    else:
        st.info("Nenhum envio de mensagem registrado até o momento.")

# -----------------------------------------------------------------------------
# ABA 6: GERENCIAR USUÁRIOS (Restrito ao Master)
# -----------------------------------------------------------------------------
with aba6:
    if st.session_state["perfil"] != "master":
        st.error(
            "⛔ Acesso Negado! Apenas o perfil **MASTER** pode gerenciar usuários."
        )
    else:
        st.subheader("Gerenciamento de Usuários do Sistema")

        with st.form("form_novo_usuario"):
            u_nome = st.text_input("Novo Nome de Usuário")
            u_senha = st.text_input("Senha", type="password")
            u_perfil = st.selectbox("Nível de Acesso", ["comum", "master"])
            if st.form_submit_button("Criar Usuário") and u_nome:
                if colecao_usuarios.find_one({"usuario": u_nome}):
                    st.error("Usuário já existe.")
                else:
                    colecao_usuarios.insert_one(
                        {
                            "usuario": u_nome.strip(),
                            "senha": u_senha.strip(),
                            "perfil": u_perfil,
                        }
                    )
                    st.success(f"Usuário {u_nome} criado!")
                    st.rerun()

        st.divider()
        st.markdown("### Usuários Cadastrados")
        for u in colecao_usuarios.find():
            c_u1, c_u2, c_u3 = st.columns([3, 2, 1])
            c_u1.write(f"• **{u['usuario']}** ({u['perfil'].upper()})")
            if u["usuario"] != "master":
                if c_u2.button(
                    "Excluir Usuário", key=f"del_user_{str(u['_id'])}"
                ):
                    colecao_usuarios.delete_one({"_id": u["_id"]})
                    st.rerun()