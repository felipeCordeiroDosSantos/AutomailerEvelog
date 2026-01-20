import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# --------------------------------------------------
# CONFIG STREAMLIT
# --------------------------------------------------
st.set_page_config(
    page_title="AutoMailer",
    layout="wide"
)

st.title("📮 Envio Automático de E-mails")

# --------------------------------------------------
# BASE DE E-MAILS DAS UNIDADES
# A = Unidade | B = Emails
# --------------------------------------------------
@st.cache_data
def carregar_emails_unidades():
    df = pd.read_excel("emails_unidades.xlsx", header=0)
    df.columns = ["Unidade", "Emails"]

    df["Unidade"] = df["Unidade"].astype(str).str.strip().str.upper()
    df["Emails"] = df["Emails"].astype(str).str.strip()

    return dict(zip(df["Unidade"], df["Emails"]))


emails_unidades = carregar_emails_unidades()

# --------------------------------------------------
# FORMULÁRIO
# --------------------------------------------------
col_form, _ = st.columns([1, 2])

with col_form:
    email_user = st.text_input(
        "E-mail remetente",
        placeholder="atendimento@evelog.com.br",
        key="email_user"
    )

    senha = st.text_input(
        "Senha",
        type="password",
        key="email_smtp"
    )


    uploaded = st.file_uploader(
        "Importar planilha",
        type=["xlsx", "xls", "csv"]
    )

# --------------------------------------------------
# PROCESSAMENTO DA PLANILHA
# --------------------------------------------------
if uploaded:

    # ---------------------------------
    # DETECÇÃO DE PLANILHA DE COLETA
    # (A2 == "ORDEM")
    # ---------------------------------
    if uploaded.name.endswith(".csv"):
        df_head = pd.read_csv(uploaded, header=1, usecols=[0], nrows=1)
    else:
        df_head = pd.read_excel(uploaded, header=1, usecols=[0], nrows=1)

    primeira_coluna = str(df_head.columns[0]).strip().upper()

    if primeira_coluna == "ORDEM":

        # leitura completa da planilha
        if uploaded.name.endswith(".csv"):
            df = pd.read_csv(uploaded, header=1)
        else:
            df = pd.read_excel(uploaded, header=1)

        import coleta
        coleta.run(df)
        st.stop()

    # ---------------------------------
    # FLUXO NORMAL (APP ATUAL)
    # ---------------------------------
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded, header=1)
    else:
        df = pd.read_excel(uploaded, header=1)

    # -----------------------------
    # COLUNAS FIXAS
    # -----------------------------
    COL_UNIDADE = df.columns[6]   # G
    COL_STATUS  = df.columns[14]  # O

    COL_DESCRICAO_STATUS = df.columns[18]  # coluna S

    df[COL_UNIDADE] = df[COL_UNIDADE].astype(str).str.strip().str.upper()
    df[COL_STATUS] = df[COL_STATUS].astype(str).str.strip().str.upper()
    df[COL_DESCRICAO_STATUS] = df[COL_DESCRICAO_STATUS].astype(str).str.strip().str.upper()
    df[COL_DESCRICAO_STATUS] = (df[COL_DESCRICAO_STATUS].astype(str).str.strip().str.upper())

    # --------------------------------------------------
    # SELETOR DE STATUS
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("📌 Filtro de status")

    status_disponiveis = sorted(df[COL_STATUS].dropna().unique())

    status_selecionado = st.selectbox(
        "Selecione o status para envio",
        status_disponiveis
    )

    df_filtrado = df[df[COL_STATUS] == status_selecionado]

    # --------------------------------------------------
    # REGRA ESPECIAL – CUSTODIA
    # --------------------------------------------------
    if "CUSTODIA" in status_selecionado:

        descricoes = (
            df_filtrado[COL_DESCRICAO_STATUS]
            .dropna()
            .unique()
            .tolist()
        )

        descricoes = sorted([d for d in descricoes if d and d != "NAN"])

        descricao_selecionada = st.selectbox(
            "Selecione a descrição da custódia",
            descricoes
        )

        df_filtrado = df_filtrado[
            df_filtrado[COL_DESCRICAO_STATUS] == descricao_selecionada
        ]


    if df_filtrado.empty:
        st.warning("Nenhum registro encontrado para o filtro selecionado.")
        st.stop()

    grupos = df_filtrado.groupby(COL_UNIDADE)

    # --------------------------------------------------
    # CONFIGURAÇÃO DO E-MAIL
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("✉️ Configuração do e-mail")

    cc_input = st.text_input(
        "CC (separados por vírgula)",
        placeholder="email1@evelog.com.br, email2@evelog.com.br"
    )

    assunto = st.text_input(
        "Assunto",
        placeholder="Assunto do e-mail"
    )

    texto_base = st.text_area(
        "Corpo do e-mail",
        placeholder="Digite a mensagem",
        height=150
    )

    # --------------------------------------------------
    # ENVIO
    # --------------------------------------------------
    if st.button("🚀 Enviar e-mails por unidade"):

        if not email_user or not senha:
            st.error("Informe o e-mail e a senha.")
            st.stop()

        if not assunto or not texto_base:
            st.error("Preencha o assunto e o corpo do e-mail.")
            st.stop()

        # CCs
        cc_list = []
        if cc_input:
            cc_list = [e.strip() for e in cc_input.split(",") if e.strip()]

        # CC fixo: remetente
        if email_user not in cc_list:
            cc_list.append(email_user)

        log_envio = []
        sem_email = []

        total_grupos = len(grupos)
        emails_enviados = 0

        progress_bar = st.progress(0)
        contador_placeholder = st.empty()

        with st.spinner("📨 Enviando e-mails..."):

            try:
                with smtplib.SMTP_SSL("email-ssl.com.br", 465) as smtp:
                    smtp.login(email_user, senha)

                    for unidade, pedidos_unidade in grupos:

                        emails_raw = emails_unidades.get(unidade)

                        if not emails_raw:
                            sem_email.append(unidade)
                            continue

                        emails_to = [e.strip() for e in emails_raw.split(",") if e.strip()]

                        texto_html = texto_base.replace("\n", "<br>")

                        # -----------------------------
                        # TABELA DO E-MAIL
                        # A,B,C,D,G,H,J,O,Q,R
                        # -----------------------------
                        if "CUSTODIA" in status_selecionado:
                            tabela = pedidos_unidade.iloc[
                                :, [0,1,2,3,6,7,9,14,16,17,18]
                            ]

                            tabela.columns = [
                                "Codigo",
                                "Nota Fiscal",
                                "Pedido",
                                "Cliente",
                                "Destino",
                                "Cidade",
                                "UF",
                                "Status",
                                "Dt Evento",
                                "Previsao",
                                "Descrição"
                            ]
                        else:
                            tabela = pedidos_unidade.iloc[
                                :, [0,1,2,3,6,7,9,14,16,17]
                            ]

                            tabela.columns = [
                                "Codigo",
                                "Nota Fiscal",
                                "Pedido",
                                "Cliente",
                                "Destino",
                                "Cidade",
                                "UF",
                                "Status",
                                "Dt Evento",
                                "Previsao"
                            ]

                        tabela_html = tabela.to_html(index=False, border=1)

                        corpo_html = f"""
                        <p>{texto_html}</p>
                        {tabela_html}
                        <p><strong><u>SE NÃO ESTIVER NA SUA UNIDADE, FAVOR DESCONSIDERAR.</u></strong></p>
                        <p><i>Mensagem automática.</i></p>
                        """

                        msg = MIMEText(corpo_html, "html")
                        msg["From"] = email_user
                        msg["To"] = ", ".join(emails_to)
                        msg["Subject"] = f"{assunto} – Unidade {unidade}"
                        msg["Cc"] = ", ".join(cc_list)

                        smtp.send_message(
                            msg,
                            to_addrs=emails_to + cc_list
                        )

                        emails_enviados += 1

                        percentual = int((emails_enviados / total_grupos) * 100)
                        progress_bar.progress(percentual)

                        contador_placeholder.markdown(
                            f"""
                            **📧 E-mails enviados:** {emails_enviados}  
                            **🏢 Unidades acionadas:** {emails_enviados} / {total_grupos}
                            """
                        )

                        log_envio.append({
                            "Unidade": unidade,
                            "Status": status_selecionado,
                            "Qtd registros": len(pedidos_unidade),
                            "Para": ", ".join(emails_to),
                            "CC": ", ".join(cc_list)
                        })

                    progress_bar.progress(100)
                    st.success("✅ Envio concluído com sucesso!")


                st.success(f"✅ {len(log_envio)} e-mails enviados com sucesso!")

                if log_envio:
                    st.subheader("📄 Log de envio")
                    st.dataframe(pd.DataFrame(log_envio))

                if sem_email:
                    st.warning("⚠️ Unidades sem e-mail cadastrado:")
                    st.write(sem_email)

            except Exception as e:
                st.error(f"Erro de conexão SMTP: {e}")
