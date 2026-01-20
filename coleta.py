import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# --------------------------------------------------
# BASE DE E-MAILS DAS UNIDADES
# --------------------------------------------------
@st.cache_data
def carregar_emails_unidades():
    df = pd.read_excel("emails_unidades.xlsx", header=0)
    df.columns = ["UNIDADE", "EMAILS"]

    df["UNIDADE"] = df["UNIDADE"].astype(str).str.strip().str.upper()
    df["EMAILS"] = df["EMAILS"].astype(str).str.strip()

    return dict(zip(df["UNIDADE"], df["EMAILS"]))


def run(df):

    st.set_page_config(
        page_title="Coleta de Pedidos",
        layout="wide"
    )

    # --------------------------------------------------
    # NORMALIZA COLUNAS
    # --------------------------------------------------
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    COL_ORDEM = "ORDEM"
    COL_ORIGEM = "ORIGEM"

    if COL_ORDEM not in df.columns or COL_ORIGEM not in df.columns:
        st.error("A planilha não contém as colunas obrigatórias (ORDEM, ORIGEM).")
        st.stop()

    df[COL_ORDEM] = df[COL_ORDEM].astype(str).str.strip()
    df[COL_ORIGEM] = df[COL_ORIGEM].astype(str).str.strip().str.upper()

    # --------------------------------------------------
    # UPLOAD DOS PDFs
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("📎 Upload dos PDFs")

    col_pdfs, _ = st.columns([1, 2])

    with col_pdfs:
        pdfs = st.file_uploader(
            "Importar PDFs",
            type=["pdf"],
            accept_multiple_files=True
        )

        if not pdfs:
            st.info("Aguardando upload dos PDFs.")
            st.stop()

    pdf_map = {
        pdf.name.replace(".pdf", "").strip(): pdf
        for pdf in pdfs
    }

    df["TEM_PDF"] = df[COL_ORDEM].isin(pdf_map.keys())

    # --------------------------------------------------
    # FILTRA APENAS PEDIDOS COM PDF
    # --------------------------------------------------
    df_envio = df[df["TEM_PDF"]].copy()

    if df_envio.empty:
        st.warning("Nenhum pedido com PDF encontrado.")
        st.stop()

    # --------------------------------------------------
    # AGRUPAMENTO POR UNIDADE (ORIGEM)
    # --------------------------------------------------
    grupos = df_envio.groupby(COL_ORIGEM)

    st.markdown("---")
    st.subheader("📊 Resumo por unidade")

    resumo = grupos.size().reset_index(name="Qtd pedidos")
    st.dataframe(resumo)

    # --------------------------------------------------
    # CONFIGURAÇÃO DO E-MAIL
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("✉️ Configuração do e-mail")

    cc_input = st.text_input(
        "CC (separados por vírgula)",
        placeholder="email1@evelog.com.br, email2@evelog.com.br",
        key="coleta_cc"
    )

    texto_base = st.text_area(
        "Corpo do e-mail",
        placeholder="Digite a mensagem",
        height=150,
        key="coleta_texto"
    )

    # --------------------------------------------------
    # ENVIO
    # --------------------------------------------------
    if st.button("🚀 Enviar e-mails de coleta", key="coleta_enviar"):

        email_user = st.session_state.get("email_user")
        senha = st.session_state.get("email_smtp")

        if not email_user or not senha:
            st.error("Informe o e-mail remetente e a senha no app principal.")
            st.stop()

        if not texto_base or not texto_base.strip():
            st.error("Preencha o corpo do e-mail.")
            st.stop()

        # CCs
        cc_list = []
        if cc_input:
            cc_list = [e.strip() for e in cc_input.split(",") if e.strip()]

        # CC fixo: remetente
        if email_user not in cc_list:
            cc_list.append(email_user)

        emails_unidades = carregar_emails_unidades()

        enviados = 0
        sem_email = []

        progress = st.progress(0)
        total = len(grupos)

        log_envio = []
        sem_email = []

        emails_enviados = 0
        total_unidades = len(grupos)

        progress_bar = st.progress(0)
        contador_placeholder = st.empty()

        with st.spinner("📨 Enviando e-mails de coleta..."):
            try:
                with smtplib.SMTP_SSL("email-ssl.com.br", 465) as smtp:
                    smtp.login(email_user, senha)

                    for i, (unidade, pedidos_unidade) in enumerate(grupos, start=1):

                        emails_raw = emails_unidades.get(unidade)

                        if not emails_raw:
                            sem_email.append(unidade)
                            continue

                        emails_to = [
                            e.strip() for e in emails_raw.split(",") if e.strip()
                        ]

                        ordens = pedidos_unidade["ORDEM"].tolist()
                        ordens_txt = ", ".join(ordens)

                        assunto = (
                            "PRÉ ALERTA DE COLETA TRAMONTINA - "
                            f"{ordens_txt}"
                        )

                        texto_html = texto_base.replace("\n", "<br>")

                        tabela_email = pedidos_unidade.drop(columns=["TEM_PDF"], errors="ignore")

                        tabela_html = tabela_email.to_html(
                            index=False,
                            border=1
                        )

                        corpo_html = f"""
                        <p>{texto_html}</p>
                        {tabela_html}
                        <p><i>Mensagem automática.</i></p>
                        """

                        msg = MIMEMultipart()
                        msg["From"] = email_user
                        msg["To"] = ", ".join(emails_to)
                        msg["Subject"] = assunto
                        msg["Cc"] = ", ".join(cc_list)

                        msg.attach(MIMEText(corpo_html, "html"))

                        # ANEXA PDFs DA UNIDADE
                        for ordem in ordens:
                            pdf = pdf_map.get(ordem)
                            if pdf:
                                anexo = MIMEApplication(pdf.read(), _subtype="pdf")
                                anexo.add_header(
                                    "Content-Disposition",
                                    "attachment",
                                    filename=pdf.name
                                )
                                msg.attach(anexo)

                        smtp.send_message(
                            msg,
                            to_addrs=emails_to + cc_list
                        )

                        log_envio.append({
                            "Unidade": unidade,
                            "Qtd registros": len(pedidos_unidade),
                            "Para": ", ".join(emails_to),
                            "CC": ", ".join(cc_list)
                        })

                        emails_enviados += 1

                        percentual = int((emails_enviados / total_unidades) * 100)
                        progress_bar.progress(percentual)

                        contador_placeholder.markdown(
                            f"""
                            **📧 E-mails enviados:** {emails_enviados}  
                            **🏢 Unidades acionadas:** {emails_enviados} / {total_unidades}
                            """
                        )

                        enviados += 1
                        progress.progress(int((i / total) * 100))

                st.success(f"✅ {enviados} e-mails enviados com sucesso!")

                if log_envio:
                    st.subheader("📄 Log de envio")
                    st.dataframe(pd.DataFrame(log_envio))

                if sem_email:
                    st.warning("⚠️ Unidades sem e-mail cadastrado:")
                    st.write(sem_email)

            except Exception as e:
                st.error(f"Erro no envio: {e}")
