import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# --------------------------------------------------
# CONFIG STREAMLIT
# --------------------------------------------------

st.set_page_config(page_title="AutoMailer", layout="wide")
st.title("📮Envio Automático de emails")

# --------------------------------------------------
# BASE DE E-MAILS
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
# FORMULÁRIO (1/3 DA TELA)
# --------------------------------------------------
col_form, _ = st.columns([1, 2])

with col_form:
    email_user = st.text_input(
        "E-mail",
        placeholder="atendimento@evelog.com.br"
    )

    senha = st.text_input(
        "Senha",
        type="password"
    )

    uploaded = st.file_uploader(
        "Importar planilha",
        type=["xlsx", "xls", "csv"]
    )

# --------------------------------------------------
# PROCESSAMENTO
# --------------------------------------------------

if uploaded:

    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded, header=1)
    else:
        df = pd.read_excel(uploaded, header=1)

    # Colunas fixas
    COL_UNIDADE = df.columns[6]   # G
    COL_STATUS = df.columns[14]  # O

    df[COL_UNIDADE] = df[COL_UNIDADE].astype(str).str.strip().str.upper()
    df[COL_STATUS] = df[COL_STATUS].astype(str).str.strip().str.upper()

    entrada = df[df[COL_STATUS] == "ENTRADA"]

    if entrada.empty:
        st.warning("Nenhum pedido com status ENTRADA encontrado.")
        st.stop()

    grupos = entrada.groupby(COL_UNIDADE)

    st.markdown("---")
    st.subheader("✉️ Configuração do e-mail")

    cc_input = st.text_input(
    "CC (Separados por vírgula)",
    placeholder="atendimento1@evelog.com.br,atendimento2@evelog.com.br"
    )

    assunto = st.text_input(
        "Assunto",
        placeholder="Digite o assunto"
    )

    texto_base = st.text_area(
        "Corpo do e-mail",
        placeholder=("Digite a mensagem"),
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

        cc_list = []

        # CC digitado pelo usuário
        if cc_input:
            cc_list = [e.strip() for e in cc_input.split(",") if e.strip()]

        # CC fixo: remetente
        if email_user not in cc_list:
            cc_list.append(email_user)

        log_envio = []
        sem_email = []
        erros = []

        try:
            with smtplib.SMTP_SSL("email-ssl.com.br", 465) as smtp:
                smtp.login(email_user, senha)

                for unidade, pedidos_unidade in grupos:

                    emails_raw = emails_unidades.get(unidade)

                    if not emails_raw:
                        sem_email.append(unidade)
                        continue

                    emails_to = [e.strip() for e in emails_raw.split(",") if e.strip()]


                    # Texto HTML
                    texto_html = texto_base.replace("\n", "<br>")

                    # Tabela HTML (A,B,C,D,G,H,J,O,Q,R)
                    colunas_email = pedidos_unidade.iloc[
                        :, [0, 1, 2, 3, 6, 7, 9, 14, 16, 17]
                    ].copy()

                    colunas_email.columns = [
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

                    tabela_html = colunas_email.to_html(
                        index=False,
                        border=1
                    )

                    corpo_html = f"""
                    <p>{texto_html}</p>
                    {tabela_html}
                    <p><i>Mensagem automática.</i></p>
                    """

                    msg = MIMEText(corpo_html, "html")
                    msg["From"] = email_user
                    msg["To"] = ", ".join(emails_to)
                    msg["Subject"] = f"{assunto} – Unidade {unidade}"

                    if cc_list:
                        msg["Cc"] = ", ".join(cc_list)

                    destinatarios = emails_to + cc_list

                    smtp.send_message(msg, to_addrs=destinatarios)

                    log_envio.append({
                        "Unidade": unidade,
                        "Para": ", ".join(emails_to),
                        "CC": ", ".join(cc_list) if cc_list else "-",
                        "Qtd pedidos": len(pedidos_unidade)
                    })

            st.success(f"✅ {len(log_envio)} e-mails enviados com sucesso!")

            if log_envio:
                st.subheader("📄 Detalhes dos envios")
                st.dataframe(pd.DataFrame(log_envio))

            if sem_email:
                st.warning("⚠️ Unidades sem e-mail cadastrado:")
                st.write(sem_email)

            if erros:
                st.error("❌ Erros no envio:")
                st.dataframe(pd.DataFrame(erros))

        except Exception as e:
            st.error(f"Erro no envio SMTP: {e}")
