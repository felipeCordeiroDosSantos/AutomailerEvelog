import streamlit as st
import pandas as pd
import re
import smtplib
from email.mime.text import MIMEText


# --------------------------------------------------
# BASE DE EMAILS DOS RESTAURANTES
# --------------------------------------------------
@st.cache_data
def carregar_emails_restaurantes():
    df = pd.read_excel("emails_restaurantes.xlsx", header=0)
    df.columns = ["RESTAURANTE", "EMAILS"]

    df["RESTAURANTE"] = (
        df["RESTAURANTE"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["EMAILS"] = df["EMAILS"].astype(str).str.strip()

    return dict(zip(df["RESTAURANTE"], df["EMAILS"]))


emails_restaurantes = carregar_emails_restaurantes()


# --------------------------------------------------
# PARSER DOS TXT
# --------------------------------------------------
def parse_txt(arquivo):
    linhas = arquivo.read().decode("latin-1").splitlines()
    dados = []

    for linha in linhas[1:]:  # pula cabeçalho
        if not linha.strip():
            continue

        partes = re.split(r"\s{2,}", linha.strip())

        try:
            restaurante = partes[0]
            pedido = partes[1]
            data = partes[2]
            item = partes[3]
            qtde = partes[4]
            descricao = partes[5]
            preco_rs = partes[6]

            idx = 7
            preco_usd = None
            if idx < len(partes) and re.match(r"^[\d\.,]+$", partes[idx]):
                preco_usd = partes[idx]
                idx += 1

            responsavel = partes[idx]
            idx += 1

            observacao = (
                " ".join(partes[idx:-2])
                if len(partes) > idx + 2
                else None
            )

            oc = partes[-2]
            cnpj = partes[-1]

            dados.append([
                restaurante, pedido, data, item, qtde,
                descricao, preco_rs, preco_usd,
                responsavel, observacao, oc, cnpj,
                arquivo.name
            ])

        except Exception:
            continue

    return pd.DataFrame(dados, columns=[
        "RESTAURANTE",
        "PEDIDO",
        "DATA",
        "ITEM",
        "QTDE",
        "DESCRICAO",
        "PRECO_UNIT_RS",
        "PRECO_UNIT_USD",
        "RESPONSAVEL",
        "OBSERVACAO",
        "OC",
        "CNPJ",
        "ARQUIVO_ORIGEM"
    ])


# --------------------------------------------------
# FLUXO PRINCIPAL
# --------------------------------------------------
def run(arquivos, email_user, senha):

    # -----------------------------
    # VARIÁVEIS DE CONTROLE
    # -----------------------------
    log_envio = []
    sem_email = []

    # -----------------------------
    # UNIFICA TODOS OS TXT
    # -----------------------------
    dfs = [parse_txt(arq) for arq in arquivos]
    df = pd.concat(dfs, ignore_index=True)

    # -----------------------------
    # TRATAMENTOS
    # -----------------------------
    df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
    df["QTDE"] = pd.to_numeric(df["QTDE"], errors="coerce")

    df["PRECO_UNIT_RS"] = (
        df["PRECO_UNIT_RS"]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df["PRECO_UNIT_RS"] = pd.to_numeric(
        df["PRECO_UNIT_RS"], errors="coerce"
    )

    # -----------------------------
    # CONFIGURAÇÃO DO EMAIL
    # -----------------------------
    st.markdown("---")
    st.subheader("✉️ Configuração do e-mail")

    cc_input = st.text_input(
        "CC (separados por vírgula)",
        placeholder="email1@evelog.com.br, email2@evelog.com.br"
    )

    # -----------------------------
    # ENVIO DOS EMAILS
    # -----------------------------
    if st.button("🚀 Enviar e-mails por pedido"):

        if not email_user or not senha:
            st.error("Credenciais não informadas no app principal.")
            st.stop()

        total = len(df)
        enviados = 0

        progress_bar = st.progress(0)
        contador = st.empty()

        with st.spinner("📨 Enviando e-mails..."):

            with smtplib.SMTP_SSL("email-ssl.com.br", 465) as smtp:
                smtp.login(email_user, senha)

                for _, pedido in df.iterrows():

                    restaurante = pedido["RESTAURANTE"]
                    emails_raw = emails_restaurantes.get(restaurante)

                    if not emails_raw:
                        sem_email.append(restaurante)
                        continue

                    emails_to = [
                        e.strip()
                        for e in emails_raw.split(",")
                        if e.strip()
                    ]

                    # CC fixo = remetente
                    cc_list = [email_user]

                    corpo_html = f"""
                    <p>Bom dia!</p>
                    <br>
                    <p><strong>{restaurante}</strong>,</p>
                    <p>
                    Foi transmitido a nós o pedido: 
                    <strong>{pedido['PEDIDO']}</strong> referentes a 
                    <strong>{pedido['DESCRICAO']}</strong>, 
                    solicitado via Central de Pedidos por 
                    <strong>{pedido['RESPONSAVEL']}</strong>.
                    </p>
                    <p>
                    Por gentileza, nos encaminhar a NOTA FISCAL 
                    para agendamento da coleta.
                    </p>
                    <br>
                    <p>Obrigado, no aguardo de um retorno.</p>
                    """

                    msg = MIMEText(corpo_html, "html")
                    msg["From"] = email_user
                    msg["To"] = ", ".join(emails_to)
                    msg["Cc"] = ", ".join(cc_list)
                    msg["Subject"] = f'SOLICITAÇÃO DE NF NIG "{pedido["PEDIDO"]}"'

                    smtp.send_message(
                        msg,
                        to_addrs=emails_to + cc_list
                    )

                    enviados += 1
                    progress_bar.progress(int((enviados / total) * 100))

                    contador.markdown(
                        f"📧 E-mails enviados: {enviados} / {total}"
                    )

                    log_envio.append({
                        "Restaurante": restaurante,
                        "Pedido": pedido["PEDIDO"],
                        "Para": ", ".join(emails_to)
                    })

        if log_envio:
            st.success("✅ Envio concluído com sucesso!")
            st.subheader("📄 Log de envio")
            st.dataframe(pd.DataFrame(log_envio), hide_index=True)

        if sem_email:
            st.warning("⚠️ Restaurantes sem e-mail cadastrado:")
            st.write(list(set(sem_email)))

