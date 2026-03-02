import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText


# ==================================================
# BASE DE E-MAILS DAS UNIDADES
# ==================================================
@st.cache_data
def carregar_emails_unidades():
    df = pd.read_excel("emails_unidades.xlsx", header=0)
    df.columns = ["Unidade", "Emails"]

    df["Unidade"] = df["Unidade"].astype(str).str.strip().str.upper()
    df["Emails"] = df["Emails"].astype(str).str.strip()

    return dict(zip(df["Unidade"], df["Emails"]))


emails_unidades = carregar_emails_unidades()


# ==================================================
# FUNÇÃO PRINCIPAL
# ==================================================
def run(arquivo, email_user, senha):

    # --------------------------------------------------
    # LEITURA DO ARQUIVO
    # --------------------------------------------------
    if arquivo.name.endswith(".csv"):
        df = pd.read_csv(arquivo)
    else:
        df = pd.read_excel(arquivo)

    df.columns = [
        "RE",
        "SIGLA",
        "TIPO",
        "CTE",
        "VINCULAR_ACERTO",
        "ORDEM",
        "SITUACAO",
        "DT_FINALIZACAO",
        "DIAS_FALTANTES",
        "SITUACAO_COLETA",
        "UNIDADE",
        "EMAIL"
    ]

    df["UNIDADE"] = df["UNIDADE"].astype(str).str.strip().str.upper()

    # --------------------------------------------------
    # CONFIGURAÇÃO DE CC
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("✉️ Configuração do envio")

    cc_input = st.text_input(
        "CC (separados por vírgula)",
        placeholder="email1@evelog.com.br, email2@evelog.com.br"
    )

    # --------------------------------------------------
    # ENVIO
    # --------------------------------------------------
    if st.button("🚀 Enviar e-mails"):

        cc_list = []
        if cc_input:
            cc_list = [e.strip() for e in cc_input.split(",") if e.strip()]

        # Remetente fixo em CC
        if email_user not in cc_list:
            cc_list.append(email_user)

        total = len(df)
        enviados = 0

        log_envio = []
        sem_email = []

        progress_bar = st.progress(0)
        contador = st.empty()

        with st.spinner("📨 Enviando e-mails..."):

            try:
                with smtplib.SMTP_SSL("email-ssl.com.br", 465) as smtp:
                    smtp.login(email_user, senha)

                    for _, linha in df.iterrows():

                        unidade = str(linha["UNIDADE"]).strip().upper()
                        ordem = linha["ORDEM"]
                        sigla = linha["SIGLA"]

                        emails_raw = emails_unidades.get(unidade)

                        if not emails_raw:
                            sem_email.append(unidade)
                            continue

                        emails_to = [
                            e.strip()
                            for e in emails_raw.split(",")
                            if e.strip()
                        ]

                        # ASSUNTO DINÂMICO
                        assunto = (
                            f"PRÉ-ALERTA - COLETA MALOTE CLIENTE MCDONALD'S "
                            f"OC - {ordem} {sigla}"
                        )

                        # CORPO HTML FORMATADO
                        corpo_html = f"""
                        <div style="font-family: Arial, sans-serif; font-size: 14px;">

                        <p style="color:red; font-weight:bold; font-size:16px;">
                        URGENTE!
                        </p>

                        <p style="background-color:#2ecc71; color:white; font-weight:bold; font-size:18px; padding:4px;">
                        COLETA DE MALOTE – DOCUMENTOS
                        </p>

                        <p>Prezados, boa tarde!</p>

                        <p style="background-color:#17c9c3; color:white; font-weight:bold; padding:4px;">
                        Por gentileza, providenciar coleta com urgência.
                        Coleta alinhada com o restaurante, o mesmo está no aguardo!!!
                        </p>

                        <p style="background-color:#d633ff; color:white; font-weight:bold; padding:4px;">
                        C/C EMISSÃO 0153080 - MALOTES
                        </p>

                        <p style="background-color:#f1c40f; font-weight:bold; padding:3px;">
                        Essa coleta deve ser feita no mesmo dia (dependendo do horário),
                        ou no dia seguinte.
                        </p>

                        <p>Não realizar a coleta em finais de semanas;</p>

                        <ul>
                        <li>
                        Emita pela tarja e nos informe o nº do CTE para que possamos
                        vincular a ordem e creditar o valor da coleta de
                        <span style="background-color:#2ecc71; font-weight:bold;">
                        R$13,20
                        </span>.
                        </li>

                        <li>Mencione o lacre no campo pedido.</li>
                        <li>Esse item é de suma importância</li>
                        </ul>

                        <p style="color:red; font-weight:bold;">
                        ATENÇÃO!
                        </p>

                        <p style="background-color:#f1c40f; font-weight:bold; padding:4px;">
                        CASO O RESTAURANTE NÃO ENVIE O MALOTE,
                        PEGUE A RESSALVA NA ORDEM (Nome legível, data e hora)
                        e nos encaminhe via e-mail para que possamos gerar a improdutiva.
                        </p>

                        <p>
                        Caso tenha alguma ordem de coleta pendente de acerto,
                        favor encaminhar em resposta a este e-mail
                        com CTE reversa / OC para que seja feito o acerto.
                        </p>

                        <p>
                        Obrigado, qualquer dúvida estou à disposição. 😊
                        </p>

                        </div>
                        """

                        msg = MIMEText(corpo_html, "html")
                        msg["From"] = email_user
                        msg["To"] = ", ".join(emails_to)
                        msg["Cc"] = ", ".join(cc_list)
                        msg["Subject"] = assunto

                        smtp.send_message(
                            msg,
                            to_addrs=emails_to + cc_list
                        )

                        enviados += 1
                        percentual = int((enviados / total) * 100)
                        progress_bar.progress(percentual)

                        contador.markdown(
                            f"📧 E-mails enviados: {enviados} / {total}"
                        )

                        log_envio.append({
                            "Unidade": unidade,
                            "Ordem": ordem,
                            "Para": ", ".join(emails_to)
                        })

                progress_bar.progress(100)
                st.success("✅ Envio concluído com sucesso!")

                if log_envio:
                    st.subheader("📄 Log de envio")
                    st.dataframe(pd.DataFrame(log_envio), hide_index=True)

                if sem_email:
                    st.warning("⚠️ Unidades sem e-mail cadastrado:")
                    st.write(list(set(sem_email)))

            except Exception as e:
                st.error(f"Erro SMTP: {e}")