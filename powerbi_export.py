import sys
import json
import time
import msal
import requests
import pandas as pd
import smtplib
import os
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
from sqlalchemy import create_engine

load_dotenv()

# ── Configurações ───────────────────────────────────────────
CLIENT_ID  = os.environ["CLIENT_ID"]
TENANT_ID  = os.environ["TENANT_ID"]
DATASET_ID = os.environ["DATASET_ID"]
SAVE_PATH  = os.environ["SAVE_PATH"]
LOG_PATH   = os.environ["LOG_PATH"]
TOKEN_PATH = os.environ["TOKEN_PATH"]

DB_URL = (
    f"mysql+pymysql://{os.environ['MARIADB_USER']}:{os.environ['MARIADB_PASSWORD']}"
    f"@{os.environ['MARIADB_HOST']}:{os.environ['MARIADB_PORT']}/{os.environ['MARIADB_DATABASE']}"
    f"?charset=utf8mb4"
)

EMAIL_REMETENTE = os.environ["EMAIL_REMETENTE"]
EMAIL_SENHA     = os.environ["EMAIL_SENHA"]
EMAIL_DESTINO   = os.environ["EMAIL_DESTINO"]
SMTP_HOST       = os.environ["SMTP_HOST"]
SMTP_PORT       = int(os.environ["SMTP_PORT"])

SCOPE = [
    "https://analysis.windows.net/powerbi/api/Workspace.Read.All",
    "https://analysis.windows.net/powerbi/api/Dataset.Read.All",
    "https://analysis.windows.net/powerbi/api/Report.Read.All",
]

DAX_QUERY = """
EVALUATE
SUMMARIZECOLUMNS(
    'Base'[status vencimento],
    'D_Empresas'[EMPRESA],
    TREATAS({"Vencido"}, 'Base'[status de atraso]),
    FILTER(
        KEEPFILTERS(VALUES('Base'[VENCIMENTO])),
        AND('Base'[VENCIMENTO] >= DATE(1900, 1, 1), 'Base'[VENCIMENTO] < TODAY())
    ),
    "saldo", 'medidas'[04.saldo]
)
ORDER BY
    'Base'[status vencimento],
    'D_Empresas'[EMPRESA]
"""

# ── Log ─────────────────────────────────────────────────────
def log(msg):
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{agora}] {msg}"
    print(linha)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(linha + "\n")

# ── E-mail ──────────────────────────────────────────────────
def enviar_email(assunto, corpo, anexo=None):
    try:
        msg = MIMEMultipart()
        msg["From"]    = EMAIL_REMETENTE
        msg["To"]      = EMAIL_DESTINO
        msg["Subject"] = assunto
        msg.attach(MIMEText(corpo, "plain", "utf-8"))

        if anexo and os.path.exists(anexo):
            with open(anexo, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(anexo)}")
                msg.attach(part)

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINO, msg.as_string())

        log("📧 E-mail enviado com sucesso!")
    except Exception as e:
        log(f"❌ Erro ao enviar e-mail: {e}")

# ── Alerta de expiração do token ────────────────────────────
def alertar_expiracao_token():
    if not os.path.exists(TOKEN_PATH):
        return
    try:
        with open(TOKEN_PATH, "r") as f:
            cache = json.load(f)

        refresh_tokens = cache.get("RefreshToken", {})
        if not refresh_tokens:
            return

        rt = list(refresh_tokens.values())[0]
        last_used = int(rt.get("last_used_on", 0))
        if last_used == 0:
            return

        dias_desde_uso = (time.time() - last_used) / 86400
        dias_restantes = int(90 - dias_desde_uso)

        if dias_restantes <= 15:
            log(f"⚠️ Token expira em ~{dias_restantes} dias!")
            enviar_email(
                assunto=f"⚠️ Token Power BI expira em {dias_restantes} dias",
                corpo=f"O token de autenticação do Power BI vai expirar em aproximadamente {dias_restantes} dias.\n\n"
                      f"Acesse a VM e execute o script manualmente para renovar o token antes que expire.\n\n"
                      f"Se expirar sem renovação, a exportação de terça-feira vai falhar."
            )
    except Exception as e:
        log(f"⚠️ Não foi possível verificar expiração do token: {e}")

# ── Autenticação com cache de token ─────────────────────────
def autenticar():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    cache = msal.SerializableTokenCache()

    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "r") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=authority, token_cache=cache)

    # Tenta usar token salvo
    contas = app.get_accounts()
    token = None
    if contas:
        token = app.acquire_token_silent(scopes=SCOPE, account=contas[0])

    # Se não tiver token salvo, faz login interativo
    if not token or "access_token" not in token:
        log("🔐 Token expirado ou ausente. Abrindo navegador para login...")
        token = app.acquire_token_interactive(scopes=SCOPE)

    # Salva o cache atualizado
    with open(TOKEN_PATH, "w") as f:
        f.write(cache.serialize())

    if "access_token" not in token:
        raise Exception("Falha na autenticação: " + token.get("error_description", ""))

    log("✅ Autenticado com sucesso!")
    alertar_expiracao_token()
    return {"Authorization": f"Bearer {token['access_token']}"}

# ── Exportar dados ──────────────────────────────────────────
def exportar_dados():
    log("=" * 50)
    log("⏳ Iniciando exportação...")

    try:
        headers = autenticar()

        payload = {
            "queries": [{"query": DAX_QUERY}],
            "serializerSettings": {"includeNulls": True}
        }

        r = requests.post(
            f"https://api.powerbi.com/v1.0/myorg/datasets/{DATASET_ID}/executeQueries",
            headers=headers,
            json=payload
        )

        if r.status_code != 200:
            raise Exception(f"Erro na API: {r.status_code} - {r.text}")

        rows = r.json()["results"][0]["tables"][0].get("rows", [])

        if not rows:
            raise Exception("Nenhum dado retornado pela query.")

        df = pd.DataFrame(rows)
        df.columns = ["status_vencimento", "empresa", "saldo"]
        df["saldo"] = df["saldo"].round(2)
        df["_ordem"] = df["status_vencimento"].str.extract(r"(\d+)").astype(int)
        df = df.sort_values(["_ordem", "empresa"]).drop(columns="_ordem").reset_index(drop=True)
        df["data_hora"]      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df["vencimento_ate"] = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

        engine = create_engine(DB_URL)
        with engine.begin() as conn:
            df.to_sql("contas_a_receber", con=conn, if_exists="append", index=False)

        log(f"✅ {len(df)} linhas salvas no banco (tabela: contas_a_receber)")

        enviar_email(
            assunto="✅ Exportação Power BI concluída",
            corpo=f"A exportação de Contas a Receber foi concluída com sucesso!\n\n"
                  f"📊 Linhas salvas: {len(df)}\n"
                  f"🗄️ Banco: {os.environ['MARIADB_DATABASE']} › contas_a_receber\n"
                  f"🕐 Horário: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        )

    except Exception as e:
        log(f"❌ Erro na exportação: {e}")
        enviar_email(
            assunto="❌ Erro na exportação Power BI",
            corpo=f"Ocorreu um erro na exportação de Contas a Receber.\n\n"
                  f"Erro: {e}\n"
                  f"🕐 Horário: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
                  f"Verifique o log em: {LOG_PATH}"
        )

os.makedirs(SAVE_PATH, exist_ok=True)
exportar_dados()