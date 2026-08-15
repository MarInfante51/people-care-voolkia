
import io
import re
import smtplib
import unicodedata
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import requests
import streamlit as st
from rapidfuzz import fuzz

st.set_page_config(
    page_title="People Care | Voolkia",
    page_icon="🟠",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# -----------------------------
# ESTILO VOOLKIA
# -----------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #F7F7F7; color: #240300; }

.block-container {
    max-width: 920px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

.v-header {
    background: linear-gradient(135deg, #FF7000 0%, #FF9100 100%);
    padding: 28px 30px;
    border-radius: 28px;
    color: white;
    margin-bottom: 20px;
}
.v-header h1 { margin: 0; font-size: 2rem; font-weight: 800; }
.v-header p { margin: 8px 0 0 0; font-size: 1rem; }

.v-card {
    background: white;
    border: 1px solid #EDEDED;
    border-radius: 24px;
    padding: 22px;
    margin: 12px 0;
}

.v-user {
    display:inline-block;
    background:#EDEDED;
    color:#240300;
    padding:8px 14px;
    border-radius:999px;
    font-size:.88rem;
    margin-bottom:12px;
}

div.stButton > button {
    width: 100%;
    border-radius: 999px;
    border: 0;
    background: #FF7000;
    color: white;
    font-weight: 700;
    min-height: 46px;
}
div.stButton > button:hover { background: #FF9100; color: #240300; }

div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div,
textarea {
    border-radius: 14px !important;
}

[data-testid="stChatMessage"] {
    background: white;
    border-radius: 20px;
    border: 1px solid #EDEDED;
    padding: 10px 14px;
}

.small-note { color:#6c6461; font-size:.82rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="v-header">
  <h1>People Care</h1>
  <p>Tu espacio de consultas frecuentes de Voolkia.</p>
</div>
""", unsafe_allow_html=True)

# -----------------------------
# CONFIG
# -----------------------------
ALERT_EMAIL = st.secrets.get("ALERT_EMAIL", "marcela@voolkia.com")
FAQ_DRIVE_FILE_ID = st.secrets.get("FAQ_DRIVE_FILE_ID", "")
FAQ_LOCAL_FILE = "People_Care_Voolkia_Base_Conocimiento.xlsx"

MATCH_THRESHOLD = 62

def normalize(text: str) -> str:
    text = str(text or "").lower().strip()
    text = ''.join(c for c in unicodedata.normalize("NFD", text)
                   if unicodedata.category(c) != "Mn")
    text = re.sub(r"[^a-z0-9ñáéíóúü\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()

@st.cache_data(ttl=300)
def load_excel():
    if FAQ_DRIVE_FILE_ID:
        url = f"https://docs.google.com/spreadsheets/d/{FAQ_DRIVE_FILE_ID}/export?format=xlsx"
        r = requests.get(url, timeout=20)
        r.raise_for_status()
        content = io.BytesIO(r.content)
        faqs = pd.read_excel(content, sheet_name="FAQs")
        content.seek(0)
        alertas = pd.read_excel(content, sheet_name="Reglas_Alertas")
        content.seek(0)
        proyectos = pd.read_excel(content, sheet_name="Proyectos")
    else:
        faqs = pd.read_excel(FAQ_LOCAL_FILE, sheet_name="FAQs")
        alertas = pd.read_excel(FAQ_LOCAL_FILE, sheet_name="Reglas_Alertas")
        proyectos = pd.read_excel(FAQ_LOCAL_FILE, sheet_name="Proyectos")
    return faqs, alertas, proyectos

def find_faq(question, faqs):
    q = normalize(question)
    best = None
    best_score = 0

    active = faqs[faqs["Activo"].astype(str).str.lower().isin(["sí","si","true","1"])]
    for _, row in active.iterrows():
        searchable = f"{row['Pregunta']} {row.get('Palabras_clave','')} {row.get('Categoría','')}"
        s = normalize(searchable)
        score = max(
            fuzz.token_set_ratio(q, s),
            fuzz.partial_ratio(q, normalize(row["Pregunta"])),
        )
        if score > best_score:
            best_score = score
            best = row

    return best, best_score

def detect_alert(question, alertas):
    q = normalize(question)
    for _, row in alertas.iterrows():
        triggers = str(row.get("Disparadores",""))
        if "MISMA_PREGUNTA_3_VECES" in triggers or "SIN_COINCIDENCIA_FAQ" in triggers:
            continue
        for trigger in triggers.split(";"):
            t = normalize(trigger)
            if t and t in q:
                return {
                    "tipo": str(row["Tipo"]),
                    "prioridad": str(row["Prioridad"]),
                    "accion": str(row["Acción"]),
                    "trigger": trigger.strip(),
                }
    return None

def send_alert(subject, body):
    host = st.secrets.get("SMTP_HOST", "")
    port = int(st.secrets.get("SMTP_PORT", 587))
    user = st.secrets.get("SMTP_USER", "")
    password = st.secrets.get("SMTP_PASSWORD", "")

    if not all([host, user, password]):
        return False, "SMTP no configurado"

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ALERT_EMAIL
    msg.set_content(body)

    with smtplib.SMTP(host, port, timeout=20) as server:
        server.starttls()
        server.login(user, password)
        server.send_message(msg)
    return True, "enviado"

def alert_body(question, answer, status, alert_type):
    return f"""Alerta automática - People Care Voolkia

Fecha/hora: {datetime.now().strftime("%d/%m/%Y %H:%M")}
Colaborador: {st.session_state.nombre} {st.session_state.apellido}
Proyecto/Cliente: {st.session_state.proyecto}
Tipo: {alert_type}
Estado: {status}

Pregunta:
{question}

Respuesta del bot:
{answer}

Esta consulta fue registrada para revisión de People Care.
"""

try:
    faqs, alertas, proyectos_df = load_excel()
except Exception as e:
    st.error("No pude cargar la base de conocimiento. Revisá el archivo Excel o la configuración de Google Drive.")
    st.caption(str(e))
    st.stop()

project_list = proyectos_df[
    proyectos_df["Activo"].astype(str).str.lower().isin(["sí","si","true","1"])
]["Proyecto_Cliente"].dropna().astype(str).tolist()

# -----------------------------
# IDENTIFICACIÓN
# -----------------------------
if "identified" not in st.session_state:
    st.session_state.identified = False

if not st.session_state.identified:
    st.markdown('<div class="v-card">', unsafe_allow_html=True)
    st.subheader("Antes de comenzar")
    st.write("Identificate para que People Care pueda contextualizar tu consulta si necesitás atención personalizada.")

    with st.form("identificacion"):
        nombre = st.text_input("Nombre *")
        apellido = st.text_input("Apellido *")
        proyecto = st.selectbox("Proyecto / Cliente asignado *", ["Seleccionar..."] + project_list)
        otro = ""
        if proyecto == "Otro":
            otro = st.text_input("Indicá tu proyecto / cliente *")
        accepted = st.checkbox("Entiendo que este asistente responde consultas generales y que los casos personales pueden ser derivados a People Care.")
        submit = st.form_submit_button("Ingresar a People Care")

    if submit:
        final_project = otro.strip() if proyecto == "Otro" else proyecto
        if not nombre.strip() or not apellido.strip() or proyecto == "Seleccionar..." or (proyecto == "Otro" and not otro.strip()):
            st.error("Completá nombre, apellido y proyecto/cliente.")
        elif not accepted:
            st.error("Necesitás aceptar la aclaración para continuar.")
        else:
            st.session_state.nombre = nombre.strip()
            st.session_state.apellido = apellido.strip()
            st.session_state.proyecto = final_project
            st.session_state.identified = True
            st.session_state.messages = []
            st.session_state.question_counts = {}
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
    st.caption("No ingreses documentación, datos médicos, contraseñas ni otra información sensible.")
    st.stop()

# -----------------------------
# CHAT
# -----------------------------
st.markdown(
    f'<span class="v-user">{st.session_state.nombre} {st.session_state.apellido} · {st.session_state.proyecto}</span>',
    unsafe_allow_html=True
)

if st.button("Cambiar colaborador / proyecto"):
    for k in ["identified","nombre","apellido","proyecto","messages","question_counts"]:
        st.session_state.pop(k, None)
    st.rerun()

if not st.session_state.messages:
    st.session_state.messages = [{
        "role":"assistant",
        "content":f"Hola, {st.session_state.nombre}. Soy el asistente de People Care de Voolkia. Puedo ayudarte con vacaciones, licencias, recibos, beneficios, capacitaciones y consultas operativas. ¿Qué necesitás saber?"
    }]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

question = st.chat_input("Escribí tu consulta...")

if question:
    st.session_state.messages.append({"role":"user","content":question})
    with st.chat_message("user"):
        st.write(question)

    nq = normalize(question)
    st.session_state.question_counts[nq] = st.session_state.question_counts.get(nq, 0) + 1

    direct_alert = detect_alert(question, alertas)
    repeated = st.session_state.question_counts[nq] >= 3

    faq, score = find_faq(question, faqs)

    if direct_alert:
        answer = (
            "Esta consulta requiere atención personalizada de People Care. "
            "Voy a dejarla señalada para que el equipo pueda revisarla. "
            "Si se trata de una urgencia, utilizá también el canal habitual de People Care."
        )
        status = "RED_FLAG"
        alert_type = direct_alert["tipo"]
        send_alert(
            f"[People Care] Alerta {direct_alert['prioridad']} - {alert_type}",
            alert_body(question, answer, status, alert_type)
        )

    elif repeated:
        answer = (
            "Veo que necesitás más ayuda con este tema. Para evitar que sigas dando vueltas con la misma consulta, "
            "la voy a derivar a People Care para atención personalizada."
        )
        status = "REITERADA"
        alert_type = "Consulta reiterada"
        send_alert(
            "[People Care] Consulta reiterada",
            alert_body(question, answer, status, alert_type)
        )

    elif faq is not None and score >= MATCH_THRESHOLD:
        answer = str(faq["Respuesta"])
        status = "RESPONDIDA"
        alert_type = ""
    else:
        answer = (
            "No encontré una respuesta suficientemente clara en la base de People Care. "
            "Prefiero no darte información incorrecta. Voy a registrar tu consulta para que el equipo pueda revisarla "
            "y evaluar incorporarla a las FAQs."
        )
        status = "SIN_RESPUESTA"
        alert_type = "Sin respuesta"
        send_alert(
            "[People Care] Nueva consulta sin respuesta",
            alert_body(question, answer, status, alert_type)
        )

    st.session_state.messages.append({"role":"assistant","content":answer})
    with st.chat_message("assistant"):
        st.write(answer)

st.markdown("---")
st.markdown(
    '<div class="small-note">Este asistente brinda información general basada en la base autorizada de People Care. '
    'No reemplaza la atención personalizada de HR.</div>',
    unsafe_allow_html=True
)
