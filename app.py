
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


# =========================================================
# CONFIGURACIÓN
# =========================================================
st.set_page_config(
    page_title="People Care | Voolkia",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ALERT_EMAIL = st.secrets.get("ALERT_EMAIL", "marcela@voolkia.com")
FAQ_DRIVE_FILE_ID = st.secrets.get("FAQ_DRIVE_FILE_ID", "")
FAQ_LOCAL_FILE = "People_Care_Voolkia_Base_Conocimiento.xlsx"
MATCH_THRESHOLD = 62


# =========================================================
# ESTILOS
# =========================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

:root {
    --orange: #FF7000;
    --orange2: #FF9100;
    --dark: #240300;
    --gray: #EDEDED;
    --bg: #F7F5F3;
    --muted: #6E625D;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: linear-gradient(180deg, #FBFAF9 0%, #F4F1EF 100%);
    color: var(--dark);
}

.block-container {
    max-width: 1120px;
    padding-top: 1.5rem;
    padding-bottom: 3rem;
}

/* HERO */
.hero {
    background: linear-gradient(135deg, #FF7000 0%, #FF9100 100%);
    color: white;
    border-radius: 28px;
    padding: 28px 32px;
    margin-bottom: 24px;
    box-shadow: 0 16px 40px rgba(36,3,0,.10);
}

.hero-brand {
    font-size: .76rem;
    font-weight: 800;
    letter-spacing: .16em;
    text-transform: uppercase;
    margin-bottom: 8px;
}

.hero-title {
    font-size: 2.4rem;
    line-height: 1;
    font-weight: 800;
    margin: 0;
}

.hero-text {
    font-size: 1rem;
    line-height: 1.5;
    margin-top: 10px;
    max-width: 720px;
    color: white;
}

/* CARDS */
.card {
    background: #FFFFFF;
    border: 1px solid #E9E1DD;
    border-radius: 24px;
    padding: 26px;
    box-shadow: 0 10px 28px rgba(36,3,0,.05);
}

.info-card {
    background: #FFF8F3;
    border: 1px solid #F1D5C2;
    border-radius: 24px;
    padding: 26px;
    box-shadow: 0 10px 28px rgba(36,3,0,.04);
    height: 100%;
}

.badge {
    display: inline-block;
    background: #240300;
    color: #FFFFFF;
    border-radius: 999px;
    padding: 7px 12px;
    font-size: .76rem;
    font-weight: 800;
    margin-bottom: 14px;
}

.title {
    color: #240300;
    font-size: 1.65rem;
    font-weight: 800;
    margin: 0 0 8px 0;
}

.subtitle {
    color: #6E625D;
    font-size: .96rem;
    line-height: 1.5;
    margin-bottom: 4px;
}

.info-title {
    color: #240300;
    font-size: 1.15rem;
    font-weight: 800;
    margin-bottom: 16px;
}

.info-item {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    margin-bottom: 13px;
    color: #443936;
    font-size: .92rem;
    line-height: 1.45;
}

.dot {
    width: 9px;
    height: 9px;
    min-width: 9px;
    border-radius: 50%;
    background: #FF7000;
    margin-top: 6px;
}

.note {
    margin-top: 20px;
    padding: 14px 15px;
    border-radius: 16px;
    background: #FFF0E4;
    border: 1px solid #FFD0B1;
    color: #654634;
    font-size: .87rem;
    line-height: 1.45;
}

.security {
    margin-top: 15px;
    color: #776B66;
    font-size: .78rem;
    line-height: 1.4;
}

/* LABELS */
.stTextInput label,
.stSelectbox label,
.stCheckbox label,
.stTextArea label {
    color: #240300 !important;
    font-weight: 600 !important;
}

/* INPUT */
.stTextInput input,
.stTextArea textarea {
    background: #FFFFFF !important;
    color: #240300 !important;
    border: 1px solid #BFB5B0 !important;
    border-radius: 12px !important;
}

.stTextInput input::placeholder {
    color: #9A8F89 !important;
}

/* SELECTBOX */
div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    border: 1px solid #BFB5B0 !important;
    border-radius: 12px !important;
}

div[data-baseweb="select"] span {
    color: #240300 !important;
}

/* CHECKBOX */
.stCheckbox p {
    color: #463B37 !important;
    font-size: .87rem !important;
}

/* BOTONES */
div.stButton > button,
div.stFormSubmitButton > button {
    width: 100%;
    border: none !important;
    border-radius: 999px !important;
    min-height: 46px;
    background: linear-gradient(135deg, #FF7000 0%, #FF9100 100%) !important;
    color: white !important;
    font-weight: 800 !important;
    box-shadow: 0 8px 18px rgba(255,112,0,.22);
}

div.stButton > button:hover,
div.stFormSubmitButton > button:hover {
    color: white !important;
    border: none !important;
}

/* USER BAR */
.user-pill {
    display: inline-block;
    background: #FFFFFF;
    border: 1px solid #E2DAD6;
    border-radius: 999px;
    padding: 8px 13px;
    margin: 0 6px 10px 0;
    color: #240300;
    font-size: .82rem;
    font-weight: 650;
}

.chat-help {
    background: #FFFFFF;
    border: 1px solid #E8E1DD;
    border-radius: 18px;
    padding: 14px 16px;
    color: #5B4F4A;
    font-size: .9rem;
    margin: 6px 0 15px 0;
}

/* CHAT */
[data-testid="stChatMessage"] {
    background: #FFFFFF;
    border: 1px solid #E8E1DD;
    border-radius: 18px;
    padding: 10px 14px;
}

.footer {
    color: #766A65;
    font-size: .78rem;
    line-height: 1.45;
}

@media (max-width: 800px) {
    .block-container {
        padding: 1rem;
    }
    .hero {
        padding: 23px 21px;
        border-radius: 22px;
    }
    .hero-title {
        font-size: 2rem;
    }
    .card, .info-card {
        padding: 20px;
        border-radius: 20px;
    }
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# HERO
# =========================================================
st.markdown(
    '<div class="hero">'
    '<div class="hero-brand">VOOLKIA · PEOPLE CARE</div>'
    '<div class="hero-title">People Care</div>'
    '<div class="hero-text">Un espacio simple para resolver consultas frecuentes de HR y encontrar rápidamente la información que necesitás.</div>'
    '</div>',
    unsafe_allow_html=True,
)


# =========================================================
# FUNCIONES
# =========================================================
def normalize(text):
    text = str(text or "").lower().strip()
    text = "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


@st.cache_data(ttl=300)
def load_excel():
    if FAQ_DRIVE_FILE_ID:
        url = f"https://docs.google.com/spreadsheets/d/{FAQ_DRIVE_FILE_ID}/export?format=xlsx"
        r = requests.get(url, timeout=30)
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

    active = faqs[
        faqs["Activo"].astype(str).str.lower().isin(["sí", "si", "true", "1"])
    ]

    for _, row in active.iterrows():
        searchable = (
            f"{row['Pregunta']} "
            f"{row.get('Palabras_clave', '')} "
            f"{row.get('Categoría', '')}"
        )

        score = max(
            fuzz.token_set_ratio(q, normalize(searchable)),
            fuzz.partial_ratio(q, normalize(row["Pregunta"])),
        )

        if score > best_score:
            best_score = score
            best = row

    return best, best_score


def detect_alert(question, alertas):
    q = normalize(question)

    for _, row in alertas.iterrows():
        triggers = str(row.get("Disparadores", ""))

        if "MISMA_PREGUNTA_3_VECES" in triggers or "SIN_COINCIDENCIA_FAQ" in triggers:
            continue

        for trigger in triggers.split(";"):
            t = normalize(trigger)
            if t and t in q:
                return {
                    "tipo": str(row["Tipo"]),
                    "prioridad": str(row["Prioridad"]),
                    "accion": str(row["Acción"]),
                }

    return None


def send_alert(subject, body):
    host = st.secrets.get("SMTP_HOST", "")
    port = int(st.secrets.get("SMTP_PORT", 587))
    user = st.secrets.get("SMTP_USER", "")
    password = st.secrets.get("SMTP_PASSWORD", "")

    if not all([host, user, password]):
        return False

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ALERT_EMAIL
    msg.set_content(body)

    with smtplib.SMTP(host, port, timeout=20) as server:
        server.starttls()
        server.login(user, password)
        server.send_message(msg)

    return True


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
"""


# =========================================================
# CARGAR BASE
# =========================================================
try:
    faqs, alertas, proyectos_df = load_excel()
except Exception as e:
    st.error("No pude cargar la base de conocimiento.")
    st.caption(str(e))
    st.stop()


project_list = (
    proyectos_df[
        proyectos_df["Activo"].astype(str).str.lower().isin(["sí", "si", "true", "1"])
    ]["Proyecto_Cliente"]
    .dropna()
    .astype(str)
    .tolist()
)


# =========================================================
# SESIÓN
# =========================================================
if "identified" not in st.session_state:
    st.session_state.identified = False


# =========================================================
# LOGIN
# =========================================================
if not st.session_state.identified:

    col_form, col_info = st.columns([1.12, 0.88], gap="large")

    with col_form:
        st.markdown(
            '<div class="card">'
            '<div class="badge">Ingreso de colaborador</div>'
            '<div class="title">Antes de comenzar</div>'
            '<div class="subtitle">Identificate para que People Care pueda contextualizar tu consulta si necesitás atención personalizada.</div>'
            '</div>',
            unsafe_allow_html=True,
        )

        st.write("")

        with st.form("identificacion"):
            nombre = st.text_input("Nombre *", placeholder="Ej.: Marcela")
            apellido = st.text_input("Apellido *", placeholder="Ej.: Infante")

            proyecto = st.selectbox(
                "Proyecto / Cliente asignado *",
                ["Seleccionar..."] + project_list,
            )

            otro = ""
            if proyecto == "Otro":
                otro = st.text_input(
                    "Indicá tu proyecto / cliente *",
                    placeholder="Ej.: Cliente / proyecto",
                )

            accepted = st.checkbox(
                "Entiendo que este asistente responde consultas generales y "
                "que los casos personales pueden ser derivados a People Care."
            )

            submit = st.form_submit_button("Ingresar a People Care")

    with col_info:
        st.markdown(
            '<div class="info-card">'
            '<div class="info-title">¿Qué podés consultar?</div>'
            '<div class="info-item"><span class="dot"></span><span>Vacaciones, licencias y certificados.</span></div>'
            '<div class="info-item"><span class="dot"></span><span>Recibos de sueldo y consultas frecuentes.</span></div>'
            '<div class="info-item"><span class="dot"></span><span>Beneficios, prepaga y referidos.</span></div>'
            '<div class="info-item"><span class="dot"></span><span>Capacitaciones y Voolkia Learning.</span></div>'
            '<div class="info-item"><span class="dot"></span><span>Cambio de domicilio, equipamiento y consultas operativas.</span></div>'
            '<div class="note">Si el asistente no encuentra una respuesta clara, no inventará información. La consulta podrá ser derivada a People Care para revisión.</div>'
            '<div class="security">No ingreses contraseñas, documentación médica, certificados, datos bancarios ni otra información sensible.</div>'
            '</div>',
            unsafe_allow_html=True,
        )

    if submit:
        final_project = otro.strip() if proyecto == "Otro" else proyecto

        if (
            not nombre.strip()
            or not apellido.strip()
            or proyecto == "Seleccionar..."
            or (proyecto == "Otro" and not otro.strip())
        ):
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

    st.stop()


# =========================================================
# CHAT
# =========================================================
st.markdown(
    f'<span class="user-pill">{st.session_state.nombre} {st.session_state.apellido}</span>'
    f'<span class="user-pill">Proyecto / Cliente: {st.session_state.proyecto}</span>',
    unsafe_allow_html=True,
)

left_btn, _ = st.columns([0.3, 0.7])

with left_btn:
    if st.button("Cambiar colaborador / proyecto"):
        for key in [
            "identified",
            "nombre",
            "apellido",
            "proyecto",
            "messages",
            "question_counts",
        ]:
            st.session_state.pop(key, None)
        st.rerun()


st.markdown(
    '<div class="chat-help">Escribí tu consulta con tus propias palabras. El asistente buscará una respuesta en la base autorizada de People Care.</div>',
    unsafe_allow_html=True,
)


if "messages" not in st.session_state or not st.session_state.messages:
    st.session_state.messages = [{
        "role": "assistant",
        "content": (
            f"Hola, {st.session_state.nombre}. Soy el asistente de People Care de Voolkia. "
            "Puedo ayudarte con vacaciones, licencias, recibos, beneficios, "
            "capacitaciones y consultas operativas. ¿Qué necesitás saber?"
        )
    }]


for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])


question = st.chat_input("Escribí tu consulta...")


if question:
    st.session_state.messages.append({"role": "user", "content": question})

    with st.chat_message("user"):
        st.write(question)

    normalized = normalize(question)
    st.session_state.question_counts[normalized] = (
        st.session_state.question_counts.get(normalized, 0) + 1
    )

    direct_alert = detect_alert(question, alertas)
    repeated = st.session_state.question_counts[normalized] >= 3
    faq, score = find_faq(question, faqs)

    if direct_alert:
        answer = (
            "Esta consulta requiere atención personalizada de People Care. "
            "Voy a dejarla señalada para que el equipo pueda revisarla."
        )
        status = "RED_FLAG"
        alert_type = direct_alert["tipo"]

        send_alert(
            f"[People Care] Alerta {direct_alert['prioridad']} - {alert_type}",
            alert_body(question, answer, status, alert_type),
        )

    elif repeated:
        answer = (
            "Veo que necesitás más ayuda con este tema. "
            "La consulta será derivada a People Care para atención personalizada."
        )
        status = "REITERADA"
        alert_type = "Consulta reiterada"

        send_alert(
            "[People Care] Consulta reiterada",
            alert_body(question, answer, status, alert_type),
        )

    elif faq is not None and score >= MATCH_THRESHOLD:
        answer = str(faq["Respuesta"])

    else:
        answer = (
            "No encontré una respuesta suficientemente clara en la base de People Care. "
            "Prefiero no darte información incorrecta. "
            "La consulta quedará identificada para que People Care pueda revisarla."
        )

        send_alert(
            "[People Care] Nueva consulta sin respuesta",
            alert_body(question, answer, "SIN_RESPUESTA", "Sin respuesta"),
        )

    st.session_state.messages.append({"role": "assistant", "content": answer})

    with st.chat_message("assistant"):
        st.write(answer)


st.markdown("---")
st.markdown(
    '<div class="footer">Este asistente brinda información general basada en la base autorizada de People Care. No reemplaza la atención personalizada de HR.</div>',
    unsafe_allow_html=True,
)
