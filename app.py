
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
# CONFIGURACIÓN GENERAL
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
# ESTILO / FRONT VOOLKIA
# =========================================================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

:root {
    --v-orange: #FF7000;
    --v-orange-2: #FF9100;
    --v-dark: #240300;
    --v-gray: #EDEDED;
    --v-bg: #F7F5F3;
    --v-white: #FFFFFF;
    --v-muted: #6F625D;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top right, rgba(255,145,0,.10), transparent 28%),
        linear-gradient(180deg, #FAF9F8 0%, #F4F1EF 100%);
    color: var(--v-dark);
}

[data-testid="stHeader"] {
    background: transparent;
}

[data-testid="stToolbar"] {
    right: 1rem;
}

.block-container {
    max-width: 1180px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

/* ---------- HERO ---------- */
.pc-hero {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, #FF7000 0%, #FF9100 100%);
    border-radius: 30px;
    padding: 30px 34px;
    color: white;
    margin-bottom: 24px;
    box-shadow: 0 18px 45px rgba(36, 3, 0, .12);
}

.pc-hero:after {
    content: "";
    position: absolute;
    width: 240px;
    height: 240px;
    border-radius: 50%;
    right: -80px;
    top: -100px;
    background: rgba(255,255,255,.12);
}

.pc-brand {
    font-size: .78rem;
    font-weight: 800;
    letter-spacing: .16em;
    text-transform: uppercase;
    margin-bottom: 10px;
}

.pc-hero h1 {
    margin: 0;
    font-size: clamp(2rem, 4vw, 3rem);
    line-height: 1;
    font-weight: 800;
}

.pc-hero p {
    margin: 12px 0 0 0;
    max-width: 700px;
    font-size: 1rem;
    line-height: 1.55;
    color: white;
}

/* ---------- CARDS ---------- */
.pc-card {
    background: rgba(255,255,255,.96);
    border: 1px solid #E9E3DF;
    border-radius: 26px;
    padding: 28px;
    box-shadow: 0 12px 32px rgba(36, 3, 0, .06);
}

.pc-info-card {
    background: linear-gradient(180deg, #FFF7F1 0%, #FFFFFF 100%);
    border: 1px solid #F0D7C6;
    border-radius: 26px;
    padding: 28px;
    box-shadow: 0 12px 32px rgba(36, 3, 0, .05);
    min-height: 100%;
}

.pc-kicker {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: var(--v-dark);
    color: white;
    padding: 8px 13px;
    border-radius: 999px;
    font-size: .78rem;
    font-weight: 700;
    margin-bottom: 16px;
}

.pc-title {
    font-size: 1.75rem;
    font-weight: 800;
    color: var(--v-dark);
    line-height: 1.15;
    margin: 0 0 8px 0;
}

.pc-subtitle {
    color: var(--v-muted);
    font-size: .98rem;
    line-height: 1.55;
    margin-bottom: 20px;
}

.pc-info-title {
    font-size: 1.12rem;
    font-weight: 800;
    color: var(--v-dark);
    margin: 4px 0 14px 0;
}

.pc-list {
    margin: 0;
    padding: 0;
    list-style: none;
}

.pc-list li {
    display: flex;
    gap: 10px;
    align-items: flex-start;
    margin-bottom: 14px;
    color: #443936;
    font-size: .93rem;
    line-height: 1.45;
}

.pc-dot {
    width: 9px;
    height: 9px;
    background: var(--v-orange);
    border-radius: 50%;
    margin-top: 6px;
    flex: 0 0 9px;
}

.pc-note {
    background: #FFF2E7;
    border: 1px solid #FFD7BC;
    border-radius: 17px;
    padding: 14px 16px;
    color: #6B4935;
    font-size: .88rem;
    line-height: 1.5;
    margin-top: 20px;
}

.pc-security {
    margin-top: 14px;
    color: #7A6D68;
    font-size: .78rem;
    line-height: 1.45;
}

/* ---------- INPUTS: CONTRASTE CORREGIDO ---------- */
.stTextInput label,
.stSelectbox label,
.stCheckbox label,
.stTextArea label {
    color: var(--v-dark) !important;
    font-weight: 650 !important;
    font-size: .91rem !important;
}

.stTextInput input,
.stTextArea textarea {
    background: #FFFFFF !important;
    color: var(--v-dark) !important;
    border: 1px solid #CFC7C2 !important;
    border-radius: 14px !important;
    min-height: 46px;
    caret-color: var(--v-orange) !important;
}

.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: #9B918D !important;
    opacity: 1 !important;
}

div[data-baseweb="select"] > div {
    background: #FFFFFF !important;
    border: 1px solid #CFC7C2 !important;
    border-radius: 14px !important;
    min-height: 46px;
}

div[data-baseweb="select"] span,
div[data-baseweb="select"] div {
    color: var(--v-dark) !important;
}

[data-baseweb="popover"] {
    color: var(--v-dark) !important;
}

[data-baseweb="menu"] {
    background: #FFFFFF !important;
}

[data-baseweb="menu"] li {
    color: var(--v-dark) !important;
}

.stCheckbox p {
    color: #493D39 !important;
    font-size: .87rem !important;
    line-height: 1.4 !important;
}

/* ---------- BOTONES ---------- */
div.stButton > button,
div.stForm button {
    width: 100%;
    min-height: 48px;
    border: none !important;
    border-radius: 999px !important;
    background: linear-gradient(135deg, #FF7000 0%, #FF9100 100%) !important;
    color: #FFFFFF !important;
    font-weight: 800 !important;
    box-shadow: 0 8px 20px rgba(255,112,0,.24);
    transition: all .18s ease;
}

div.stButton > button:hover,
div.stForm button:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 24px rgba(255,112,0,.30);
    color: #FFFFFF !important;
}

/* ---------- ALERTAS ---------- */
[data-testid="stAlert"] {
    border-radius: 16px;
}

/* ---------- CHAT ---------- */
.pc-userbar {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
    margin-bottom: 16px;
}

.pc-userpill {
    display: inline-block;
    background: #FFFFFF;
    border: 1px solid #E1DAD6;
    color: var(--v-dark);
    padding: 9px 14px;
    border-radius: 999px;
    font-size: .83rem;
    font-weight: 650;
}

.pc-chat-intro {
    background: #FFFFFF;
    border: 1px solid #EAE3DF;
    border-radius: 22px;
    padding: 18px 20px;
    margin-bottom: 16px;
    color: #4F423E;
    box-shadow: 0 8px 22px rgba(36,3,0,.04);
}

[data-testid="stChatMessage"] {
    background: #FFFFFF;
    border: 1px solid #E9E3DF;
    border-radius: 20px;
    padding: 10px 14px;
    box-shadow: 0 6px 18px rgba(36,3,0,.035);
}

[data-testid="stChatInput"] {
    background: #FFFFFF;
    border-radius: 18px;
}

hr {
    border-color: #E5DEDA !important;
}

.pc-footer {
    color: #756A66;
    font-size: .79rem;
    line-height: 1.45;
    margin-top: 12px;
}

@media (max-width: 800px) {
    .block-container {
        padding-top: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }

    .pc-hero {
        padding: 24px 22px;
        border-radius: 24px;
    }

    .pc-card,
    .pc-info-card {
        padding: 22px;
        border-radius: 22px;
    }
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# CABECERA
# =========================================================
st.markdown(
    """
<div class="pc-hero">
    <div class="pc-brand">VOOLKIA · PEOPLE CARE</div>
    <h1>People Care</h1>
    <p>
        Un espacio simple para resolver consultas frecuentes de HR y ayudarte
        a encontrar rápidamente la información que necesitás.
    </p>
</div>
""",
    unsafe_allow_html=True,
)


# =========================================================
# FUNCIONES
# =========================================================
def normalize(text: str) -> str:
    text = str(text or "").lower().strip()
    text = "".join(
        c
        for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )
    text = re.sub(r"[^a-z0-9ñáéíóúü\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


@st.cache_data(ttl=300)
def load_excel():
    """
    Fuente principal:
    Google Sheets nativo exportado temporalmente como XLSX.

    Si FAQ_DRIVE_FILE_ID no está configurado, usa el Excel local
    incluido en el repositorio como respaldo.
    """
    if FAQ_DRIVE_FILE_ID:
        url = (
            f"https://docs.google.com/spreadsheets/d/"
            f"{FAQ_DRIVE_FILE_ID}/export?format=xlsx"
        )

        response = requests.get(url, timeout=30)
        response.raise_for_status()

        content = io.BytesIO(response.content)

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
        faqs["Activo"]
        .astype(str)
        .str.lower()
        .isin(["sí", "si", "true", "1"])
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

        if (
            "MISMA_PREGUNTA_3_VECES" in triggers
            or "SIN_COINCIDENCIA_FAQ" in triggers
        ):
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
    """
    Si SMTP todavía no está configurado, la app sigue funcionando.
    Solo omite el envío del correo.
    """
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


# =========================================================
# CARGA DE BASE
# =========================================================
try:
    faqs, alertas, proyectos_df = load_excel()
except Exception as e:
    st.error(
        "No pude cargar la base de conocimiento. "
        "Revisá el Google Sheet o la configuración de Streamlit."
    )
    st.caption(str(e))
    st.stop()


project_list = (
    proyectos_df[
        proyectos_df["Activo"]
        .astype(str)
        .str.lower()
        .isin(["sí", "si", "true", "1"])
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
# PANTALLA DE ACCESO
# =========================================================
if not st.session_state.identified:
    left, right = st.columns([1.15, 0.85], gap="large")

    with left:
        st.markdown(
            """
<div class="pc-card">
    <div class="pc-kicker">Ingreso de colaborador</div>
    <div class="pc-title">Antes de comenzar</div>
    <div class="pc-subtitle">
        Identificate para que People Care pueda contextualizar tu consulta
        si necesitás atención personalizada.
    </div>
</div>
""",
            unsafe_allow_html=True,
        )

        with st.form("identificacion"):
            nombre = st.text_input(
                "Nombre *",
                placeholder="Ej.: Marcela",
            )

            apellido = st.text_input(
                "Apellido *",
                placeholder="Ej.: Infante",
            )

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
                "Entiendo que este asistente responde consultas generales "
                "y que los casos personales pueden ser derivados a People Care."
            )

            submit = st.form_submit_button("Ingresar a People Care")

    with right:
        st.markdown(
            """
<div class="pc-info-card">
    <div class="pc-info-title">¿Qué podés consultar?</div>

    <ul class="pc-list">
        <li><span class="pc-dot"></span><span>Vacaciones, licencias y certificados.</span></li>
        <li><span class="pc-dot"></span><span>Recibos de sueldo y consultas frecuentes.</span></li>
        <li><span class="pc-dot"></span><span>Beneficios, prepaga y referidos.</span></li>
        <li><span class="pc-dot"></span><span>Capacitaciones y Voolkia Learning.</span></li>
        <li><span class="pc-dot"></span><span>Cambio de domicilio, equipamiento y consultas operativas.</span></li>
    </ul>

    <div class="pc-note">
        Si el asistente no encuentra una respuesta clara, no inventará información:
        la consulta podrá ser derivada a People Care para revisión.
    </div>

    <div class="pc-security">
        No ingreses contraseñas, documentación médica, certificados,
        datos bancarios ni otra información sensible.
    </div>
</div>
""",
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
    f"""
<div class="pc-userbar">
    <span class="pc-userpill">
        {st.session_state.nombre} {st.session_state.apellido}
    </span>
    <span class="pc-userpill">
        Proyecto / Cliente: {st.session_state.proyecto}
    </span>
</div>
""",
    unsafe_allow_html=True,
)

change_col, spacer_col = st.columns([0.28, 0.72])

with change_col:
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
    """
<div class="pc-chat-intro">
    Podés escribir tu consulta con tus propias palabras. El asistente buscará
    una respuesta en la base autorizada de People Care.
</div>
""",
    unsafe_allow_html=True,
)


if "messages" not in st.session_state or not st.session_state.messages:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                f"Hola, {st.session_state.nombre}. "
                "Soy el asistente de People Care de Voolkia. "
                "Puedo ayudarte con vacaciones, licencias, recibos, beneficios, "
                "capacitaciones y consultas operativas. ¿Qué necesitás saber?"
            ),
        }
    ]


for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])


question = st.chat_input("Escribí tu consulta...")


if question:
    st.session_state.messages.append(
        {"role": "user", "content": question}
    )

    with st.chat_message("user"):
        st.write(question)

    nq = normalize(question)

    st.session_state.question_counts[nq] = (
        st.session_state.question_counts.get(nq, 0) + 1
    )

    direct_alert = detect_alert(question, alertas)
    repeated = st.session_state.question_counts[nq] >= 3
    faq, score = find_faq(question, faqs)

    if direct_alert:
        answer = (
            "Esta consulta requiere atención personalizada de People Care. "
            "Voy a dejarla señalada para que el equipo pueda revisarla. "
            "Si se trata de una situación urgente, utilizá también el canal habitual "
            "de People Care."
        )

        status = "RED_FLAG"
        alert_type = direct_alert["tipo"]

        send_alert(
            f"[People Care] Alerta {direct_alert['prioridad']} - {alert_type}",
            alert_body(
                question,
                answer,
                status,
                alert_type,
            ),
        )

    elif repeated:
        answer = (
            "Veo que necesitás más ayuda con este tema. "
            "Para evitar que sigas dando vueltas con la misma consulta, "
            "la voy a derivar a People Care para atención personalizada."
        )

        status = "REITERADA"
        alert_type = "Consulta reiterada"

        send_alert(
            "[People Care] Consulta reiterada",
            alert_body(
                question,
                answer,
                status,
                alert_type,
            ),
        )

    elif faq is not None and score >= MATCH_THRESHOLD:
        answer = str(faq["Respuesta"])
        status = "RESPONDIDA"
        alert_type = ""

    else:
        answer = (
            "No encontré una respuesta suficientemente clara en la base de People Care. "
            "Prefiero no darte información incorrecta. "
            "Voy a registrar tu consulta para que el equipo pueda revisarla "
            "y evaluar incorporarla a las FAQs."
        )

        status = "SIN_RESPUESTA"
        alert_type = "Sin respuesta"

        send_alert(
            "[People Care] Nueva consulta sin respuesta",
            alert_body(
                question,
                answer,
                status,
                alert_type,
            ),
        )

    st.session_state.messages.append(
        {"role": "assistant", "content": answer}
    )

    with st.chat_message("assistant"):
        st.write(answer)


st.markdown("---")

st.markdown(
    """
<div class="pc-footer">
    Este asistente brinda información general basada en la base autorizada
    de People Care. No reemplaza la atención personalizada de HR.
</div>
""",
    unsafe_allow_html=True,
)
