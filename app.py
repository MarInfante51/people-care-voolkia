
import io
import re
import smtplib
import unicodedata
from datetime import datetime
from email.message import EmailMessage
from html import escape

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
# ESTILO
# =========================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: "Inter", sans-serif;
    }

    .stApp {
        background: #F7F5F3;
        color: #240300;
    }

    .block-container {
        max-width: 1080px;
        padding-top: 1.6rem;
        padding-bottom: 3rem;
    }

    .pc-hero {
        background: linear-gradient(135deg, #FF7000 0%, #FF9100 100%);
        border-radius: 28px;
        padding: 28px 32px;
        margin-bottom: 22px;
        color: #FFFFFF;
        box-shadow: 0 14px 34px rgba(36, 3, 0, 0.10);
    }

    .pc-brand {
        font-size: 0.76rem;
        font-weight: 800;
        letter-spacing: 0.14em;
        margin-bottom: 8px;
    }

    .pc-title {
        font-size: 2.35rem;
        font-weight: 800;
        line-height: 1.05;
        margin-bottom: 8px;
    }

    .pc-subtitle {
        font-size: 0.98rem;
        line-height: 1.5;
        max-width: 760px;
    }

    /* Contenedores nativos de Streamlit */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #FFFFFF;
        border-color: #E6DDD8 !important;
        border-radius: 22px !important;
        box-shadow: 0 8px 24px rgba(36, 3, 0, 0.045);
    }

    /* Labels */
    .stTextInput label,
    .stTextArea label,
    .stSelectbox label,
    .stCheckbox label {
        color: #240300 !important;
        font-weight: 600 !important;
    }

    .stCheckbox p {
        color: #493D39 !important;
        font-size: 0.88rem !important;
    }

    /* Inputs */
    .stTextInput input,
    .stTextArea textarea {
        background: #FFFFFF !important;
        color: #240300 !important;
        border: 1px solid #BDB3AE !important;
        border-radius: 12px !important;
        caret-color: #FF7000 !important;
    }

    .stTextInput input::placeholder,
    .stTextArea textarea::placeholder {
        color: #988C87 !important;
        opacity: 1 !important;
    }

    /* Select */
    div[data-baseweb="select"] > div {
        background: #FFFFFF !important;
        border: 1px solid #BDB3AE !important;
        border-radius: 12px !important;
    }

    div[data-baseweb="select"] span {
        color: #240300 !important;
    }

    /* Botones */
    div.stButton > button,
    div.stFormSubmitButton > button {
        border: 0 !important;
        border-radius: 999px !important;
        min-height: 46px;
        background: linear-gradient(135deg, #FF7000 0%, #FF9100 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        box-shadow: 0 7px 16px rgba(255, 112, 0, 0.18);
    }

    div.stButton > button:hover,
    div.stFormSubmitButton > button:hover {
        color: #FFFFFF !important;
        border: 0 !important;
    }

    .pc-pill {
        display: inline-block;
        background: #FFFFFF;
        border: 1px solid #E0D8D3;
        color: #240300;
        border-radius: 999px;
        padding: 8px 13px;
        margin: 0 6px 10px 0;
        font-size: 0.82rem;
        font-weight: 600;
    }

    .pc-help {
        background: #FFFFFF;
        border: 1px solid #E6DDD8;
        border-radius: 18px;
        padding: 14px 16px;
        margin: 8px 0 16px 0;
        color: #594C47;
        font-size: 0.91rem;
    }

    .pc-assistant {
        background: #FFFFFF;
        border: 1px solid #E5DDD8;
        border-radius: 18px 18px 18px 6px;
        padding: 15px 17px;
        margin: 10px 0;
        color: #2F2926;
        line-height: 1.52;
    }

    .pc-user {
        background: #FFF0E5;
        border: 1px solid #FFD1B3;
        border-radius: 18px 18px 6px 18px;
        padding: 15px 17px;
        margin: 10px 0 10px auto;
        color: #352B27;
        line-height: 1.52;
        max-width: 82%;
    }

    .pc-label {
        color: #FF7000;
        font-size: 0.72rem;
        font-weight: 800;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 5px;
    }

    .pc-footer {
        color: #756A65;
        font-size: 0.79rem;
        line-height: 1.45;
        margin-top: 12px;
    }

    @media (max-width: 800px) {
        .block-container {
            padding: 1rem;
        }

        .pc-hero {
            padding: 22px 20px;
            border-radius: 22px;
        }

        .pc-title {
            font-size: 2rem;
        }

        .pc-user {
            max-width: 100%;
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
        <div class="pc-title">People Care</div>
        <div class="pc-subtitle">
            Un espacio simple para resolver consultas frecuentes de HR
            y encontrar rápidamente la información que necesitás.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FUNCIONES
# =========================================================
def normalize(text):
    text = str(text or "").lower().strip()
    text = "".join(
        c
        for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


@st.cache_data(ttl=300)
def load_excel():
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
    second_score = 0

    active = faqs[
        faqs["Activo"]
        .astype(str)
        .str.lower()
        .isin(["sí", "si", "true", "1"])
    ]

    scores = []

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

        scores.append((score, row))

    scores.sort(key=lambda x: x[0], reverse=True)

    if scores:
        best_score, best = scores[0]

    if len(scores) > 1:
        second_score = scores[1][0]

    return best, best_score, second_score


def is_contextual_followup(question):
    q = normalize(question)

    if len(q.split()) <= 7:
        return True

    starters = (
        "y como",
        "como hago",
        "como lo hago",
        "como solicito",
        "como pido",
        "donde",
        "y eso",
        "y ahi",
        "y despues",
        "que hago",
        "a quien",
    )

    return q.startswith(starters)


def get_contextual_query(question):
    """
    Para preguntas cortas o de seguimiento, agrega el último tema conversado.
    No cambia la respuesta: solo mejora la búsqueda de la FAQ correcta.
    """
    if not is_contextual_followup(question):
        return question

    previous_user = ""
    previous_assistant = ""

    for msg in reversed(st.session_state.messages[:-1]):
        if not previous_assistant and msg["role"] == "assistant":
            previous_assistant = str(msg["content"])
        elif not previous_user and msg["role"] == "user":
            previous_user = str(msg["content"])

        if previous_user and previous_assistant:
            break

    context = f"{previous_user} {previous_assistant}".strip()

    if context:
        return f"{question} CONTEXTO ANTERIOR: {context}"

    return question


def direct_people_care_contact(question):
    q = normalize(question)

    mentions_people_care = (
        "people care" in q
        or "ppc" in q
        or "recursos humanos" in q
        or "rrhh" in q
        or "hr" in q
    )

    contact_intent = any(
        phrase in q
        for phrase in [
            "como contacto",
            "como escribo",
            "como solicito",
            "como pido",
            "como hablo",
            "donde escribo",
            "a que mail",
            "cual es el mail",
            "mail de",
            "correo de",
        ]
    )

    if mentions_people_care and contact_intent:
        return (
            "Podés contactar a People Care enviando un mail a "
            "ppc@voolkia.com."
        )

    return None



def detect_human_state(question):
    q = normalize(question)

    frustration_terms = [
        "no me responden",
        "no me dan respuesta",
        "nadie me responde",
        "hace dias",
        "hace 2 dias",
        "hace 3 dias",
        "hace 4 dias",
        "hace una semana",
        "estoy cansado",
        "estoy cansada",
        "estoy enojado",
        "estoy enojada",
        "estoy molesto",
        "estoy molesta",
        "esto no sirve",
        "no sirve",
        "boludear",
        "me estan dando vueltas",
        "me dan vueltas",
        "quiero hablar con alguien",
        "necesito hablar con alguien",
    ]

    if any(term in q for term in frustration_terms):
        return "frustration"

    return None


def humanized_no_answer(question):
    name = st.session_state.get("nombre", "").strip()
    prefix = f"{name}, " if name else ""

    return (
        f"{prefix}entiendo la consulta. En este momento no tengo una respuesta "
        "validada sobre ese tema en la base de People Care, y prefiero no darte "
        "información que pueda ser incorrecta. Voy a dejarla identificada para "
        "que People Care pueda revisarla."
    )


def humanized_clarification():
    name = st.session_state.get("nombre", "").strip()
    prefix = f"{name}, " if name else ""

    return (
        f"{prefix}quiero asegurarme de entenderte bien para no responderte cualquier cosa. "
        "¿Podés contarme un poco más sobre qué necesitás?"
    )


def humanized_frustration_response():
    name = st.session_state.get("nombre", "").strip()
    prefix = f"{name}, " if name else ""

    return (
        f"{prefix}entiendo la frustración. Si ya venís esperando una respuesta, "
        "no quiero que sigas dando vueltas con el mismo tema. Voy a dejar esta "
        "situación marcada para atención personalizada de People Care. "
        "Si querés, también podés escribir directamente a ppc@voolkia.com."
    )


def humanized_bot_complaint_response():
    name = st.session_state.get("nombre", "").strip()
    prefix = f"{name}, " if name else ""

    return (
        f"{prefix}entiendo lo que decís. La idea de este bot es justamente darte "
        "respuestas rápidas cuando la información está validada y no hacerte perder "
        "tiempo. Si no tengo una respuesta segura, prefiero decirlo y derivar el tema "
        "a People Care antes que inventar. Podés contarme qué necesitás resolver y "
        "lo intento de nuevo."
    )


def is_bot_complaint(question):
    q = normalize(question)
    complaint_terms = [
        "este bot",
        "el bot",
        "esto no sirve",
        "boludear",
        "me hace perder tiempo",
        "perder tiempo",
    ]
    return any(term in q for term in complaint_terms)

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
            trigger_normalized = normalize(trigger)

            if trigger_normalized and trigger_normalized in q:
                return {
                    "tipo": str(row["Tipo"]),
                    "prioridad": str(row["Prioridad"]),
                }

    return None


def send_alert(subject, body):
    host = st.secrets.get("SMTP_HOST", "")
    port = int(st.secrets.get("SMTP_PORT", 587))
    user = st.secrets.get("SMTP_USER", "")
    password = st.secrets.get("SMTP_PASSWORD", "")

    # La app sigue funcionando aunque el correo aún no esté configurado.
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
# CARGA DE LA BASE
# =========================================================
try:
    faqs, alertas, proyectos_df = load_excel()
except Exception as error:
    st.error(
        "No pude cargar la base de conocimiento. "
        "Revisá el Google Sheet o la configuración de Streamlit."
    )
    st.caption(str(error))
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
# ESTADO DE SESIÓN
# =========================================================
if "identified" not in st.session_state:
    st.session_state.identified = False

if "messages" not in st.session_state:
    st.session_state.messages = []

if "question_counts" not in st.session_state:
    st.session_state.question_counts = {}


# =========================================================
# INGRESO DEL COLABORADOR
# =========================================================
if not st.session_state.identified:
    col_form, col_info = st.columns([1.1, 0.9], gap="large")

    with col_form:
        with st.container(border=True):
            st.caption("INGRESO DE COLABORADOR")
            st.subheader("Antes de comenzar")
            st.write(
                "Identificate para que People Care pueda contextualizar tu consulta "
                "si necesitás atención personalizada."
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

                submit = st.form_submit_button(
                    "Ingresar a People Care",
                    use_container_width=True,
                )

    with col_info:
        with st.container(border=True):
            st.subheader("¿Qué podés consultar?")
            st.markdown(
                """
                - Vacaciones, licencias y certificados.
                - Recibos de sueldo y consultas frecuentes.
                - Beneficios, prepaga y referidos.
                - Capacitaciones y Voolkia Learning.
                - Cambio de domicilio, equipamiento y consultas operativas.
                """
            )

            st.info(
                "Si el asistente no encuentra una respuesta clara, no inventará "
                "información. La consulta podrá ser derivada a People Care para revisión."
            )

            st.caption(
                "No ingreses contraseñas, documentación médica, certificados, "
                "datos bancarios ni otra información sensible."
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
            st.session_state.question_counts = {}
            st.session_state.messages = [
                {
                    "role": "assistant",
                    "content": (
                        f"Hola, {nombre.strip()}. Soy el asistente de People Care "
                        "de Voolkia. Puedo ayudarte con vacaciones, licencias, recibos, "
                        "beneficios, capacitaciones y consultas operativas. "
                        "¿Qué necesitás saber?"
                    ),
                }
            ]
            st.rerun()

    st.stop()


# =========================================================
# ÁREA DE CONSULTAS
# =========================================================
st.markdown(
    f'<span class="pc-pill">{escape(st.session_state.nombre)} '
    f'{escape(st.session_state.apellido)}</span>'
    f'<span class="pc-pill">Proyecto / Cliente: '
    f'{escape(st.session_state.proyecto)}</span>',
    unsafe_allow_html=True,
)

btn_col, _ = st.columns([0.28, 0.72])

with btn_col:
    if st.button(
        "Cambiar colaborador / proyecto",
        use_container_width=True,
    ):
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
    '<div class="pc-help">'
    'Escribí tu consulta con tus propias palabras. '
    'El asistente buscará una respuesta en la base autorizada de People Care.'
    '</div>',
    unsafe_allow_html=True,
)


# =========================================================
# HISTORIAL
# =========================================================
for message in st.session_state.messages:
    safe_content = escape(str(message["content"]))

    if message["role"] == "assistant":
        st.markdown(
            '<div class="pc-assistant">'
            '<div class="pc-label">People Care</div>'
            f'{safe_content}'
            '</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="pc-user">'
            '<div class="pc-label">Tu consulta</div>'
            f'{safe_content}'
            '</div>',
            unsafe_allow_html=True,
        )


# =========================================================
# CAMPO REAL DE ESCRITURA
# =========================================================
with st.container(border=True):
    st.subheader("¿En qué podemos ayudarte?")
    st.caption("Escribí tu consulta y presioná Enviar consulta.")

    with st.form("consulta_form", clear_on_submit=True):
        question = st.text_area(
            "Tu consulta",
            placeholder="Ej.: ¿Dónde puedo ver mis días de vacaciones disponibles?",
            height=120,
        )

        send = st.form_submit_button(
            "Enviar consulta",
            use_container_width=True,
        )


# =========================================================
# PROCESAMIENTO DE LA CONSULTA
# =========================================================
if send:
    question = question.strip()

    if not question:
        st.warning("Escribí una consulta antes de enviarla.")

    else:
        st.session_state.messages.append(
            {"role": "user", "content": question}
        )

        normalized_question = normalize(question)

        st.session_state.question_counts[normalized_question] = (
            st.session_state.question_counts.get(normalized_question, 0) + 1
        )

        direct_alert = detect_alert(question, alertas)
        repeated = st.session_state.question_counts[normalized_question] >= 3
        human_state = detect_human_state(question)

        direct_contact_answer = direct_people_care_contact(question)
        search_query = get_contextual_query(question)
        faq, score, second_score = find_faq(search_query, faqs)

        # Evita responder una FAQ si la coincidencia es débil o ambigua.
        confident_match = (
            faq is not None
            and score >= 72
            and (score - second_score >= 8 or score >= 88)
        )

        if is_bot_complaint(question):
            answer = humanized_bot_complaint_response()

            send_alert(
                "[People Care] Disconformidad con el bot",
                alert_body(
                    question,
                    answer,
                    "RED_FLAG",
                    "Disconformidad / experiencia del colaborador",
                ),
            )

        elif human_state == "frustration":
            answer = humanized_frustration_response()

            send_alert(
                "[People Care] Colaborador requiere atención personalizada",
                alert_body(
                    question,
                    answer,
                    "RED_FLAG",
                    "Demora / frustración",
                ),
            )

        elif direct_contact_answer:
            answer = direct_contact_answer

        elif direct_alert:
            answer = (
                "Esta consulta requiere atención personalizada de People Care. "
                "Voy a dejarla señalada para que el equipo pueda revisarla."
            )

            send_alert(
                f"[People Care] Alerta {direct_alert['prioridad']} - "
                f"{direct_alert['tipo']}",
                alert_body(
                    question,
                    answer,
                    "RED_FLAG",
                    direct_alert["tipo"],
                ),
            )

        elif repeated:
            name = st.session_state.get("nombre", "").strip()
            prefix = f"{name}, " if name else ""
            answer = (
                f"{prefix}veo que este tema sigue sin quedar resuelto. "
                "No quiero que tengas que seguir repitiendo la misma consulta, "
                "así que voy a dejarla marcada para atención personalizada de People Care. "
                "También podés escribir a ppc@voolkia.com."
            )

            send_alert(
                "[People Care] Consulta reiterada",
                alert_body(
                    question,
                    answer,
                    "REITERADA",
                    "Consulta reiterada",
                ),
            )

        elif confident_match:
            answer = str(faq["Respuesta"])

        else:
            # Si hay una coincidencia débil/ambigua, primero pide aclaración.
            # Así evita "delirar" con una FAQ que no corresponde.
            if faq is not None and score >= 55:
                answer = humanized_clarification()
            else:
                answer = humanized_no_answer(question)

                send_alert(
                    "[People Care] Nueva consulta sin respuesta",
                    alert_body(
                        question,
                        answer,
                        "SIN_RESPUESTA",
                        "Sin respuesta",
                    ),
                )

        st.session_state.messages.append(
            {"role": "assistant", "content": answer}
        )

        st.rerun()


st.markdown(
    '<div class="pc-footer">'
    'Este asistente brinda información general basada en la base autorizada '
    'de People Care. No reemplaza la atención personalizada de HR.'
    '</div>',
    unsafe_allow_html=True,
)
