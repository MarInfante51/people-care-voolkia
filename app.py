
import io
import re
import json
import smtplib
import unicodedata
from datetime import datetime
from email.message import EmailMessage
from html import escape

import pandas as pd
import requests
import streamlit as st
from google import genai
from google.genai import types


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
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
GEMINI_MODEL = "gemini-3.5-flash-lite"


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
    
    .pc-access-grid {
        display: grid;
        grid-template-columns: 1.05fr .95fr;
        gap: 22px;
        margin-top: 8px;
    }

    .pc-access-card {
        background: #FFFFFF;
        border: 1px solid #E7DED9;
        border-radius: 24px;
        padding: 26px;
        box-shadow: 0 10px 28px rgba(36,3,0,.05);
    }

    .pc-access-soft {
        background: linear-gradient(180deg, #FFF8F3 0%, #FFFFFF 100%);
        border: 1px solid #F2D6C3;
        border-radius: 24px;
        padding: 26px;
        box-shadow: 0 10px 28px rgba(36,3,0,.04);
    }

    .pc-eyebrow {
        display: inline-block;
        background: #240300;
        color: #FFFFFF;
        border-radius: 999px;
        padding: 7px 12px;
        font-size: .74rem;
        font-weight: 800;
        letter-spacing: .04em;
        margin-bottom: 14px;
    }

    .pc-access-title {
        color: #240300;
        font-size: 1.55rem;
        font-weight: 800;
        line-height: 1.18;
        margin-bottom: 8px;
    }

    .pc-access-copy {
        color: #6E625D;
        font-size: .94rem;
        line-height: 1.55;
        margin-bottom: 16px;
    }

    .pc-mini-item {
        display: flex;
        gap: 10px;
        align-items: flex-start;
        margin-bottom: 12px;
        color: #443936;
        font-size: .9rem;
        line-height: 1.45;
    }

    .pc-mini-dot {
        width: 8px;
        height: 8px;
        min-width: 8px;
        border-radius: 50%;
        background: #FF7000;
        margin-top: 6px;
    }

    .pc-human-box {
        background: #240300;
        color: #FFFFFF;
        border-radius: 18px;
        padding: 17px 18px;
        margin-top: 18px;
    }

    .pc-human-title {
        font-size: .95rem;
        font-weight: 800;
        margin-bottom: 6px;
    }

    .pc-human-copy {
        font-size: .86rem;
        line-height: 1.5;
        color: #F7EEE9;
    }

    .pc-hours {
        background: #FFF0E4;
        border: 1px solid #FFD0AF;
        border-radius: 18px;
        padding: 16px 17px;
        margin-top: 14px;
        color: #5B4032;
    }

    .pc-hours strong {
        color: #240300;
    }

    .pc-safe {
        margin-top: 14px;
        color: #796D67;
        font-size: .78rem;
        line-height: 1.45;
    }

    @media (max-width: 800px) {
        .pc-access-grid {
            grid-template-columns: 1fr;
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



def build_knowledge_base(faqs):
    active = faqs[
        faqs["Activo"]
        .astype(str)
        .str.lower()
        .isin(["sí", "si", "true", "1"])
    ]

    blocks = []
    for _, row in active.iterrows():
        blocks.append(
            f"ID: {row.get('ID', '')}\n"
            f"CATEGORÍA: {row.get('Categoría', '')}\n"
            f"PREGUNTA: {row.get('Pregunta', '')}\n"
            f"RESPUESTA AUTORIZADA: {row.get('Respuesta', '')}\n"
            f"PALABRAS CLAVE: {row.get('Palabras_clave', '')}"
        )

    return "\n\n---\n\n".join(blocks)


def recent_conversation():
    history = []

    # Últimos 8 mensajes son suficientes para mantener el hilo
    # sin inflar innecesariamente el consumo de tokens.
    for msg in st.session_state.messages[-8:]:
        role = "COLABORADOR" if msg["role"] == "user" else "PEOPLE CARE"
        history.append(f"{role}: {msg['content']}")

    return "\n".join(history)


def ask_gemini(question, faqs):
    """
    Gemini interpreta lenguaje y contexto, pero la única fuente de verdad
    permitida son las FAQs activas del Google Sheet.
    """
    if not GEMINI_API_KEY:
        raise RuntimeError(
            "Falta GEMINI_API_KEY en los Secrets de Streamlit."
        )

    client = genai.Client(api_key=GEMINI_API_KEY)
    knowledge = build_knowledge_base(faqs)
    conversation = recent_conversation()

    system_instruction = """
Sos el asistente virtual de People Care de Voolkia.

OBJETIVO
Ayudar a colaboradores de Voolkia a resolver consultas de HR de manera rápida,
clara, cordial y natural.

REGLA MÁS IMPORTANTE
La única fuente de verdad es la BASE DE CONOCIMIENTO AUTORIZADA que se incluye
en cada solicitud. NO uses tu conocimiento general, NO completes información,
NO supongas políticas de Voolkia y NO inventes datos.

COMPORTAMIENTO
- Entendé lenguaje cotidiano, errores de tipeo, frases breves y preguntas de seguimiento.
- Conservá el contexto de la conversación. Ejemplo: si primero preguntan por prepaga
  y luego dicen "¿y cómo lo hago?", entendé que siguen hablando de prepaga.
- Si una consulta amplia coincide con varias FAQs, orientá al colaborador con opciones
  concretas tomadas de la base. Ejemplo: "vacaciones" puede abrir cómo pedirlas o dónde
  consultar días disponibles.
- Si la respuesta está respaldada por la base, respondé de inmediato en lenguaje natural.
- No copies mecánicamente la FAQ si podés expresarla de forma más conversacional,
  pero nunca cambies su sentido.
- Si la información NO está en la base, decilo claramente y no inventes.
- Si falta contexto para decidir entre dos respuestas posibles, hacé UNA pregunta breve
  de aclaración.
- Si el colaborador expresa frustración, demora, enojo o pide hablar con alguien,
  reconocé la situación con respeto y ofrecé People Care: ppc@voolkia.com.
- No seas defensivo, robótico ni excesivamente formal.
- Usá el nombre del colaborador con moderación; no hace falta repetirlo en cada mensaje.
- Nunca pidas datos médicos, bancarios, contraseñas ni documentación sensible.

ESTADOS
answered: pudiste responder usando la base.
clarify: necesitás una aclaración breve antes de responder.
escalate: la base no contiene la respuesta o requiere atención personalizada.

RED FLAG
Marcá red_flag=true cuando haya frustración relevante, reclamo, demora sin respuesta,
problemas individuales de haberes, renuncia, conflicto, situación sensible o pedido
explícito de atención humana.

Respondé SIEMPRE en JSON siguiendo el esquema solicitado.
"""

    prompt = f"""
COLABORADOR
Nombre: {st.session_state.nombre}
Apellido: {st.session_state.apellido}
Proyecto/Cliente: {st.session_state.proyecto}

BASE DE CONOCIMIENTO AUTORIZADA
{knowledge}

CONVERSACIÓN RECIENTE
{conversation}

NUEVA CONSULTA
{question}
"""

    schema = {
        "type": "object",
        "properties": {
            "answer": {
                "type": "string",
                "description": "Respuesta natural para mostrar al colaborador."
            },
            "status": {
                "type": "string",
                "enum": ["answered", "clarify", "escalate"]
            },
            "red_flag": {
                "type": "boolean"
            },
            "alert_reason": {
                "type": "string",
                "description": "Motivo breve de alerta; vacío si no corresponde."
            }
        },
        "required": ["answer", "status", "red_flag", "alert_reason"]
    }

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_instruction,
            response_mime_type="application/json",
            response_schema=schema,
            temperature=0.2,
            max_output_tokens=450,
        ),
    )

    if not response.text:
        raise RuntimeError("Gemini no devolvió una respuesta.")

    data = json.loads(response.text)

    return {
        "answer": str(data.get("answer", "")).strip(),
        "status": str(data.get("status", "escalate")).strip(),
        "red_flag": bool(data.get("red_flag", False)),
        "alert_reason": str(data.get("alert_reason", "")).strip(),
    }


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
    access_html = (
        '<div class="pc-access-grid">'
        '<div class="pc-access-card">'
        '<div class="pc-eyebrow">ASISTENTE PEOPLE CARE</div>'
        '<div class="pc-access-title">Resolvé consultas frecuentes de HR en pocos minutos</div>'
        '<div class="pc-access-copy">'
        'El asistente responde únicamente con información validada por People Care. '
        'Si no encuentra una respuesta segura, no inventa y deriva el tema para revisión.'
        '</div>'
        '<div class="pc-mini-item"><span class="pc-mini-dot"></span>'
        '<span>Vacaciones, licencias, recibos, beneficios, prepaga, referidos y capacitaciones.</span></div>'
        '<div class="pc-mini-item"><span class="pc-mini-dot"></span>'
        '<span>Consultas operativas frecuentes y orientación sobre circuitos internos.</span></div>'
        '<div class="pc-mini-item"><span class="pc-mini-dot"></span>'
        '<span>Respuestas inmediatas usando la base oficial de People Care.</span></div>'
        '<div class="pc-human-box">'
        '<div class="pc-human-title">¿Necesitás atención personalizada?</div>'
        '<div class="pc-human-copy">'
        'Para casos personales, sensibles o situaciones que requieren seguimiento, '
        'podés escribir directamente a <strong>ppc@voolkia.com</strong>.'
        '</div></div>'
        '<div class="pc-hours">'
        '<strong>HR Office Hours</strong><br>'
        'Todos los viernes de <strong>10 a 12 hs</strong> tenés un espacio abierto '
        'para conversar directamente con el equipo de HR.'
        '</div>'
        '<div class="pc-safe">'
        'No ingreses contraseñas, datos bancarios, documentación médica ni otra información sensible.'
        '</div>'
        '</div>'
        '<div class="pc-access-soft">'
        '<div class="pc-eyebrow">INGRESO</div>'
        '<div class="pc-access-title">Antes de comenzar</div>'
        '<div class="pc-access-copy">'
        'Identificate para que People Care pueda contextualizar tu consulta '
        'si necesitás atención personalizada.'
        '</div>'
        '</div>'
        '</div>'
    )

    st.markdown(access_html, unsafe_allow_html=True)

    st.write("")

    form_col, _ = st.columns([0.58, 0.42])

    with form_col:
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
                "Ingresar al asistente",
                use_container_width=True,
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

        # Gemini interpreta intención + contexto usando SOLO el Google Sheet.
        try:
            ai_result = ask_gemini(question, faqs)
            answer = ai_result["answer"]
            status = ai_result["status"]
            red_flag = ai_result["red_flag"]
            alert_reason = ai_result["alert_reason"]

            # Alerta cuando Gemini identifica una situación sensible o
            # cuando la base no permite resolver la consulta.
            if red_flag or status == "escalate":
                tipo = alert_reason or (
                    "Consulta sin respuesta validada"
                    if status == "escalate"
                    else "Atención personalizada"
                )

                send_alert(
                    f"[People Care] {tipo}",
                    alert_body(
                        question,
                        answer,
                        "RED_FLAG" if red_flag else "SIN_RESPUESTA",
                        tipo,
                    ),
                )

        except Exception as error:
            # Si Gemini falla, NO inventamos ni volvemos al fuzzy matching.
            answer = (
                "En este momento no pude procesar tu consulta correctamente. "
                "Prefiero no darte una respuesta que pueda ser incorrecta. "
                "Podés intentar nuevamente o escribir a ppc@voolkia.com."
            )

            send_alert(
                "[People Care] Error del asistente",
                alert_body(
                    question,
                    answer,
                    "ERROR_IA",
                    f"Error Gemini: {str(error)[:180]}",
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
