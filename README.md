# People Care Voolkia — MVP

## Qué hace
- Identifica al colaborador por nombre, apellido y proyecto/cliente.
- Lee FAQs autorizadas desde Excel.
- Busca la respuesta más cercana sin usar un LLM: evita inventar información.
- Si no encuentra una respuesta, la deriva a People Care.
- Detecta palabras/frases de red flag.
- Detecta la misma consulta repetida 3 veces en una sesión.
- Envía alertas por email a marcela@voolkia.com cuando SMTP está configurado.

## Archivos
- `app.py`: aplicación Streamlit.
- `People_Care_Voolkia_Base_Conocimiento.xlsx`: base editable.
- `requirements.txt`: dependencias.
- `.streamlit/config.toml`: tema.
- `.streamlit/secrets.example.toml`: ejemplo de secretos.

## Google Drive
1. Subir `People_Care_Voolkia_Base_Conocimiento.xlsx` a Drive.
2. Obtener el ID del archivo desde la URL.
3. Para este MVP, compartir el archivo como solo lectura mediante enlace.
4. En Streamlit Secrets cargar:
   `FAQ_DRIVE_FILE_ID = "ID_DEL_ARCHIVO"`

La app vuelve a leer el Excel periódicamente (cache de 5 minutos), por lo que las nuevas FAQs aparecen sin modificar el código.

## Email
Configurar SMTP en Streamlit Secrets. Nunca subir contraseñas a GitHub.

## Publicación
Subir estos archivos a un repositorio GitHub nuevo y desplegar `app.py` en Streamlit Community Cloud.

## Importante
El formulario de nombre/apellido/proyecto IDENTIFICA, pero no autentica.
Para producción corporativa con información interna o casos personales se recomienda SSO/login corporativo o despliegue privado.
No guardar en la base documentación, diagnósticos, contraseñas ni información personal sensible.
