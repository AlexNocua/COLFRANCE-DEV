import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import re  # para validar correos con regex

# ============================
# CONFIGURACIÓN GMAIL
# ============================
SMTP_SERVER = os.getenv("SMPT_SERVER_GMAIL")
SMTP_PORT = os.getenv("SMPT_SERVER_GMAIL_PORT")
EMAIL = os.getenv("EMAIL_USER_GMAIL")# tu correo de Gmail
PASSWORD = os.getenv("PASSWORD_APP_GMAIL")# contraseña de aplicación de Google

# ============================
# FUNCIÓN PARA VERIFICAR CORREOS
# ============================

def verificar_correos(ruta_excel):
    try:
        df = pd.read_excel(ruta_excel)

        if "Email Personal" not in df.columns:
            raise ValueError("❌ No se encontró la columna 'correo personal' en el Excel.")

        # Obtener lista de correos sin vacíos y únicos
        lista_correos = df["Email Personal"].dropna().unique().tolist()

        # Expresión regular simple para validar emails
        patron = r"^[\w\.-]+@[\w\.-]+\.\w+$"

        # Filtrar correos válidos
        correos_validos = [c for c in lista_correos if re.match(patron, str(c))]
        correos_invalidos = [c for c in lista_correos if not re.match(patron, str(c))]
        
        for correo in correos_validos:
            print(correo)
            
        print(f"✅ Se encontraron {len(correos_validos)} correos válidos.")
        if correos_invalidos:
            print(f"⚠️ Correos inválidos detectados: {correos_invalidos}")

        return correos_validos

    except Exception as e:
        print(f"❌ Error al leer correos: {e}")
        return []

# # ============================
# # FUNCIÓN PARA ENVIAR CORREO
# # ============================
def enviar_correo(destinatario, asunto, cuerpo, adjuntos):
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL
        msg["To"] = destinatario
        msg["Subject"] = asunto
        
        # esto es para edicion de texto plano
        # msg.attach(MIMEText(cuerpo, "plain"))
        
        # esto es para estructuras de mensajes HTML
        msg.attach(MIMEText(cuerpo, "html"))

        for i, file in enumerate(adjuntos):
            if os.path.exists(file):
                with open(file, "rb") as f:
                    mime = MIMEBase("application", "pdf")
                    mime.set_payload(f.read())
                    encoders.encode_base64(mime)

                    # Definir nombre personalizado solo para el segundo archivo
                    if i == 1:
                        filename = "PC-TH-10 Procedimiento de Gestión de Quejas de Colaboradores.pdf"
                    else:
                        filename = os.path.basename(file)

                    # 👇 Aquí se agregan las 2 cabeceras para evitar "noname"
                    mime.add_header("Content-Type", "application/pdf", name=filename)
                    mime.add_header("Content-Disposition", "attachment", filename=filename)

                    msg.attach(mime)



        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.send_message(msg)

        print(f"✅ Correo enviado a {destinatario}")
    except Exception as e:
        print(f"❌ Error al enviar a {destinatario}: {e}")

# ============================
# MAIN
# ============================
if __name__ == "__main__":
    # Verificar correos antes de enviar
    lista_correos = verificar_correos("./docs/basecorreos.xlsx")

    # Archivos adjuntos
    adjuntos = [
        "./docs/FR-TH-10-01 Formato de Queja de colaboradores.pdf",
        "./docs/PC-TH10ProcedimientodeGestióndeQuejasdeColaboradores.pdf"
    ]

    asunto = "Procedimiento de Quejas para Colaboradores"
#     cuerpo = """
# <h1>
# Estimados colaboradores,
# </h1

# Desde la Dirección de Talento Humano de Alimentos Colfrance S.A.S., nos complace informarles que hemos implementado de manera oficial el nuevo Procedimiento de Quejas para Colaboradores, con el objetivo de ofrecer un canal claro, confidencial y seguro para expresar cualquier situación que afecte su bienestar laboral.

# Este procedimiento busca garantizar:

# ✅ Escucha activa y sin prejuicios.
# ✅ Trato respetuoso y confidencial durante todo el proceso.
# ✅ Prevención de cualquier tipo de represalia contra quien presenta una queja.
# ✅ Seguimiento objetivo y justo de cada caso.

# Hemos diseñado herramientas prácticas para acompañar este proceso, como el Formato de Quejas, el Checklist de Bienestar Post-Queja y una ruta clara de atención que será liderada por el equipo de Talento Humano y el Comité de Convivencia Laboral.

# 📬 Recuerda que puedes presentar tus quejas o inquietudes a través de:

# Correo: talentohumano@colfrance.com

# Oficinas de Talento Humano (extensiones 2005)

# Buzón físico de sugerencias ubicado en la recepción de la empresa

# Si tienes dudas o deseas conocer más a fondo el proceso, no dudes en acercarte a nuestro equipo.

# Tu bienestar es clave para que sigamos creciendo juntos. ¡Tu voz cuenta!

# 🔗 Puedes consultar el procedimiento completo y descargar los formatos aquí:

# Atentamente,


# Cordialmente, 
# Talento Humano. 
# ALIMENTOS COLFRANCE
# Capellanía - Cundinamarca. Km. 2 Vía Chiquinquirá. 
# PBX: 7945760 +(57) 3175007752
# Tecnología Francesa -www.colfrance.com.co
# """

    cuerpo = """
<div style="font-family:Tahoma, Arial, sans-serif; font-size:15px; color:#073763; line-height:1.6;">
  
  <h2 style="color:#073763; text-align:left; margin-bottom:20px;">
    Estimados colaboradores,
  </h2>

  <p>
    Desde la Dirección de <strong>Talento Humano</strong> de <strong>Alimentos Colfrance S.A.S.</strong>, 
    en colaboración con el área de <strong>Sistemas</strong>, 
    nos complace informarles que hemos implementado de manera oficial el nuevo 
    <strong>Procedimiento de Quejas para Colaboradores</strong>, con el objetivo de ofrecer 
    un canal claro, confidencial y seguro para expresar cualquier situación que afecte 
    su bienestar laboral.
  </p>

  <p><strong>Este procedimiento busca garantizar:</strong></p>

  <ul style="margin:15px 0; padding-left:20px;">
    <li><strong>✅ Escucha activa y sin prejuicios.</strong></li>
    <li><strong>✅ Trato respetuoso y confidencial durante todo el proceso.</strong></li>
    <li><strong>✅ Prevención de cualquier tipo de represalia contra quien presenta una queja.</strong></li>
    <li><strong>✅ Seguimiento objetivo y justo de cada caso.</strong></li>
  </ul>

  <p>
    Hemos diseñado herramientas prácticas para acompañar este proceso, como el 
    <strong>Formato de Quejas</strong>, el <strong>Checklist de Bienestar Post-Queja</strong> y una 
    ruta clara de atención que será liderada por el equipo de <strong>Talento Humano</strong> y 
    el <strong>Comité de Convivencia Laboral</strong>.
  </p>

  <p><strong>📬 Recuerda que puedes presentar tus quejas o inquietudes a través de:</strong></p>

  <ul style="margin:15px 0; padding-left:20px;">
    <li><strong>Correo:</strong> <a href="mailto:talentohumano@colfrance.com" style="color:#073763; text-decoration:none;">talentohumano@colfrance.com</a></li>
    <li><strong>Oficinas de Talento Humano</strong> (extensiones 2005)</li>
    <li><strong>Buzón físico</strong> de sugerencias ubicado en la recepción de la empresa</li>
  </ul>

  <p>
    Si tienes dudas o deseas conocer más a fondo el proceso, no dudes en acercarte a nuestro equipo.
  </p>

  <p style="font-weight:bold; color:#000000;">
    Tu bienestar es clave para que sigamos creciendo juntos. ¡Tu voz cuenta!
  </p>

  <p>
    🔗 Puedes consultar el procedimiento completo y descargar los formatos aquí:
  </p>

  <br>

  <p><strong>Atentamente,</strong></p>
  <p>
    <strong>Area Sistemas y Talento Humano</strong><br>
    ALIMENTOS COLFRANCE<br>
    Capellanía - Cundinamarca. Km. 2 Vía Chiquinquirá.<br>
    PBX: 7945760 • +(57) 3175007752<br>
    Tecnología Francesa — 
    <a href="https://www.colfrance.com.co" style="color:#073763; text-decoration:none;">www.colfrance.com.co</a>
  </p>

  <p style="font-size:13px; color:#555; margin-top:20px;">
    <em>Nota: Este comunicado fue enviado en colaboración con el área de <strong>Sistemas</strong> y 
    <strong>Talento Humano</strong>.</em>
  </p>
</div>
"""



    # Enviar solo si hay correos válidos
    for correo in lista_correos:
        enviar_correo(correo, asunto, cuerpo, adjuntos)
