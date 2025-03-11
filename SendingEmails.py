import pandas as pd
from email.message import EmailMessage
import ssl  # Por seguridad
import smtplib  # Para poder enviar el mensaje

# Cargar datos
df = pd.read_excel("DatosUsuarios.xlsx")
print(df.head())


# Limpiar y filtrar datos
df['Genero'] = df['Genero'].str.strip().str.lower()
fm_email = df[df["Genero"] == 'femenino']['Email'].tolist()
ml_email = df[df["Genero"] == 'masculino']['Email'].tolist()

print(fm_email)
print(ml_email)

# Credenciales de correo
email_sender = 'example@gmail.com'
password = 'PASSWORD'  # contraseña de aplicación

# Configuración SSL
context = ssl.create_default_context()

# Abrir la conexión SMTP dentro del bloque with
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, password)

    # Enviar correo a mujeres
    if fm_email:
        subject = '🌸 ¡Feliz Día de la Mujer! 🌸'
        body = """Querida amiga,

Hoy queremos celebrar la fuerza, valentía y dedicación de todas las mujeres. 
Gracias por tu inspiración y por hacer del mundo un lugar mejor. 💜

¡Feliz Día de la Mujer!"""

        msg = EmailMessage()
        msg['From'] = email_sender
        msg['To'] = ", ".join(fm_email)
        msg['Subject'] = subject
        msg.set_content(body)

        smtp.sendmail(email_sender, fm_email, msg.as_string())

    # Enviar correo a hombres
    if ml_email:
        subject = '🌟 ¡Feliz Día del Hombre! 🌟'
        body = """Querido amigo,

Hoy queremos celebrar la fuerza, valentía y dedicación de todos los hombres. 
Gracias por tu inspiración y por hacer del mundo un lugar mejor. 💜

¡Feliz Día del Hombre!"""

        msg = EmailMessage()
        msg['From'] = email_sender
        msg['To'] = ", ".join(ml_email)
        msg['Subject'] = subject
        msg.set_content(body)

        smtp.sendmail(email_sender, ml_email, msg.as_string())
