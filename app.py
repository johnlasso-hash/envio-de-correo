import os
import shutil
import zipfile
import smtplib
import re
import sqlite3
import time
from datetime import datetime, timedelta
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import gradio as gr

# ==========================================
# CONFIGURACIÓN Y BASE DE DATOS LOCAL
# ==========================================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
DB_HISTORIAL = "historial_envios.sqlite"

def inicializar_bd():
    """Crea la tabla de historial si no existe."""
    conn = sqlite3.connect(DB_HISTORIAL)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS envios (
            correo TEXT PRIMARY KEY,
            contribuyente TEXT,
            fecha_envio TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def correo_enviado_recientemente(correo, horas=24):
    """Verifica si se le ha enviado un correo al destinatario en las últimas X horas."""
    conn = sqlite3.connect(DB_HISTORIAL)
    cursor = conn.cursor()
    cursor.execute('SELECT fecha_envio FROM envios WHERE correo = ?', (correo.lower(),))
    row = cursor.fetchone()
    conn.close()

    if row:
        fecha_ultimo_envio = datetime.fromisoformat(row[0])
        if datetime.now() - fecha_ultimo_envio < timedelta(horas=horas):
            return True, fecha_ultimo_envio.strftime("%Y-%m-%d %H:%M:%S")
    return False, None

def registrar_envio_exitoso(correo, contribuyente):
    """Registra o actualiza el envío exitoso en la base de datos."""
    conn = sqlite3.connect(DB_HISTORIAL)
    cursor = conn.cursor()
    now_str = datetime.now().isoformat()
    cursor.execute('''
        INSERT OR REPLACE INTO envios (correo, contribuyente, fecha_envio)
        VALUES (?, ?, ?)
    ''', (correo.lower(), contribuyente, now_str))
    conn.commit()
    conn.close()

def es_correo_valido(correo):
    patron = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return bool(re.match(patron, str(correo).strip()))

inicializar_bd()

# ==========================================
# LÓGICA PRINCIPAL DE PROCESAMIENTO
# ==========================================
def procesar_y_enviar(correo_emisor, clave_app, archivo_excel, archivo_zip, asunto_input, cuerpo_input, progreso=gr.Progress()):
    if not correo_emisor or not clave_app:
        return "❌ **Error:** Debe ingresar el correo emisor y la contraseña de aplicación.", None

    if not es_correo_valido(correo_emisor):
        return f"❌ **Error:** El correo emisor (`{correo_emisor}`) no tiene un formato válido.", None

    if archivo_excel is None or archivo_zip is None:
        return "❌ **Error:** Debe cargar la plantilla Excel y el archivo ZIP.", None

    if not asunto_input or not cuerpo_input:
        return "⚠️ **Error:** Debe ingresar el asunto y el cuerpo del correo.", None

    temp_zip_dir = "./archivos_extraidos"
    if os.path.exists(temp_zip_dir):
        shutil.rmtree(temp_zip_dir)
    os.makedirs(temp_zip_dir)

    try:
        ruta_zip = archivo_zip if isinstance(archivo_zip, str) else archivo_zip.name
        with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
            zip_ref.extractall(temp_zip_dir)
    except Exception as e:
        return f"❌ **Error al descomprimir el archivo ZIP:** {str(e)}", None

    try:
        ruta_excel = archivo_excel if isinstance(archivo_excel, str) else archivo_excel.name
        df = pd.read_excel(ruta_excel)
        df.columns = df.columns.str.strip().str.lower()
    except Exception as e:
        return f"❌ **Error al leer el Excel:** {str(e)}", None

    total_correos = len(df)
    entregados = 0
    fallidos = 0
    omitidos_sin_adjunto = 0
    omitidos_24h = 0
    detalles_log = []
    registros_auditoria = []

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(correo_emisor.strip(), clave_app.strip())
    except Exception as e:
        return f"❌ **Error de autenticación SMTP:** Verifique sus credenciales.\n\n*Detalle:* {str(e)}", None

    for index, row in progreso.tqdm(df.iterrows(), total=total_correos, desc="Enviando Correos"):
        correo_destino = str(row.get('correo electronico', '')).strip().lower()
        nombre_contribuyente = str(row.get('contribuyente', 'Contribuyente')).strip()
        base_nombre_archivo = str(row.get('nombre archivo', '')).strip()

        if not es_correo_valido(correo_destino):
            fallidos += 1
            msg_log = f"Fila {index+2} ({nombre_contribuyente}): ❌ Correo no válido ({correo_destino})."
            detalles_log.append(msg_log)
            registros_auditoria.append({"Fila": index+2, "Correo": correo_destino, "Estado": "Fallido", "Motivo": "Formato inválido"})
            continue

        ya_enviado, fecha_prev = correo_enviado_recientemente(correo_destino, horas=24)
        if ya_enviado:
            omitidos_24h += 1
            msg_log = f"Fila {index+2} ({correo_destino}): ⏭️ OMITIDO (Ya enviado previamente el {fecha_prev})."
            detalles_log.append(msg_log)
            registros_auditoria.append({"Fila": index+2, "Correo": correo_destino, "Estado": "Omitido", "Motivo": f"Enviado hace <24h ({fecha_prev})"})
            continue

        posibles_nombres = [base_nombre_archivo, f"{base_nombre_archivo}.pdf"]
        ruta_adjunto = None
        nombre_adjunto_final = ""

        for nombre in posibles_nombres:
            ruta_test = os.path.join(temp_zip_dir, nombre)
            if os.path.exists(ruta_test) and os.path.isfile(ruta_test):
                ruta_adjunto = ruta_test
                nombre_adjunto_final = nombre
                break

        if not ruta_adjunto:
            omitidos_sin_adjunto += 1
            msg_log = f"Fila {index+2} ({correo_destino}): ⚠️ OMITIDO (Adjunto '{base_nombre_archivo}' no encontrado en el ZIP)."
            detalles_log.append(msg_log)
            registros_auditoria.append({"Fila": index+2, "Correo": correo_destino, "Estado": "Omitido", "Motivo": "Sin archivo adjunto"})
            continue

        cuerpo_personalizado = cuerpo_input.replace("{contribuyente}", nombre_contribuyente)
        msg = MIMEMultipart()
        msg['From'] = correo_emisor.strip()
        msg['To'] = correo_destino
        msg['Subject'] = asunto_input
        msg.attach(MIMEText(cuerpo_personalizado, 'plain'))

        try:
            with open(ruta_adjunto, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename= {nombre_adjunto_final}")
                msg.attach(part)
        except Exception as e:
            fallidos += 1
            detalles_log.append(f"Fila {index+2} ({correo_destino}): ❌ Error adjuntando archivo: {str(e)}")
            registros_auditoria.append({"Fila": index+2, "Correo": correo_destino, "Estado": "Fallido", "Motivo": f"Error adjunto: {str(e)}"})
            continue

        try:
            server.sendmail(correo_emisor.strip(), correo_destino, msg.as_string())
            entregados += 1
            registrar_envio_exitoso(correo_destino, nombre_contribuyente)
            detalles_log.append(f"Fila {index+2} ({correo_destino}): ✅ Enviado a {nombre_contribuyente}.")
            registros_auditoria.append({"Fila": index+2, "Correo": correo_destino, "Estado": "Exitoso", "Motivo": "Enviado correctamente"})
            time.sleep(0.5)
        except Exception as e:
            fallidos += 1
            detalles_log.append(f"Fila {index+2} ({correo_destino}): ❌ Error de envío: {str(e)}")
            registros_auditoria.append({"Fila": index+2, "Correo": correo_destino, "Estado": "Fallido", "Motivo": f"Error SMTP: {str(e)}"})

    server.quit()
    shutil.rmtree(temp_zip_dir)

    df_auditoria = pd.DataFrame(registros_auditoria)
    ruta_csv = f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    df_auditoria.to_csv(ruta_csv, index=False, encoding='utf-8-sig')

    logs_texto = "\n".join(detalles_log)
    resumen_informe = f"""### 📊 Informe de Gestión de Envíos - Municipio de Palmira

**Emisor Configurado:** `{correo_emisor.strip()}`  
**Fecha de Ejecución:** `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`

| Métrica | Cantidad |
| :--- | :---: |
| 📋 Total de Registros Procesados | **{total_correos}** |
| ✅ Envíos Entregados con Éxito | **{entregados}** |
| ⏭️ Omitidos por Falta de Adjunto | **{omitidos_sin_adjunto}** |
| 🛡️ Omitidos por Control 24h | **{omitidos_24h}** |
| ❌ Envíos Fallidos | **{fallidos}** |

---
#### 📄 Detalle Operativo de Envíos:
```text
{logs_texto}
```"""

    return resumen_informe, ruta_csv

def exportar_historial_completo():
    conn = sqlite3.connect(DB_HISTORIAL)
    df = pd.read_sql_query("SELECT * FROM envios", conn)
    conn.close()
    ruta_historial = "historial_completo_bd.csv"
    df.to_csv(ruta_historial, index=False, encoding='utf-8-sig')
    return ruta_historial

# ==========================================
# ESTILOS CSS PERSONALIZADOS (ALTO CONTRASTE)
# ==========================================
custom_css = """
:root {
    color-scheme: dark !important;
}

/* Ajustes generales de la aplicación */
body, .gradio-container {
    background-color: #12181b !important;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
    color: #f0f0f0 !important;
}

/* Encabezado Principal */
.header-bar {
    background-color: #007a53;
    color: white;
    padding: 18px 28px;
    border-radius: 8px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 20px;
}
.header-title h1 { color: #ffffff !important; font-size: 22px !important; margin: 0 !important; font-weight: bold; }
.header-title p { color: #e1f5fe !important; font-size: 14px !important; margin: 0 !important; }
.badge-app { background-color: rgba(255, 255, 255, 0.2); color: #ffffff; padding: 6px 14px; border-radius: 20px; font-size: 13px; font-weight: bold; }

/* Contadores y Pestañas */
button.tabnav-tab {
    color: #e0e0e0 !important;
    font-size: 15px !important;
    font-weight: 600 !important;
}
button.tabnav-tab.selected {
    color: #4caf50 !important;
    border-bottom-color: #4caf50 !important;
}

/* Títulos, etiquetas y textos generales */
label span, h1, h2, h3, h4, p, span, div, .gr-form {
    color: #f0f0f0 !important;
}

/* Cuadro de Reglas */
.info-rules-box {
    background-color: #1b2e2b !important;
    border: 1px solid #007a53 !important;
    border-radius: 10px;
    padding: 18px;
}
.info-rules-box h4 { color: #81c784 !important; margin-top: 0; }
.info-rules-box ul li { color: #e0e0e0 !important; margin-bottom: 6px; }

/* Botón Principal */
.btn-primary-palmira {
    background-color: #007a53 !important;
    color: #ffffff !important;
    border-radius: 8px !important;
    padding: 12px 20px !important;
    font-size: 16px !important;
    font-weight: bold !important;
}
"""

cuerpo_por_defecto = """Apreciado(a) contribuyente {contribuyente}:

La Secretaría de Hacienda Municipal le extiende un cordial saludo y expresa su sincero agradecimiento por mantenerse al día en el pago del Impuesto Predial Unificado."""

with gr.Blocks(theme=gr.themes.Soft(), css=custom_css, title="Alcaldía de Palmira - Notificaciones") as demo:

    gr.HTML("""
        <div class="header-bar">
            <div class="header-title">
                <h1>🏛️ Alcaldía de Palmira</h1>
                <p>Subsecreatría de Ingresos y Tesorería -Secretaría de Hacienda - Sistema de Notificaciones Masivas</p>
            </div>
            <div class="badge-app">By John Lasso</div>
        </div>
    """)

    with gr.Tabs():
        with gr.Tab("🚀 Panel de Envíos Masivos"):
            with gr.Row():
                with gr.Column(scale=12):
                    with gr.Accordion("🔑 Credenciales del Emisor", open=True):
                        with gr.Row():
                            txt_emisor = gr.Textbox(label="Correo Emisor", placeholder="ejemplo@palmira.gov.co", scale=1)
                            txt_clave = gr.Textbox(label="Contraseña de Aplicación", type="password", placeholder="xxxx xxxx xxxx xxxx", scale=1)
                        with gr.Accordion("❓ ¿Cómo obtener la Contraseña de Aplicación?", open=False):
                            gr.Markdown("1. Entra a tu Cuenta de Google ([myaccount.google.com](https://myaccount.google.com/)).\n2. Activa **Verificación en dos pasos**.\n3. Busca **Contraseñas de aplicaciones**, crea una clave y copia los 16 caracteres.")

                    gr.Markdown("### 📁 1. Carga de Insumos Locales")
                    with gr.Row():
                        file_excel = gr.File(label="Plantilla Excel (.xlsx, .xls)", file_types=[".xlsx", ".xls"], scale=1)
                        file_zip = gr.File(label="Archivo ZIP con PDFs (.zip)", file_types=[".zip"], scale=1)

                    gr.Markdown("### ✉️ 2. Configuración del Mensaje")
                    txt_asunto = gr.Textbox(label="Asunto del Correo", value="Notificación Estado de Cuenta IPU 2026")
                    txt_cuerpo = gr.Textbox(label="Cuerpo del Mensaje (Soporta variable {contribuyente})", lines=5, value=cuerpo_por_defecto)

                    btn_ejecutar = gr.Button("🚀 Procesar Archivos e Iniciar Envío", elem_classes="btn-primary-palmira")

                with gr.Column(scale=10):
                    gr.Markdown("### 📊 Reporte de Ejecución")
                    txt_reporte = gr.Markdown("*Los resultados del envío masivo y el informe de trazabilidad se mostrarán aquí en tiempo real...*")
                    file_descarga_reporte = gr.File(label="📥 Descargar Reporte en CSV", interactive=False)

                    gr.HTML("""
                        <div class="info-rules-box">
                            <h4>🛡️ Reglas del Sistema Activas:</h4>
                            <ul>
                                <li><b>Adjunto Obligatorio:</b> Si el archivo de la columna <i>Nombre Archivo</i> no está en el ZIP, el envío se cancela.</li>
                                <li><b>Protección 24h:</b> No se repiten envíos a correos notificados exitosamente en las últimas 24 horas.</li>
                                <li><b>Envíos Diarios:</b> Límite máximo de 1.500 correos por día según políticas de Google.</li>
                            </ul>
                        </div>
                    """)

        with gr.Tab("📜 Historial de Envíos y Base de Datos"):
            gr.Markdown("### 🛡️ Base de Datos Local Anti-Duplicados")
            gr.Markdown("Consulte el histórico completo de correos procesados anteriormente.")
            btn_descargar_bd = gr.Button("📥 Exportar Registro Histórico Completo (BD)", variant="secondary")
            file_bd_csv = gr.File(label="Historial Consolidado (CSV)")

            btn_descargar_bd.click(fn=exportar_historial_completo, outputs=[file_bd_csv])

    btn_ejecutar.click(
        fn=procesar_y_enviar,
        inputs=[txt_emisor, txt_clave, file_excel, file_zip, txt_asunto, txt_cuerpo],
        outputs=[txt_reporte, file_descarga_reporte]
    )

# Configuración de puerto para Render
port = int(os.environ.get("PORT", 7860))
demo.launch(server_name="0.0.0.0", server_port=port)
