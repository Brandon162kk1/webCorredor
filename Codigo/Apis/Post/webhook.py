import requests
import logging
import os
#from textwrap import dedent
from LinuxDebian.Carpetas.rutas import obtener_imagenes_error
from jinja2 import Environment, FileSystemLoader

# --- Variables de Entorno ---
url_n8n_enviar_correo_general = os.getenv("url_n8n_enviar_correo_general")

para_venv = os.getenv("para_wc")
para_lista = para_venv.split(",") if para_venv else []
copia_venv = os.getenv("copia_wc")
copias_lista = copia_venv.split(",") if copia_venv else []

ruta_plantilla = "/app/Codigo/Plantillas/Correo"
env = Environment(loader=FileSystemLoader(ruta_plantilla))

def enviar_error_general(cliente,ctx_ramo,palabra_clave,error,detalle_error,ruta_carpeta,const):

    template = env.get_template("error.html")

    imagenes = obtener_imagenes_error(ruta_carpeta, const)

    lista_tramas = []

    if ctx_ramo.trama:
        lista_tramas.append({
            "nombre": f"{ctx_ramo.poliza}.xlsx",
            "url": ctx_ramo.trama
        })

    if ctx_ramo.trama_97:
        lista_tramas.append({
            "nombre": f"{ctx_ramo.poliza}_97.xlsx",
            "url": ctx_ramo.trama_97
        })

    #tramas = []

    # if ctx_ramo.trama:
    #     tramas.append(ctx_ramo.trama)

    # if ctx_ramo.trama_97:
    #     tramas.append(ctx_ramo.trama_97)

    # mensaje = dedent(f"""Hubo problemas en la {palabra_clave} con la póliza {ctx_ramo.poliza} del cliente {cliente}.

    #     Datos:

    #     Vigencia: {ctx_ramo.f_inicio} al {ctx_ramo.f_fin}
    #     Tramas: {' y '.join(tramas) if tramas else 'No disponibles'}
    #     Sede: {ctx_ramo.sede}
    #     Compañía: {ctx_ramo.compania.capitalize()}

    #     Error Técnico y evidencia visual:

    #     Tipo de Error: {error}
    #     Detalle del Error:
    #     {str(detalle_error)}
    # """)

    # # 💥 LIMPIEZA FINAL (clave)
    # mensaje = "\n".join(line.strip() for line in mensaje.splitlines())

    html = template.render(
        titulo=f"⚠️ Problemas en la {palabra_clave}",
        cliente=cliente,
        poliza=ctx_ramo.poliza,
        rhumano=const.capitalize(),
        error=error,
        detalle_error=str(detalle_error),
        compania=ctx_ramo.compania.capitalize(),
        sede=ctx_ramo.sede,
        vigencia=f"{ctx_ramo.f_inicio} al {ctx_ramo.f_fin}",
        tramas=lista_tramas, #" y ".join(tramas) if tramas else "No disponibles",
        screenshot = (
            f"data:image/png;base64,{imagenes[0]}"
            if imagenes else None
        )
    )

    payload = {
        "Para": para_lista,
        "Copia": copias_lista,
        "Asunto": f"Error en la {palabra_clave} - Póliza: {ctx_ramo.poliza}",
        "Mensaje": html
        #"imagen_nombre": f"Error_{const}.png",
        #"imagen_base64": imagenes[0] if imagenes else None
    }

    try:
        response = requests.post(url_n8n_enviar_correo_general,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Notificación enviada al equipo Jishu")
        else:
            logging.error(f"❌ Problemas en el envio de notificación al equipo Jishu - {response.status_code} - {response.text}")

    except Exception as e:
        logging.error(f"❌ Error enviando la notificación por el webhook, Motivo : {e}")
