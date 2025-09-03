from babel.dates import format_date, format_datetime
from datetime import date, datetime
import pandas as pd

# -------------------------------------------------------------------
# Función segura para formatear fechas en español con Babel
# -------------------------------------------------------------------
def formatear_fecha_es(fecha, con_hora=False):
    """
    Formatea una fecha en español de forma segura usando Babel.
    - fecha: puede ser datetime, date, string o None.
    - con_hora: si True, usa formato con hora.
    """
    if not fecha or pd.isna(fecha):
        return "Fecha desconocida"

    try:
        if not isinstance(fecha, (date, datetime)):
            fecha = pd.to_datetime(fecha, errors="coerce")
        if pd.isna(fecha):
            return "Fecha desconocida"

        if con_hora:
            return format_datetime(fecha, format="long", locale="es")
        else:
            return format_date(fecha, format="long", locale="es")
    except Exception:
        return str(fecha)

# -------------------------------------------------------------------
# Sustituciones en el código original:
# -------------------------------------------------------------------
# 1. Se eliminaron todos los bloques con locale.setlocale(...)
# 2. Todas las llamadas a fecha.strftime(...) se reemplazaron por formatear_fecha_es(fecha)
# 3. En informes, eventos, retos y proyectos ahora usan formatear_fecha_es
# -------------------------------------------------------------------

# Ejemplo en generar_informe_persona:
# fecha_evento = evento_info.get('Fecha', 'N/D')
# fecha_str = formatear_fecha_es(fecha_evento)

# Ejemplo en eventos asistidos:
# fecha_str = formatear_fecha_es(evento.get("Fecha"))

# Ejemplo en retos tecnológicos:
# fecha_cierre_str = formatear_fecha_es(reto_info.get("Fecha cierre"))
# doc.add_paragraph(f"Fecha de cierre: {fecha_cierre_str}")

# Ejemplo en retos pasados:
# fecha_envio_str = formatear_fecha_es(reto_info.get("Fecha envío"))
# doc.add_paragraph(f"Fecha de envío: {fecha_envio_str}")

# Ejemplo en proyectos:
# fecha_final_str = formatear_fecha_es(fila.get(" Final"))
# doc.add_paragraph(f"Finalización del proyecto: {fecha_final_str}")

# -------------------------------------------------------------------
# Nota: Con estos cambios, cualquier fecha en eventos, retos o proyectos
# aparecerá en formato largo en español, ej: "3 de septiembre de 2025".
# -------------------------------------------------------------------
