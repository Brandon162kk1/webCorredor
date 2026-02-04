import pytz
import calendar

from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta

# Definir la zona horaria de Lima
tz_peru = pytz.timezone("America/Lima")

def get_fecha_hoy():
    return datetime.now(tz_peru)

def get_timestamp():
    return datetime.now(tz_peru).strftime('%Y%m%d_%H%M%S')

def get_hora_minuto_segundo():
    return datetime.now(tz_peru).strftime("%H-%M-%S")

def get_fecha_actual():
    return datetime.now(tz_peru).strftime("%Y-%m-%d")

def get_fecha_dmy():
    return datetime.now(tz_peru).strftime("%d-%m-%Y")

def get_anio():
    return datetime.now(tz_peru).strftime("%Y")

def get_dia():
    return datetime.now(tz_peru).strftime("%d")

def get_mes():
    return datetime.now(tz_peru).strftime("%m")

def get_hora():
    return datetime.now(tz_peru).strftime("%H")

def get_minuto():
    return datetime.now(tz_peru).strftime("%M")

def get_pos_fecha_dmy():
    return datetime.now(tz_peru).strftime("%d/%m/%Y")

def sumar_x_dias_habiles(fecha_inicio, dias_habiles):
    fecha = fecha_inicio
    contador = 0
    while contador < dias_habiles:
        fecha += timedelta(days=1)
        if fecha.weekday() < 5:  # 0=Lunes, ..., 4=Viernes
            contador += 1
    return fecha

def tipo_vigencia(fecha_ini_inc, fecha_fin_inc):
    formato = "%d/%m/%Y"

    ini = datetime.strptime(fecha_ini_inc, formato)
    fin = datetime.strptime(fecha_fin_inc, formato)

    diff = relativedelta(fin, ini)
    total_meses = diff.years * 12 + diff.months
    dias = diff.days

    # ----------- NUEVA REGLA ESPECIAL -----------
    # Si inicia el día 1 y termina el ultimo día del mes → mensual
    ultimo_dia_mes_fin = calendar.monthrange(fin.year, fin.month)[1]
    if ini.day == 1 and fin.day == ultimo_dia_mes_fin:
        return "Mensual"
    # --------------------------------------------

    # Si los dias extra son ≤ 7 → ignorar (normal en polizas)
    if dias <= 7:
        if total_meses == 1:
            return "Mensual"
        elif total_meses == 2:
            return "Bimestral"
        elif total_meses == 3:
            return "Trimestral"
        elif total_meses == 6:
            return "Semestral"
        elif total_meses == 12:
            return "Anual"
        else:
            return f"Vigencia fuera de rango estándar ({total_meses} meses {dias} días)"

    # Si supera 7 días → muy irregular
    return f"Vigencia atípica ({total_meses} meses y {dias} días)"