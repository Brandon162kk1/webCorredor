import pytz
import calendar
import random
import time

from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta

# Definir la zona horaria de Lima
tz_peru = pytz.timezone("America/Lima")

#datetime.strptime(fecha, "%d/%m/%Y")

def time_espera_alea(min_seg,max_seg):
     return time.sleep(random.uniform(min_seg,max_seg))

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

def get_fecha_menos_x_dias(dia):
    fecha = datetime.now(tz_peru) - timedelta(days=dia)
    return fecha.strftime("%d/%m/%Y")

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

    if fin < ini:
        raise ValueError("La fecha fin no puede ser anterior a la fecha de inicio")

    diff = relativedelta(fin, ini)
    total_meses = diff.years * 12 + diff.months
    dias = diff.days

    # ----------- NUEVA REGLA ESPECIAL -----------
    # Si inicia el día 1 y termina el ultimo día del mes → mensual
    ultimo_dia_mes_fin = calendar.monthrange(fin.year, fin.month)[1]
    if (
        ini.year == fin.year
        and ini.month == fin.month
        and ini.day == 1
        and fin.day == ultimo_dia_mes_fin
    ):
        return "Mensual"
    # if ini.day == 1 and fin.day == ultimo_dia_mes_fin:
    #     return "Mensual"
    # --------------------------------------------

    # Ajustar vigencias que abarcan meses completos de calendario
    if ini.day == 1 and fin.day == ultimo_dia_mes_fin:
        meses_calendario = (
            (fin.year - ini.year) * 12
            + fin.month - ini.month + 1
        )

        if meses_calendario == 12:
            return "Anual"

    # Clasificación por duración
    if dias <= 7:
        modalidades = {
            1: "Mensual",
            2: "Bimestral",
            3: "Trimestral",
            6: "Semestral",
            12: "Anual",
        }

        if total_meses in modalidades:
            return modalidades[total_meses]

        return (
            f"Vigencia fuera de rango estándar "
            f"({total_meses} meses {dias} días)"
        )

    # Si supera 7 días → muy irregular
    return f"Vigencia atípica ({total_meses} meses y {dias} días)"
