import datetime
import pandas as pd

PUESTOS = [
    {
        "nombre": "puesto_1",
        "nocturno": False,
    },
    {
        "nombre": "puesto_2",
        "nocturno": False,
    },
    {
        "nombre": "puesto_3 (noche)",
        "nocturno": True,
    },
    {
        "nombre": "puesto_4",
        "nocturno": False,
    }
]

EMPLEADOS = [{"nombre": f"Empleado_{i}"} for i in range(5)]

def crear_clase_empleado (e):
    return {
        "nombre": e['nombre'],
        "turnos_dia": 0,
        "turnos_noche": 0,
        "descansos": 0,
        'bloqueos_dia': e['bloqueos_dia'],
        'bloqueos_noche': e['bloqueos_noche'],
    }

def cronograma_diario_vacio (dia, puestos):
    cr_dia = {"fecha": dia}
    for p in puestos:
        cr_dia[p["nombre"]]= None
    return cr_dia

def actualizar_empleados (empleados, nombre, nocturno):
    actualizados = []
    for empleado in empleados:
        if empleado["nombre"] == nombre:
            objeto_modificado = empleado.copy()
            if nocturno:
                objeto_modificado["turnos_noche"] += 1
            else:
                objeto_modificado["turnos_dia"] += 1
            actualizados.append(objeto_modificado)
        else:
            actualizados.append(empleado)
    return actualizados

def actualizar_descansos (empleados, dia, dia_anterior):
    actualizados = []
    empleados_dia = list(dia.values())
    empleados_dia_anterior = list(dia_anterior.values())
    
    for empleado in empleados:
        if empleado["nombre"] not in empleados_dia and empleado["nombre"] not in empleados_dia_anterior:
            objeto_modificado = empleado.copy()
            objeto_modificado["descansos"] += 1
            actualizados.append(objeto_modificado)
        else:
            actualizados.append(empleado)
    return actualizados

def formatear_fecha (fecha):
    return fecha.date()

if __name__ == "__main__":
    # Carga de datos
    pages = xls = pd.ExcelFile('Entradas.xlsm').sheet_names
    df_configs = pd.read_excel('Entradas.xlsm', sheet_name='Configs', header=None, names=['A', 'B'])
    fecha_inicio = formatear_fecha(df_configs.at[0, 'B'])
    fecha_fin = formatear_fecha(df_configs.at[1, 'B'])

    PUESTOS = []
    for index, row in df_configs.iterrows():
        if (index > 3):
            PUESTOS.append({
                'nombre': row['A'],
                'nocturno': row['B']
            })
    pages.remove('Configs')
    pages.remove('Empleados')
    EMPLEADOS = []
    for empleado in pages:
        bloqueos_dia = []
        bloqueos_noche = []
        df_empleado = pd.read_excel('Entradas.xlsm', sheet_name=empleado, names=['fecha', 'dia', 'noche'])
        for index, row in df_empleado.iterrows():
            if (not pd.isna(row['dia'])):
                bloqueos_dia.append(formatear_fecha(row['fecha']))
            if (not pd.isna(row['noche'])):
                bloqueos_noche.append(formatear_fecha(row['fecha']))
        EMPLEADOS.append({
            'nombre': empleado,
            'bloqueos_dia': bloqueos_dia,
            'bloqueos_noche':bloqueos_noche,
        })
    #Cuadro de turnos
    
    delta = fecha_fin - fecha_inicio
    dias = [fecha_inicio + datetime.timedelta(days=i) for i in range(delta.days + 1)]
    empleados = [crear_clase_empleado(e) for e in EMPLEADOS]
    cronograma = [cronograma_diario_vacio(dia, PUESTOS) for dia in dias]

    for i in range(len(cronograma)):
        dia = cronograma[i]
        trabajadores_bloqueados = []
        bloqueos_descanso = []
        bloqueos_dia = []
        bloqueos_noche = []
        # Bloquea segun dias libres
        for empleado in empleados:
            dias_trabajados = empleado["turnos_noche"] + empleado["turnos_dia"]
            if dias_trabajados // 3 > empleado["descansos"]:
                bloqueos_descanso.append(empleado["nombre"])
            if dia['fecha'] in empleado['bloqueos_dia']:
                bloqueos_dia.append(empleado["nombre"])
            if dia['fecha'] in empleado['bloqueos_noche']:
                bloqueos_noche.append(empleado["nombre"])        
        # Bloquear empleados que trabajaron la noche anterior 
        if i > 0:
            dia_anterior = cronograma[i - 1]
            puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
            for p in puestos_nocturnos:
                trabajadores_bloqueados.append(dia_anterior[p["nombre"]])
        
        # Asignar puestos
        for puesto in PUESTOS:
            es_nocturno = puesto["nocturno"]
            
            #Bloqueo por asignacion dia actual o noche anterior
            if es_nocturno:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_descanso + bloqueos_noche]
            else:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_descanso + bloqueos_dia]
            
            # Si no encuentra un trabajador ideal, sacrifica condcion dia libre
            if len(empleados_habilitados) == 0:
                if es_nocturno:
                    empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_noche]
                else:
                    empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_dia]
                
            # Solo si hay trabajador disponible
            if len(empleados_habilitados) > 0:
                # Encuentra el trabajador disponible que menos ha trabajado
                empleado_a_asignar = min(empleados_habilitados, key=lambda item: item["turnos_noche"]) if es_nocturno else min(empleados_habilitados, key=lambda item: item["turnos_dia"])
                # Bloquea el trabajador elegido para no asignarlo a otro puesto el mismo dia
                trabajadores_bloqueados.append(empleado_a_asignar["nombre"])
                # Actualiza la cuenta de jornadas
                empleados = actualizar_empleados(empleados, empleado_a_asignar["nombre"], es_nocturno)
                #Actualiza el calendario final
                dia[puesto["nombre"]] = empleado_a_asignar["nombre"]
        cronograma[i] = dia
        
        if i > 0:
            # Actualizar los dias de descanso
            empleados = actualizar_descansos(empleados, dia, cronograma[i -1])

    data_excel = {
        'Puestos': [puesto['nombre'] for puesto in PUESTOS],
    }
    for dia in cronograma:
        data_excel[dia['fecha']] = [dia[puesto["nombre"]] for puesto in PUESTOS]
    df = pd.DataFrame(data_excel)
    df.to_excel(f"Cronograma-{cronograma[0]['fecha']}-{cronograma[-1]['fecha']}.xlsx", sheet_name='Cronograma', index=False)
    # for dia in cronograma:
    #     print(f"\n{dia["fecha"]}:")
    #     for puesto in PUESTOS:
    #         print(f"  - {puesto["nombre"]}: {dia[puesto["nombre"]]}")
