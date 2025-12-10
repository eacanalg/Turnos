import datetime
import pandas as pd

def get_excel_column_name(column_number):
    column_name = ""
    while column_number > 0:
        # Adjust for 0-based remainder (A=1 becomes 0, B=2 becomes 1, etc.)
        column_number -= 1
        # Find the remainder to determine the current character
        remainder = column_number % 26
        # Convert the remainder to a character (A is ASCII 65)
        column_name = chr(65 + remainder) + column_name
        # Move to the next "digit"
        column_number //= 26
    return column_name

def crear_clase_empleado (e):
    return {
        "nombre": e['nombre'],
        "turnos_dia": 0,
        "turnos_noche": 0,
        "descansos": 0,
        'bloqueos_dia': e['bloqueos_dia'],
        'bloqueos_noche': e['bloqueos_noche'],
        'puestos_habilitados': e['puestos_habilitados']
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

def actualizar_descansos (empleados, dia, dia_anterior, PUESTOS):
    puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
    actualizados = []
    empleados_dia = list(dia.values())
    empleados_dia_anterior = [dia_anterior[puesto["nombre"]] for puesto in puestos_nocturnos]
    
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
    df_empleados = pd.read_excel('Entradas.xlsm', sheet_name='Empleados', header=None, names=['nombre', 'puestos_habilitados'])
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
        puestos_empleado = df_empleados.loc[df_empleados['nombre'] == empleado].iloc[0]['puestos_habilitados'].split(', ')
        for index, row in df_empleado.iterrows():
            if (not pd.isna(row['dia'])):
                bloqueos_dia.append(formatear_fecha(row['fecha']))
            if (not pd.isna(row['noche'])):
                bloqueos_noche.append(formatear_fecha(row['fecha']))
        EMPLEADOS.append({
            'nombre': empleado,
            'bloqueos_dia': bloqueos_dia,
            'bloqueos_noche':bloqueos_noche,
            'puestos_habilitados': puestos_empleado
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
        # Bloqueos
        for empleado in empleados:
            # Bloqueo dias libres
            dias_trabajados = empleado["turnos_noche"] + empleado["turnos_dia"]
            if 2 *(dias_trabajados // 5) > empleado["descansos"]:
                bloqueos_descanso.append(empleado["nombre"])

            # Bloqueo por cronograma
            if dia['fecha'] in empleado['bloqueos_dia']:
                bloqueos_dia.append(empleado["nombre"])
            if dia['fecha'] in empleado['bloqueos_noche']:
                bloqueos_noche.append(empleado["nombre"])
          
        # Bloquear empleados que trabajaron la noche anterior 
        if i > 0:
            dia_anterior = cronograma[i - 1]
            puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
            for p in puestos_nocturnos:
                bloqueos_dia.append(dia_anterior[p["nombre"]])
        
        # Asignar puestos
        for puesto in PUESTOS:
            es_nocturno = puesto["nocturno"]
            
            #Bloqueo por puesto
            bloqueos_puesto = [e['nombre'] for e in empleados if puesto['nombre'] not in e['puestos_habilitados']]
            #Empleados habilitados
            if es_nocturno:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_descanso + bloqueos_noche + bloqueos_puesto]
            else:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_descanso + bloqueos_dia + bloqueos_puesto]
            
            # Si no encuentra un trabajador ideal, sacrifica condcion dia libre
            if len(empleados_habilitados) == 0:
                if es_nocturno:
                    empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_noche + bloqueos_puesto]
                else:
                    empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_dia + bloqueos_puesto]
                
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
            empleados = actualizar_descansos(empleados, dia, cronograma[i -1], PUESTOS)

    #Generar excel
    data_excel = {
        'Puestos': [puesto['nombre'] for puesto in PUESTOS],
        'Nocturno': [puesto['nocturno'] for puesto in PUESTOS],
    }
    for dia in cronograma:
        data_excel[dia['fecha']] = [dia[puesto["nombre"]] for puesto in PUESTOS]
    df = pd.DataFrame(data_excel)
    excel_file_path = f"Cronograma-{cronograma[0]['fecha']}-{cronograma[-1]['fecha']}.xlsx"
    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        workbook.set_calc_mode('auto')
        workbook.calc_on_load = True
        df.to_excel(writer, sheet_name='Cronograma', index=False)
        for empleado in empleados:
            cr_empleado = {
                "Día": [f'=SUMPRODUCT((Cronograma!B2:B{len(PUESTOS ) + 1}=FALSE) * (COUNTIF(D2:{get_excel_column_name(len(cronograma) + 3)}2, Cronograma!A2:A{len(PUESTOS ) + 1})))'], 
                "Noche": [f'=SUMPRODUCT((Cronograma!B2:B{len(PUESTOS ) + 1}=TRUE) * (COUNTIF(D2:{get_excel_column_name(len(cronograma) + 3)}2, Cronograma!A2:A{len(PUESTOS ) + 1})))'], 
                "Descanso": [0]
            }
            for index, dia in enumerate(cronograma):
                cr_empleado[dia['fecha']] = [f'=_xlfn.XLOOKUP("{empleado["nombre"]}",Cronograma!{get_excel_column_name(index + 3)}2:{get_excel_column_name(index + 3)}{len(PUESTOS ) + 1},Cronograma!A2:A{len(PUESTOS ) + 1},"")']
            de = pd.DataFrame(cr_empleado)
            de.to_excel(writer, sheet_name=empleado["nombre"], index=False)
            worksheet = writer.sheets[empleado["nombre"]]
            worksheet.write_formula("C2", f'=SUMPRODUCT((E2:{get_excel_column_name(len(cronograma) + 3)}2="") * NOT(IFERROR(VLOOKUP(N(D2:{get_excel_column_name(len(cronograma) + 2)}2),Cronograma!A2:B{len(PUESTOS ) + 1},2,FALSE), FALSE)))')
    # for dia in cronograma:
    #     print(f"\n{dia["fecha"]}:")
    #     for puesto in PUESTOS:
    #         print(f"  - {puesto["nombre"]}: {dia[puesto["nombre"]]}")
