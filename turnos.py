import datetime
import pandas as pd
import networkx as nx

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
        'dias_sin_descanso': 0,
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
            objeto_modificado['dias_sin_descanso'] = 0
            actualizados.append(objeto_modificado)
        else:
            objeto_modificado = empleado.copy()
            objeto_modificado['dias_sin_descanso'] += 1
            actualizados.append(objeto_modificado)
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
            if 2 *(dias_trabajados // 5) > empleado["descansos"] or empleado['dias_sin_descanso'] > 4:
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
        
        puestos_con_disponibilidad = []
        for puesto in PUESTOS:
            nuevo_puesto = puesto.copy()
            es_nocturno = puesto["nocturno"]
            
            #Bloqueo por puesto
            bloqueos_puesto = [e['nombre'] for e in empleados if puesto['nombre'] not in e['puestos_habilitados']]
            #Empleados habilitados
            empleados_habilitados = []
            if es_nocturno:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_descanso + bloqueos_noche + bloqueos_puesto]
            else:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in trabajadores_bloqueados + bloqueos_descanso + bloqueos_dia + bloqueos_puesto]
            nuevo_puesto["empleados_disponibles"] = empleados_habilitados
            puestos_con_disponibilidad.append(nuevo_puesto)

        # Se busca la combinacion mas eficiente usando grafos
        G = nx.Graph()
        objetos_grafo = {}

        for puesto in puestos_con_disponibilidad:
            objetos_grafo[puesto["nombre"]] = [e["nombre"] for e in puesto["empleados_disponibles"]]
        puestos_grafo = list(objetos_grafo.keys())

        for puesto in puestos_grafo:
            G.add_node(puesto, bipartite=0)

        for personas in objetos_grafo.values():
            for persona in personas:
                G.add_node(persona, bipartite=1)
        
        for puesto, personas in objetos_grafo.items():
            for persona in personas:
                G.add_edge(puesto, persona, weight=50)

        matching = nx.algorithms.matching.max_weight_matching(G, maxcardinality=True)

        resultado = {}

        for u, v in matching:
            if u in puestos_grafo:
                resultado[u] = v
            else:
                resultado[v] = u

        # Asignar puestos
        for puesto in puestos_con_disponibilidad:
            es_nocturno = puesto["nocturno"]
            empleado_a_asignar = next((empleado for empleado in empleados if empleado['nombre'] == resultado.get(puesto['nombre'])), None)
            # Solo si hay trabajador disponible
            if empleado_a_asignar:
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
        worksheet = writer.sheets["Cronograma"]
        worksheet.autofit()
        format_red = workbook.add_format({'bg_color': '#FF0000'})
        worksheet.conditional_format(f'C2:{get_excel_column_name(len(cronograma) + 2)}{len(PUESTOS) + 1}', {
            'type':     'formula',
            'criteria': f'=C2=""',
            'format':   format_red
        })

        # Hoja 2
        format_yellow = workbook.add_format({'bg_color': '#FFFF00'})
        format_blue = workbook.add_format({'bg_color': '#0033CC'})
        format_green = workbook.add_format({'bg_color': '#0DBF33'})

        cr_individual = {
            "Nombre": [empleado["nombre"] for empleado in empleados],
            "Día": [f'=SUMPRODUCT((Cronograma!B2:B{len(PUESTOS ) + 1}=FALSE) * (COUNTIF(E{index+2}:{get_excel_column_name(len(cronograma) + 4)}{index+2}, Cronograma!A2:A{len(PUESTOS ) + 1})))' for index, empleado in enumerate(empleados)], 
            "Noche": [f'=SUMPRODUCT((Cronograma!B2:B{len(PUESTOS ) + 1}=TRUE) * (COUNTIF(E{index+2}:{get_excel_column_name(len(cronograma) + 4)}{index+2}, Cronograma!A2:A{len(PUESTOS ) + 1})))' for index, empleado in enumerate(empleados)], 
            "Descanso": [empleado["nombre"] for empleado in empleados],
        }
        for index, dia in enumerate(cronograma):
            cr_individual[dia['fecha']] = [f'=_xlfn.XLOOKUP("{empleado["nombre"]}",Cronograma!{get_excel_column_name(index + 3)}2:{get_excel_column_name(index + 3)}{len(PUESTOS ) + 1},Cronograma!A2:A{len(PUESTOS ) + 1},"")' for empleado in empleados]
        
        di = pd.DataFrame(cr_individual)
        di.to_excel(writer, sheet_name="Empleados", index=False)
        worksheet = writer.sheets["Empleados"]
        
        for index, empleado in enumerate(empleados):
            worksheet.write_dynamic_array_formula(f'D{index+2}', f'=SUMPRODUCT((F{index+2}:{get_excel_column_name(len(cronograma) + 4)}{index+2}="") * NOT(IFERROR(VLOOKUP((E{index+2}:{get_excel_column_name(len(cronograma) + 3)}{index+2}),Cronograma!A2:B{len(PUESTOS ) + 1},2,FALSE), FALSE)))')
        worksheet.conditional_format(f'E2:{get_excel_column_name(len(cronograma) + 4)}{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=VLOOKUP((E2),Cronograma!$A$2:$B${len(PUESTOS ) + 1},2,FALSE)=FALSE',
            'format':   format_yellow
        })
        worksheet.conditional_format(f'B2:B{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=TRUE',
            'format':   format_yellow
        })
        worksheet.conditional_format(f'E2:{get_excel_column_name(len(cronograma) + 4)}{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=VLOOKUP((E2),Cronograma!$A$2:$B${len(PUESTOS ) + 1},2,FALSE)=TRUE',
            'format':   format_blue
        })
        worksheet.conditional_format(f'C2:C{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=TRUE',
            'format':   format_blue
        })
        worksheet.conditional_format(f'F2:{get_excel_column_name(len(cronograma) + 4)}{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=AND(NOT(IFERROR(VLOOKUP((E2),Cronograma!$A$2:$B${len(PUESTOS ) + 1},2,FALSE), FALSE)=TRUE), F2="")',
            'format':   format_green
        })
        worksheet.conditional_format(f'D2:D{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=TRUE',
            'format':   format_green
        })
        
        worksheet.autofit()
        worksheet.protect(options={'format_columns': True, 'format_rows': True})