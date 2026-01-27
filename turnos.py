import datetime
import pandas as pd
import networkx as nx
import sys
import os

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
        'ultimo_turno': None,  # 'dia', 'noche', o None si descansó
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
                objeto_modificado["ultimo_turno"] = 'noche'
            else:
                objeto_modificado["turnos_dia"] += 1
                objeto_modificado["ultimo_turno"] = 'dia'
            actualizados.append(objeto_modificado)
        else:
            actualizados.append(empleado)
    return actualizados

def actualizar_descansos (empleados, dia, dia_anterior, dia_siguiente, PUESTOS):
    """
    Actualiza los descansos de los empleados.
    Condición idéntica a Excel: Un día es descanso si:
    (no trabajar ese día Y no trabajar la noche anterior) O (no trabajar ese día Y no trabajar el día siguiente)
    """
    puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
    puestos_diurnos = [item for item in PUESTOS if not item["nocturno"]]
    actualizados = []
    
    # Empleados que trabajan en el día actual (cualquier turno)
    empleados_dia_actual = list(dia.values())
    empleados_dia_actual = [e for e in empleados_dia_actual if e is not None]
    
    # Empleados que trabajan en turnos nocturnos del día anterior
    empleados_nocturno_anterior = []
    if dia_anterior:
        for puesto in puestos_nocturnos:
            empleado = dia_anterior.get(puesto["nombre"])
            if empleado:
                empleados_nocturno_anterior.append(empleado)
    
    # Empleados que trabajan en turnos diurnos del día siguiente
    empleados_diurno_siguiente = []
    if dia_siguiente:
        for puesto in puestos_diurnos:
            empleado = dia_siguiente.get(puesto["nombre"])
            if empleado:
                empleados_diurno_siguiente.append(empleado)
    
    for empleado in empleados:
        nombre_empleado = empleado["nombre"]
        
        # Verificar si no trabaja el día actual
        no_trabaja_hoy = nombre_empleado not in empleados_dia_actual
        
        if no_trabaja_hoy:
            # Condición 1: no trabajar hoy Y no trabajar la noche anterior
            # IMPORTANTE: Solo aplicar si hay día anterior (no es el primer día)
            condicion1 = False
            if dia_anterior:
                nocturno_anterior_libre = nombre_empleado not in empleados_nocturno_anterior
                condicion1 = nocturno_anterior_libre
            # Si no hay día anterior (primer día), condicion1 = False (no se aplica)
            
            # Condición 2: no trabajar hoy Y no trabajar el día siguiente (diurno)
            # IMPORTANTE: Solo aplicar si hay día siguiente (no es el último día)
            condicion2 = False
            if dia_siguiente:
                diurno_siguiente_libre = nombre_empleado not in empleados_diurno_siguiente
                condicion2 = diurno_siguiente_libre
            # Si no hay día siguiente (último día), condicion2 = False (no se aplica)
            
            # Es día libre si cumple AL MENOS UNA de las dos condiciones aplicables (OR)
            # Si ninguna condición es aplicable (primer y último día en cronograma de 1 día), NO es descanso
            es_dia_libre = condicion1 or condicion2
            
            if es_dia_libre:
                objeto_modificado = empleado.copy()
                objeto_modificado["descansos"] += 1
                objeto_modificado['dias_sin_descanso'] = 0
                objeto_modificado['ultimo_turno'] = None  # Resetear al descansar
                actualizados.append(objeto_modificado)
            else:
                # No es día libre según la condición, pero no trabajó hoy
                objeto_modificado = empleado.copy()
                objeto_modificado['dias_sin_descanso'] += 1
                actualizados.append(objeto_modificado)
        else:
            # Trabajó hoy
            objeto_modificado = empleado.copy()
            objeto_modificado['dias_sin_descanso'] += 1
            actualizados.append(objeto_modificado)
    
    return actualizados

def es_descanso_valido(empleado_nombre, dia_intermedio, dia_anterior_intermedio, dia_siguiente_intermedio, PUESTOS, trabajara_diurno_hoy=None):
    """
    Verifica si un día fue descanso válido para un empleado.
    Un día es descanso válido si:
    (no trabajar ese día Y no trabajar la noche anterior) O (no trabajar ese día Y no trabajar el día siguiente)
    """
    puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
    puestos_diurnos = [item for item in PUESTOS if not item["nocturno"]]
    
    # Verificar si el empleado trabajó ese día (cualquier turno)
    empleado_trabajo_ese_dia = False
    for puesto_nombre in dia_intermedio:
        if puesto_nombre != 'fecha' and dia_intermedio[puesto_nombre] == empleado_nombre:
            empleado_trabajo_ese_dia = True
            break
    
    # Si trabajó ese día, no es descanso
    if empleado_trabajo_ese_dia:
        return False
    
    # Condición 1: no trabajar ese día Y no trabajar la noche anterior
    # IMPORTANTE: Solo aplicar si hay día anterior (no es el primer día)
    condicion1 = False
    if dia_anterior_intermedio:
        nocturno_anterior_libre = True
        for puesto in puestos_nocturnos:
            if dia_anterior_intermedio.get(puesto["nombre"]) == empleado_nombre:
                nocturno_anterior_libre = False
                break
        condicion1 = nocturno_anterior_libre
    # Si no hay día anterior (primer día), condicion1 = False (no se aplica)
    
    # Condición 2: no trabajar ese día Y no trabajar el día siguiente (diurno)
    # IMPORTANTE: Solo aplicar si hay día siguiente o se especifica explícitamente
    condicion2 = False
    if trabajara_diurno_hoy is not None:
        # Si se especifica explícitamente si trabajará diurno hoy, usar esa información
        diurno_siguiente_libre = not trabajara_diurno_hoy
        condicion2 = diurno_siguiente_libre
    elif dia_siguiente_intermedio:
        # Verificar en el cronograma del día siguiente
        diurno_siguiente_libre = True
        for puesto in puestos_diurnos:
            if dia_siguiente_intermedio.get(puesto["nombre"]) == empleado_nombre:
                diurno_siguiente_libre = False
                break
        condicion2 = diurno_siguiente_libre
    # Si no hay día siguiente y no se especifica, condicion2 = False (no se aplica)
    
    # Es descanso válido si cumple AL MENOS UNA de las dos condiciones
    return condicion1 or condicion2

def formatear_fecha (fecha):
    return fecha.date()

def calcular_peso_persona (
    empleado,
    puesto,
    dias,
    todos_empleados=None,
    empleados_disponibles_puesto=None,
    cronograma=None,
    dia_idx=None,
    PUESTOS=None,
):
    peso_base = 10
    dias_trabajados = empleado['turnos_dia'] + empleado['turnos_noche']
    nocturno=puesto['nocturno']
    # Que baja el peso:  pocos dias de descanso, muchos dias seguidos, muchos turnos de este horario

    # Dias de descanso el balance debe ser de 2 dias de descanso por 5 de trabajo. En la proporción cada dia representa 0.2 y el valor ideal es 0.4.
    # Cuando se trabaja mucho (ratio bajo), se debe desincentivar trabajar más (peso más bajo = variable negativa)
    # PENALIZACIÓN MUY ALTA para empleados con 0 días trabajados
    if (dias_trabajados == 0):
        variable_por_descansos = 100  # Aumentado de 5 a 100 para evitar empleados con 0 días trabajados
    else:
        # Corregido: la fórmula original estaba invertida. 
        # Cuando ratio es bajo (trabaja mucho, pocos descansos) → variable debe ser negativa para bajar peso y desincentivar
        # Cuando ratio es alto (descansa mucho) → variable debe ser positiva para subir peso e incentivar
        ratio_descansos = empleado['descansos'] / (empleado['turnos_dia'] + empleado['turnos_noche'])
        # Fórmula corregida (invertida): (0.4 - ratio) * -10 es equivalente a (ratio - 0.4) * 10
        # pero escrita de forma invertida para claridad
        variable_por_descansos = (0.4 - ratio_descansos) * -10
    
    # Balanceo de descansos: penalizar si tiene muchos más descansos que el promedio
    variable_balance_descansos = 0
    if todos_empleados and len(todos_empleados) > 1:
        # Calcular promedio de descansos
        total_descansos = sum(e['descansos'] for e in todos_empleados)
        promedio_descansos = total_descansos / len(todos_empleados)
        # Si tiene más descansos que el promedio, penalizar (incentivar a trabajar)
        diferencia_descansos = empleado['descansos'] - promedio_descansos
        if diferencia_descansos > 0:
            # Penalizar proporcionalmente (más descansos = más penalización) - AUMENTADO
            variable_balance_descansos = -diferencia_descansos * 10  # Aumentado de 3 a 10
    
    # Dias de trabajo seguidos (streak real calculado desde el cronograma)
    # OJO: en este algoritmo, empleado['dias_sin_descanso'] se actualiza al final, así que NO sirve
    # para influir decisiones durante el armado. Por eso lo calculamos dinámicamente aquí.
    dias_seguidos_trabajados = None
    if cronograma is not None and dia_idx is not None:
        streak = 0
        for j in range(dia_idx - 1, -1, -1):
            dia_verificar = cronograma[j]
            trabajo = False
            for puesto_nombre in dia_verificar:
                if puesto_nombre != 'fecha' and dia_verificar[puesto_nombre] == empleado['nombre']:
                    trabajo = True
                    break
            if trabajo:
                streak += 1
            else:
                break
        dias_seguidos_trabajados = streak
    else:
        dias_seguidos_trabajados = empleado.get('dias_sin_descanso', 0)

    # Penalización progresiva para desincentivar llegar a 5 días seguidos
    # + INCENTIVOS AUMENTADOS para consolidar bloques más largos (1-3 días) sin empujar a violar la regla de 5.
    dias_sin_descanso = dias_seguidos_trabajados
    if dias_sin_descanso >= 6:
        variable_por_trabajos = -200  # Penalización muy alta después de 5 días (casi imposible de seleccionar)
    elif dias_sin_descanso == 5:
        variable_por_trabajos = -100  # Penalización muy alta al llegar a 5 días
    elif dias_sin_descanso == 4:
        variable_por_trabajos = -50   # Penalización alta para desincentivar llegar a 5
    elif dias_sin_descanso == 3:
        variable_por_trabajos = 15    # INCENTIVO AUMENTADO: continuar el bloque (aumentado de -20 a +15)
    elif dias_sin_descanso == 2:
        variable_por_trabajos = 25    # INCENTIVO AUMENTADO: continuar el bloque (aumentado de 5 a 25)
    elif dias_sin_descanso == 1:
        variable_por_trabajos = 20    # INCENTIVO AUMENTADO: empezar a formar un bloque (aumentado de 3 a 20)
    else:
        variable_por_trabajos = 0     # Sin penalización ni incentivo para 0 días

    # Variable por disparidad entre jornadas
    # Cuando ya se trabajó mucho de este tipo de turno, se debe desincentivar trabajar más (peso más bajo = variable negativa)
    # PERO incentivamos cambio de jornada después de descanso
    # PENALIZACIÓN MUY AGRESIVA para evitar que empleados terminen con 0 turnos de un tipo
    variable_jornadas = 0
    turnos_dia = empleado['turnos_dia']
    turnos_noche = empleado['turnos_noche']
    
    # CASO CRÍTICO: Si tiene 0 turnos de un tipo, penalización EXTREMADAMENTE ALTA para el tipo que ya tiene
    # Esto debe ser tan alto que se prefiera dejar el turno vacío antes que asignarlo
    # La penalización debe crecer con el número de turnos del tipo opuesto que ya tiene
    if turnos_dia == 0 and nocturno:
        # Tiene 0 diurnos y se le está ofreciendo otro nocturno - PENALIZACIÓN EXTREMADAMENTE ALTA
        # Penalización base muy alta + penalización adicional por cada turno nocturno que ya tiene
        # Esto hace que sea preferible dejar el turno vacío si ya tiene varios turnos nocturnos
        variable_jornadas = -5000 - (turnos_noche * 1000)  # Base -5000, -1000 por cada turno nocturno adicional
    elif turnos_noche == 0 and not nocturno:
        # Tiene 0 nocturnos y se le está ofreciendo otro diurno - PENALIZACIÓN EXTREMADAMENTE ALTA
        # Penalización base muy alta + penalización adicional por cada turno diurno que ya tiene
        variable_jornadas = -5000 - (turnos_dia * 1000)  # Base -5000, -1000 por cada turno diurno adicional
    elif turnos_dia == 0 and not nocturno:
        # Tiene 0 diurnos y se le ofrece un diurno - INCENTIVO MUY ALTO
        # Si hay múltiples empleados con 0 diurnos, normalizar pesos para distribución equitativa
        if todos_empleados:
            num_empleados_0_diurnos = sum(1 for e in todos_empleados if e['turnos_dia'] == 0)
            if num_empleados_0_diurnos > 1:
                # Base fija para que todos tengan pesos similares
                variable_jornadas = 200
                # Pequeño ajuste basado en turnos nocturnos para evitar empates exactos
                # pero mantener pesos muy cercanos para distribución equitativa
                variable_jornadas += empleado['turnos_noche'] * 3  # Muy pequeño para mantener similitud
            else:
                # Solo un empleado con 0 diurnos, incentivo normal más alto
                variable_jornadas = 200 + (empleado['turnos_noche'] * 20)
        else:
            variable_jornadas = 200
    elif turnos_noche == 0 and nocturno:
        # Tiene 0 nocturnos y se le ofrece un nocturno - INCENTIVO MUY ALTO
        # Si hay múltiples empleados con 0 nocturnos, normalizar pesos para distribución equitativa
        if todos_empleados:
            num_empleados_0_nocturnos = sum(1 for e in todos_empleados if e['turnos_noche'] == 0)
            if num_empleados_0_nocturnos > 1:
                # Base fija para que todos tengan pesos similares
                variable_jornadas = 200
                # Pequeño ajuste basado en turnos diurnos para evitar empates exactos
                variable_jornadas += empleado['turnos_dia'] * 3  # Muy pequeño para mantener similitud
            else:
                # Solo un empleado con 0 nocturnos, incentivo normal más alto
                variable_jornadas = 200 + (empleado['turnos_dia'] * 20)
        else:
            variable_jornadas = 200
    else:
        # Caso normal: penalizar según el desbalance, pero de forma MUY AGRESIVA
        # La penalización crece exponencialmente con la diferencia para evitar casos extremos
        if nocturno:
            diferencia = turnos_noche - turnos_dia
            if diferencia > 0:
                # Si la diferencia es muy grande (más de 5), penalización EXTREMA similar al caso de 0 turnos
                if diferencia > 5:
                    variable_jornadas = -3000 - (diferencia * 500)  # Penalización extrema para diferencias muy grandes
                else:
                    # Penalización base más agresiva: multiplicar por 100 (aumentado de 20)
                    variable_jornadas = -diferencia * 100
                    # Penalización escalonada que crece rápidamente
                    if diferencia > 1:
                        variable_jornadas = -diferencia * 200  # Diferencia de 2 o más
                    if diferencia > 2:
                        variable_jornadas = -diferencia * 400  # Diferencia de 3 o más
                    if diferencia > 3:
                        variable_jornadas = -diferencia * 800  # Diferencia de 4 o más
                    if diferencia > 4:
                        variable_jornadas = -diferencia * 1500  # Diferencia de 5 o más
            else:
                # Si tiene más diurnos que nocturnos, incentivar para balancear
                variable_jornadas = abs(diferencia) * 100  # Aumentado de 20 a 100
        else:
            diferencia = turnos_dia - turnos_noche
            if diferencia > 0:
                # Si la diferencia es muy grande (más de 5), penalización EXTREMA similar al caso de 0 turnos
                if diferencia > 5:
                    variable_jornadas = -3000 - (diferencia * 500)  # Penalización extrema para diferencias muy grandes
                else:
                    # Penalización base más agresiva: multiplicar por 100 (aumentado de 20)
                    variable_jornadas = -diferencia * 100
                    # Penalización escalonada que crece rápidamente
                    if diferencia > 1:
                        variable_jornadas = -diferencia * 200  # Diferencia de 2 o más
                    if diferencia > 2:
                        variable_jornadas = -diferencia * 400  # Diferencia de 3 o más
                    if diferencia > 3:
                        variable_jornadas = -diferencia * 800  # Diferencia de 4 o más
                    if diferencia > 4:
                        variable_jornadas = -diferencia * 1500  # Diferencia de 5 o más
            else:
                # Si tiene más nocturnos que diurnos, incentivar para balancear
                variable_jornadas = abs(diferencia) * 100  # Aumentado de 20 a 100
    
    # INCENTIVO: Cambio de jornada después de descanso
    # Si hay un descanso válido entre el último turno del tipo opuesto y hoy,
    # y el turno de HOY es del tipo opuesto, dar incentivo.
    variable_cambio_jornada = 0
    if cronograma is not None and dia_idx is not None and dia_idx > 0 and PUESTOS is not None:
        # Buscar el último turno del tipo opuesto trabajado
        puestos_nocturnos = [p for p in PUESTOS if p['nocturno']]
        puestos_diurnos = [p for p in PUESTOS if not p['nocturno']]
        puestos_tipo_opuesto = puestos_diurnos if nocturno else puestos_nocturnos
        
        ultimo_dia_tipo_opuesto = None
        ultimo_tipo = None
        
        # Buscar el último día que trabajó el tipo opuesto
        for j in range(dia_idx - 1, -1, -1):
            dia_verificar = cronograma[j]
            trabajo_tipo_opuesto = False
            puesto_encontrado = None
            
            for p_opuesto in puestos_tipo_opuesto:
                if dia_verificar.get(p_opuesto["nombre"]) == empleado['nombre']:
                    trabajo_tipo_opuesto = True
                    puesto_encontrado = p_opuesto["nombre"]
                    ultimo_dia_tipo_opuesto = j
                    break
            
            if trabajo_tipo_opuesto:
                ultimo_tipo = 'noche' if nocturno else 'dia'  # El tipo opuesto al que se le está ofreciendo
                break
        
        # Si encontró un turno del tipo opuesto, verificar si hay descanso válido
        if ultimo_dia_tipo_opuesto is not None:
            hay_descanso_valido = False
            
            # Verificar todos los días intermedios entre el último turno del tipo opuesto y hoy
            for dia_intermedio_idx in range(ultimo_dia_tipo_opuesto + 1, dia_idx):
                dia_intermedio = cronograma[dia_intermedio_idx]
                dia_anterior_intermedio = cronograma[dia_intermedio_idx - 1] if dia_intermedio_idx > 0 else None
                dia_siguiente_intermedio = cronograma[dia_intermedio_idx + 1] if dia_intermedio_idx < len(cronograma) - 1 else None
                
                # Si es el día inmediatamente anterior a hoy, usar la información de que trabajará hoy
                if dia_intermedio_idx == dia_idx - 1:
                    if es_descanso_valido(empleado['nombre'], dia_intermedio,
                                         dia_anterior_intermedio, dia_siguiente_intermedio, PUESTOS,
                                         trabajara_diurno_hoy=not nocturno):
                        hay_descanso_valido = True
                        break
                else:
                    # Para días intermedios que no son el último, verificar normalmente
                    if es_descanso_valido(empleado['nombre'], dia_intermedio,
                                         dia_anterior_intermedio, dia_siguiente_intermedio, PUESTOS,
                                         trabajara_diurno_hoy=None):
                        hay_descanso_valido = True
                        break
            
            # Si hay descanso válido y el turno de hoy es opuesto, aplicar incentivo
            if hay_descanso_valido:
                # Incentivo base por cambio de jornada después de descanso
                variable_cambio_jornada = 300  # Aumentado de 100 a 300
                # Aumentar el incentivo si hay desbalance significativo
                diferencia = abs(empleado['turnos_dia'] - empleado['turnos_noche'])
                if diferencia > 0:
                    # Incentivo adicional proporcional al desbalance
                    variable_cambio_jornada += diferencia * 150  # Aumentado de 50 a 150
                # Incentivo extra si tiene muy pocos turnos de un tipo
                if nocturno and empleado['turnos_dia'] < 3:
                    variable_cambio_jornada += 500  # Aumentado de 200 a 500
                elif not nocturno and empleado['turnos_noche'] < 3:
                    variable_cambio_jornada += 500  # Aumentado de 200 a 500
        else:
            # Si no hay turno previo del tipo opuesto, pero hay desbalance, aplicar incentivo menor
            # (esto cubre el caso de empleados que nunca han trabajado un tipo)
            if empleado['turnos_dia'] > empleado['turnos_noche'] and nocturno:
                # Incentivo alto si tiene más diurnos y se le ofrece nocturno
                diferencia = empleado['turnos_dia'] - empleado['turnos_noche']
                variable_cambio_jornada = 150 + (diferencia * 100)  # Aumentado de 50+30 a 150+100
            elif empleado['turnos_noche'] > empleado['turnos_dia'] and not nocturno:
                # Incentivo alto si tiene más nocturnos y se le ofrece diurno
                diferencia = empleado['turnos_noche'] - empleado['turnos_dia']
                variable_cambio_jornada = 150 + (diferencia * 100)  # Aumentado de 50+30 a 150+100
            # Incentivo adicional si tiene muy pocos turnos de un tipo
            if nocturno and empleado['turnos_dia'] < 3:
                variable_cambio_jornada += 400  # Aumentado de 150 a 400
            elif not nocturno and empleado['turnos_noche'] < 3:
                variable_cambio_jornada += 400  # Aumentado de 150 a 400
    
    # Incentivo por especialización del empleado: menos puestos habilitados = más incentivo
    # Empleados más especializados (menos puestos) deben tener prioridad
    num_puestos_empleado = len(empleado.get('puestos_habilitados', []))
    if num_puestos_empleado > 0:
        # Incentivo inversamente proporcional: menos puestos = más incentivo
        # Normalizar: si tiene 1 puesto, incentivo máximo; si tiene muchos, menos incentivo
        # Usamos 1/num_puestos como base, multiplicado por un factor
        variable_especializacion_empleado = (1.0 / num_puestos_empleado) * 5
    else:
        variable_especializacion_empleado = 0
    
    # Incentivo por especialización del puesto: menos empleados disponibles = más incentivo
    # Puestos que pueden ser cubiertos por menos empleados deben llenarse más rápido
    if empleados_disponibles_puesto is not None:
        num_empleados_puesto = len(empleados_disponibles_puesto)
        if num_empleados_puesto > 0:
            # Incentivo inversamente proporcional: menos empleados = más incentivo
            # Normalizar: si solo 1 empleado puede cubrirlo, incentivo máximo
            variable_especializacion_puesto = (1.0 / num_empleados_puesto) * 5
        else:
            variable_especializacion_puesto = 0
    else:
        variable_especializacion_puesto = 0

    return (
        peso_base
        + variable_por_descansos
        + variable_por_trabajos
        + variable_jornadas
        + variable_balance_descansos
        + variable_especializacion_empleado
        + variable_especializacion_puesto
        + variable_cambio_jornada
    )

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
        print('\n\n------'+dia['fecha'].strftime("%Y-%m-%d"))
        bloqueos_dia = []
        bloqueos_noche = []
        
        # Bloqueos
        for empleado in empleados:

            # Bloqueo por cronograma
            if dia['fecha'] in empleado['bloqueos_dia']:
                bloqueos_dia.append(empleado["nombre"])
            if dia['fecha'] in empleado['bloqueos_noche']:
                bloqueos_noche.append(empleado["nombre"])
            
            # Bloqueo obligatorio: no trabajar más de 5 días seguidos
            # Calcular días seguidos trabajados directamente desde el cronograma
            dias_seguidos_trabajados = 0
            for j in range(i - 1, -1, -1):  # Retroceder desde el día anterior
                dia_verificar = cronograma[j]
                # Verificar si el empleado trabajó ese día (en cualquier puesto)
                empleado_trabajo = False
                for puesto_nombre in dia_verificar:
                    if puesto_nombre != 'fecha' and dia_verificar[puesto_nombre] == empleado["nombre"]:
                        empleado_trabajo = True
                        break
                
                if empleado_trabajo:
                    dias_seguidos_trabajados += 1
                else:
                    # Si no trabajó, verificar si fue día libre según la nueva condición
                    # (esto requiere verificar el día anterior y siguiente, pero para simplificar,
                    #  si no trabajó, consideramos que descansó)
                    break
            
            # Si ya trabajó 5 días seguidos, bloquear
            if dias_seguidos_trabajados >= 5:
                # Bloquear para todos los turnos (diurnos y nocturnos)
                if empleado["nombre"] not in bloqueos_dia:
                    bloqueos_dia.append(empleado["nombre"])
                if empleado["nombre"] not in bloqueos_noche:
                    bloqueos_noche.append(empleado["nombre"])
          
        # Bloquear cambio de tipo de turno sin descanso válido
        # Para cambiar de jornada (diurno a nocturno o viceversa) se requiere un día de descanso válido
        puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
        puestos_diurnos = [item for item in PUESTOS if not item["nocturno"]]
        
        # Para cada puesto, verificar si algún empleado que trabajó el tipo opuesto puede trabajar aquí
        for puesto in PUESTOS:
            es_nocturno = puesto["nocturno"]
            tipo_opuesto = "diurno" if es_nocturno else "nocturno"
            puestos_tipo_opuesto = puestos_diurnos if es_nocturno else puestos_nocturnos
            
            # Buscar el último día que cada empleado trabajó el tipo opuesto
            for empleado in empleados:
                nombre_empleado = empleado["nombre"]
                
                # Buscar el último día que trabajó el tipo opuesto (hacia atrás desde hoy)
                ultimo_dia_tipo_opuesto = None
                for j in range(i - 1, -1, -1):  # Desde el día anterior hacia atrás
                    dia_verificar = cronograma[j]
                    trabajo_tipo_opuesto = False
                    
                    # Verificar si trabajó en algún puesto del tipo opuesto ese día
                    for p_opuesto in puestos_tipo_opuesto:
                        if dia_verificar.get(p_opuesto["nombre"]) == nombre_empleado:
                            trabajo_tipo_opuesto = True
                            ultimo_dia_tipo_opuesto = j
                            break
                    
                    if trabajo_tipo_opuesto:
                        break
                
                # Si encontró un día donde trabajó el tipo opuesto, verificar que haya descanso válido
                if ultimo_dia_tipo_opuesto is not None:
                    # Verificar que haya al menos un día de descanso válido entre el último día del tipo opuesto y hoy
                    hay_descanso_valido = False
                    
                    # Verificar todos los días intermedios
                    for dia_intermedio_idx in range(ultimo_dia_tipo_opuesto + 1, i):
                        dia_intermedio = cronograma[dia_intermedio_idx]
                        dia_anterior_intermedio = cronograma[dia_intermedio_idx - 1] if dia_intermedio_idx > 0 else None
                        dia_siguiente_intermedio = cronograma[dia_intermedio_idx + 1] if dia_intermedio_idx < len(cronograma) - 1 else None
                        
                        # Si es el día inmediatamente anterior a hoy, usar la información de que trabajará hoy
                        if dia_intermedio_idx == i - 1:
                            if es_descanso_valido(nombre_empleado, dia_intermedio,
                                                 dia_anterior_intermedio, dia_siguiente_intermedio, PUESTOS,
                                                 trabajara_diurno_hoy=not es_nocturno):
                                hay_descanso_valido = True
                                break
                        else:
                            # Para días intermedios que no son el último, verificar normalmente
                            if es_descanso_valido(nombre_empleado, dia_intermedio,
                                                 dia_anterior_intermedio, dia_siguiente_intermedio, PUESTOS,
                                                 trabajara_diurno_hoy=None):
                                hay_descanso_valido = True
                                break
                    
                    # Si no hay descanso válido, bloquear
                    if not hay_descanso_valido:
                        if es_nocturno and nombre_empleado not in bloqueos_noche:
                            bloqueos_noche.append(nombre_empleado)
                        elif not es_nocturno and nombre_empleado not in bloqueos_dia:
                            bloqueos_dia.append(nombre_empleado)
        
        puestos_con_disponibilidad = []
        for puesto in PUESTOS:
            nuevo_puesto = puesto.copy()
            es_nocturno = puesto["nocturno"]
            
            #Bloqueo por puesto
            bloqueos_puesto = [e['nombre'] for e in empleados if puesto['nombre'] not in e['puestos_habilitados']]
            #Empleados habilitados
            empleados_habilitados = []
            if es_nocturno:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in bloqueos_noche + bloqueos_puesto]
            else:
                empleados_habilitados = [item for item in empleados if item["nombre"] not in bloqueos_dia + bloqueos_puesto]
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
            print('\n   '+puesto)
            # Obtener la lista de empleados disponibles para este puesto
            puesto_info = next((p for p in puestos_con_disponibilidad if p['nombre'] == puesto), None)
            empleados_disponibles_puesto = puesto_info['empleados_disponibles'] if puesto_info else []
            
            for persona in personas:
                empleado_info = next((empleado for empleado in empleados if empleado['nombre'] == persona), None)
                puesto_info_peso = next((p for p in PUESTOS if p['nombre'] == puesto), None)
                
                peso = calcular_peso_persona(
                    empleado_info,
                    puesto_info_peso,
                    len(cronograma),
                    todos_empleados=empleados,  # Pasar todos los empleados para balanceo
                    empleados_disponibles_puesto=empleados_disponibles_puesto,  # Pasar empleados disponibles para este puesto
                    cronograma=cronograma,  # Pasar cronograma para verificar descansos
                    dia_idx=i,  # Pasar índice del día actual
                    PUESTOS=PUESTOS,
                )
                empleadofull = next((empleado for empleado in empleados if empleado['nombre'] == persona), None)
                print(f'      {persona}: {str(peso)}, {empleadofull["descansos"]}, {empleadofull["dias_sin_descanso"]}, {empleadofull["turnos_dia"]}, {empleadofull["turnos_noche"]}')
                G.add_edge(
                    puesto, 
                    persona, 
                    weight=peso
                )

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
    
    # POST-PROCESAMIENTO: Corregir desbalances extremos (empleados con 0 turnos de un tipo)
    # Buscar empleados con 0 turnos de un tipo y oportunidades de intercambio
    empleados_con_0_diurnos = [e for e in empleados if e['turnos_dia'] == 0 and e['turnos_noche'] > 0]
    empleados_con_0_nocturnos = [e for e in empleados if e['turnos_noche'] == 0 and e['turnos_dia'] > 0]
    
    # Para cada empleado con 0 diurnos, buscar oportunidades de intercambio
    for empleado_0_diurnos in empleados_con_0_diurnos:
        nombre_0_diurnos = empleado_0_diurnos['nombre']
        
        # Buscar en el cronograma días donde este empleado trabajó nocturno
        # y hay otros empleados con muchos diurnos que podrían intercambiar
        for i in range(len(cronograma)):
            dia = cronograma[i]
            
            # Buscar si este empleado trabajó en un puesto nocturno este día
            puesto_nocturno_empleado = None
            for puesto in PUESTOS:
                if puesto['nocturno'] and dia.get(puesto['nombre']) == nombre_0_diurnos:
                    puesto_nocturno_empleado = puesto['nombre']
                    break
            
            if puesto_nocturno_empleado:
                # Buscar empleados con muchos diurnos que trabajaron diurno este día
                for otro_empleado in empleados:
                    if otro_empleado['turnos_dia'] >= 3 and otro_empleado['turnos_noche'] == 0:
                        nombre_otro = otro_empleado['nombre']
                        
                        # Buscar si este otro empleado trabajó en un puesto diurno este día
                        puesto_diurno_otro = None
                        for puesto in PUESTOS:
                            if not puesto['nocturno'] and dia.get(puesto['nombre']) == nombre_otro:
                                puesto_diurno_otro = puesto['nombre']
                                break
                        
                        # Si encontramos un intercambio posible, verificar que sea válido
                        if puesto_diurno_otro:
                            # Verificar que ambos empleados estén habilitados para los puestos
                            # Obtener puestos habilitados desde df_empleados
                            try:
                                puestos_habilitados_0_diurnos = df_empleados.loc[df_empleados['nombre'] == nombre_0_diurnos].iloc[0]['puestos_habilitados'].split(', ')
                                puestos_habilitados_otro = df_empleados.loc[df_empleados['nombre'] == nombre_otro].iloc[0]['puestos_habilitados'].split(', ')
                                
                                if (puesto_diurno_otro in puestos_habilitados_0_diurnos and
                                    puesto_nocturno_empleado in puestos_habilitados_otro):
                                    
                                    # Verificar que no haya bloqueos (descanso válido, días seguidos, etc.)
                                    # Para simplificar, solo intercambiar si ambos tienen descanso válido
                                    dia_anterior = cronograma[i - 1] if i > 0 else None
                                    dia_siguiente = cronograma[i + 1] if i < len(cronograma) - 1 else None
                                    
                                    # Verificar descanso válido para el cambio
                                    hay_descanso_0_diurnos = es_descanso_valido(nombre_0_diurnos, dia_anterior if dia_anterior else dia,
                                                                               None, dia_siguiente, PUESTOS, trabajara_diurno_hoy=True)
                                    hay_descanso_otro = es_descanso_valido(nombre_otro, dia_anterior if dia_anterior else dia,
                                                                          None, dia_siguiente, PUESTOS, trabajara_diurno_hoy=False)
                                    
                                    if hay_descanso_0_diurnos and hay_descanso_otro:
                                        # Realizar el intercambio
                                        dia[puesto_diurno_otro] = nombre_0_diurnos
                                        dia[puesto_nocturno_empleado] = nombre_otro
                                        
                                        # Actualizar contadores
                                        empleado_0_diurnos['turnos_dia'] += 1
                                        empleado_0_diurnos['turnos_noche'] -= 1
                                        otro_empleado['turnos_dia'] -= 1
                                        otro_empleado['turnos_noche'] += 1
                                        
                                        cronograma[i] = dia
                                        # Salir del bucle interno después de un intercambio
                                        break
                            except (IndexError, KeyError):
                                # Si no se encuentra el empleado en df_empleados, continuar
                                pass
                
                # Si se hizo un intercambio, salir del bucle de días
                if empleado_0_diurnos['turnos_dia'] > 0:
                    break
    
    # Similar para empleados con 0 nocturnos (intercambiar con empleados con muchos nocturnos)
    for empleado_0_nocturnos in empleados_con_0_nocturnos:
        nombre_0_nocturnos = empleado_0_nocturnos['nombre']
        
        # Buscar en el cronograma días donde este empleado trabajó diurno
        for i in range(len(cronograma)):
            dia = cronograma[i]
            
            # Buscar si este empleado trabajó en un puesto diurno este día
            puesto_diurno_empleado = None
            for puesto in PUESTOS:
                if not puesto['nocturno'] and dia.get(puesto['nombre']) == nombre_0_nocturnos:
                    puesto_diurno_empleado = puesto['nombre']
                    break
            
            if puesto_diurno_empleado:
                # Buscar empleados con muchos nocturnos que trabajaron nocturno este día
                for otro_empleado in empleados:
                    if otro_empleado['turnos_noche'] >= 3 and otro_empleado['turnos_dia'] == 0:
                        nombre_otro = otro_empleado['nombre']
                        
                        # Buscar si este otro empleado trabajó en un puesto nocturno este día
                        puesto_nocturno_otro = None
                        for puesto in PUESTOS:
                            if puesto['nocturno'] and dia.get(puesto['nombre']) == nombre_otro:
                                puesto_nocturno_otro = puesto['nombre']
                                break
                        
                        # Si encontramos un intercambio posible, verificar que sea válido
                        if puesto_nocturno_otro:
                            # Verificar que ambos empleados estén habilitados para los puestos
                            # Obtener puestos habilitados desde df_empleados
                            try:
                                puestos_habilitados_0_nocturnos = df_empleados.loc[df_empleados['nombre'] == nombre_0_nocturnos].iloc[0]['puestos_habilitados'].split(', ')
                                puestos_habilitados_otro = df_empleados.loc[df_empleados['nombre'] == nombre_otro].iloc[0]['puestos_habilitados'].split(', ')
                                
                                if (puesto_nocturno_otro in puestos_habilitados_0_nocturnos and
                                    puesto_diurno_empleado in puestos_habilitados_otro):
                                    
                                    # Verificar descanso válido para el cambio
                                    dia_anterior = cronograma[i - 1] if i > 0 else None
                                    dia_siguiente = cronograma[i + 1] if i < len(cronograma) - 1 else None
                                    
                                    hay_descanso_0_nocturnos = es_descanso_valido(nombre_0_nocturnos, dia_anterior if dia_anterior else dia,
                                                                                  None, dia_siguiente, PUESTOS, trabajara_diurno_hoy=False)
                                    hay_descanso_otro = es_descanso_valido(nombre_otro, dia_anterior if dia_anterior else dia,
                                                                          None, dia_siguiente, PUESTOS, trabajara_diurno_hoy=True)
                                    
                                    if hay_descanso_0_nocturnos and hay_descanso_otro:
                                        # Realizar el intercambio
                                        dia[puesto_nocturno_otro] = nombre_0_nocturnos
                                        dia[puesto_diurno_empleado] = nombre_otro
                                        
                                        # Actualizar contadores
                                        empleado_0_nocturnos['turnos_noche'] += 1
                                        empleado_0_nocturnos['turnos_dia'] -= 1
                                        otro_empleado['turnos_noche'] -= 1
                                        otro_empleado['turnos_dia'] += 1
                                        
                                        cronograma[i] = dia
                                        # Salir del bucle interno después de un intercambio
                                        break
                            except (IndexError, KeyError):
                                # Si no se encuentra el empleado en df_empleados, continuar
                                pass
                
                # Si se hizo un intercambio, salir del bucle de días
                if empleado_0_nocturnos['turnos_noche'] > 0:
                    break
    
    # Actualizar los días de descanso después de completar todo el cronograma
    # (necesitamos el día siguiente para la nueva condición)
    for i in range(len(cronograma)):
        dia = cronograma[i]
        dia_anterior = cronograma[i - 1] if i > 0 else None
        dia_siguiente = cronograma[i + 1] if i < len(cronograma) - 1 else None
        empleados = actualizar_descansos(empleados, dia, dia_anterior, dia_siguiente, PUESTOS)

    # Verificar si se está ejecutando desde optimize_pesos.py para no generar Excel
    # (optimize_pesos.py ejecuta el código compilado con nombre 'turnos.py(inmem)')
    import inspect
    generar_excel = True
    try:
        frame = inspect.currentframe()
        if frame and frame.f_back:
            filename = str(frame.f_back.f_globals.get('__file__', ''))
            if 'turnos.py(inmem)' in filename:
                # Se está ejecutando desde optimize_pesos.py, no generar Excel
                generar_excel = False
    except:
        pass
    
    if generar_excel:
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
        
        # Asegurar que la hoja Empleados esté disponible
        if "Empleados" not in writer.sheets:
            raise Exception("La hoja Empleados no se creó correctamente")
        
        # Fórmula para calcular descansos según la nueva lógica:
        # Un día es descanso si: ese día no hay turnos asignados Y el día anterior no se trabajó nocturno
        # O si: ese día no se trabajó Y al día siguiente no se trabaja diurno
        # OPTIMIZACIÓN: Dividir en bloques para evitar fórmulas muy largas que excedan el límite de Excel
        for index, empleado in enumerate(empleados):
            nombre_emp = empleado["nombre"]
            num_dias = len(cronograma)
            # Usar referencia específica para cada fila
            # Cada fórmula se escribe en D{index+2}, y necesita referirse a A{index+2}
            fila_actual = index + 2
            nombre_emp_excel = f'$A{fila_actual}'  # Columna A absoluta, fila específica
            
            # Dividir en bloques de máximo 15 días para evitar fórmulas muy largas
            # Con 31 días, bloques de 30 generan fórmulas de ~11000 caracteres (muy largas)
            tamano_bloque = 15
            bloques_formula = []
            
            for bloque_inicio in range(0, num_dias, tamano_bloque):
                bloque_fin = min(bloque_inicio + tamano_bloque, num_dias)
                partes_bloque = []
                
                for dia_idx in range(bloque_inicio, bloque_fin):
                    col_dia_cronograma = get_excel_column_name(3 + dia_idx)
                    # Versión más compacta: usar SUMPRODUCT para verificar si NO trabaja ese día
                    dia_vacio = f'SUMPRODUCT((Cronograma!{col_dia_cronograma}$2:{col_dia_cronograma}${len(PUESTOS) + 1}={nombre_emp_excel})*1)=0'
                    
                    condiciones_descanso = []
                    
                    # Construir condiciones de forma más compacta
                    if dia_idx > 0:
                        col_anterior = get_excel_column_name(3 + dia_idx - 1)
                        # Verificar si trabajó nocturno ayer
                        trabajo_ayer_nocturno = f'SUMPRODUCT((Cronograma!{col_anterior}$2:{col_anterior}${len(PUESTOS) + 1}={nombre_emp_excel})*(Cronograma!$B$2:$B${len(PUESTOS) + 1}=TRUE))>0'
                        cond1 = f'({dia_vacio})*NOT({trabajo_ayer_nocturno})'
                        condiciones_descanso.append(cond1)
                    
                    if dia_idx < num_dias - 1:
                        col_siguiente = get_excel_column_name(3 + dia_idx + 1)
                        # Verificar si trabajará diurno mañana
                        trabajo_manana_diurno = f'SUMPRODUCT((Cronograma!{col_siguiente}$2:{col_siguiente}${len(PUESTOS) + 1}={nombre_emp_excel})*(Cronograma!$B$2:$B${len(PUESTOS) + 1}=FALSE))>0'
                        cond2 = f'({dia_vacio})*NOT({trabajo_manana_diurno})'
                        condiciones_descanso.append(cond2)
                    
                    # IMPORTANTE: Un día es descanso solo si cumple AL MENOS UNA de las condiciones aplicables
                    if condiciones_descanso:
                        if len(condiciones_descanso) == 1:
                            # Solo una condición: debe cumplirse (1 si TRUE, 0 si FALSE)
                            partes_bloque.append(f'({condiciones_descanso[0]})')
                        else:
                            # Dos condiciones: debe cumplirse AL MENOS UNA (OR)
                            partes_bloque.append(f'(({condiciones_descanso[0]})+({condiciones_descanso[1]})>0)')
                    else:
                        # No hay condiciones aplicables: NO es descanso
                        partes_bloque.append('0')
                
                if partes_bloque:
                    bloques_formula.append(f'SUM({",".join(partes_bloque)})')
            
            # Combinar todos los bloques
            if len(bloques_formula) == 0:
                # Si no hay bloques (caso muy raro), usar fórmula por defecto
                formula_final = '=0'
                if index == 0:
                    print(f"DEBUG - No hay bloques para {empleado['nombre']}, usando =0")
            elif len(bloques_formula) == 1:
                formula_final = f'={bloques_formula[0]}'
            else:
                formula_final = f'=SUM({",".join(bloques_formula)})'
            
            # Debug: imprimir información sobre la fórmula inicial
            if index == 0:
                print(f"DEBUG - Empleado: {empleado['nombre']}, Fila: {fila_actual}")
                print(f"DEBUG - Bloques: {len(bloques_formula)}, Tamaño fórmula inicial: {len(formula_final)}")
                if len(formula_final) < 500:
                    print(f"DEBUG - Fórmula inicial completa: {formula_final}")
                else:
                    print(f"DEBUG - Fórmula inicial (primeros 300 chars): {formula_final[:300]}...")
            
            # Escribir la fórmula - SIEMPRE usar fórmula, nunca valor de Python
            # Si la fórmula es muy larga, dividir en más bloques
            max_intentos = 3
            tamano_bloque_actual = tamano_bloque
            intento = 0
            formula_exitosa = False
            
            # Intentar escribir la fórmula inicial si es válida y no necesita reconstrucción
            if formula_final and len(formula_final) > 0 and len(formula_final) < 8000:
                try:
                    worksheet.write_formula(f'D{fila_actual}', formula_final)
                    formula_exitosa = True
                    if index == 0:
                        print(f"DEBUG - Fórmula inicial escrita exitosamente (sin reconstrucción)")
                except Exception as e:
                    if index == 0:
                        print(f"DEBUG - Error al escribir fórmula inicial: {e}")
                    # Continuar al bucle de intentos para reconstruir
            
            while intento < max_intentos and not formula_exitosa:
                try:
                    # Si la fórmula excede el límite o está vacía, reducir tamaño de bloque y reconstruir
                    if not formula_final or len(formula_final) >= 8000:
                        if index == 0:
                            print(f"DEBUG - Intento {intento + 1}: Fórmula muy larga ({len(formula_final) if formula_final else 0} chars), reconstruyendo con bloque de {tamano_bloque_actual}")
                        # Reducir tamaño de bloque progresivamente
                        if intento == 0:
                            tamano_bloque_actual = 10
                        elif intento == 1:
                            tamano_bloque_actual = 5
                        else:
                            tamano_bloque_actual = 3
                        bloques_formula_nuevos = []
                        
                        for bloque_inicio in range(0, num_dias, tamano_bloque_actual):
                            bloque_fin = min(bloque_inicio + tamano_bloque_actual, num_dias)
                            partes_bloque = []
                            
                            for dia_idx in range(bloque_inicio, bloque_fin):
                                col_dia_cronograma = get_excel_column_name(3 + dia_idx)
                                # Versión más compacta: usar SUMPRODUCT para verificar si NO trabaja ese día
                                dia_vacio = f'SUMPRODUCT((Cronograma!{col_dia_cronograma}$2:{col_dia_cronograma}${len(PUESTOS) + 1}=$A{fila_actual})*1)=0'
                                
                                condiciones_descanso = []
                                
                                # Condición 1: Solo aplicar si hay día anterior (dia_idx > 0)
                                if dia_idx > 0:
                                    col_anterior = get_excel_column_name(3 + dia_idx - 1)
                                    trabajo_ayer_nocturno = f'SUMPRODUCT((Cronograma!{col_anterior}$2:{col_anterior}${len(PUESTOS) + 1}=$A{fila_actual})*(Cronograma!$B$2:$B${len(PUESTOS) + 1}=TRUE))>0'
                                    cond1 = f'({dia_vacio})*NOT({trabajo_ayer_nocturno})'
                                    condiciones_descanso.append(cond1)
                                
                                # Condición 2: Solo aplicar si hay día siguiente (dia_idx < num_dias - 1)
                                if dia_idx < num_dias - 1:
                                    col_siguiente = get_excel_column_name(3 + dia_idx + 1)
                                    trabajo_manana_diurno = f'SUMPRODUCT((Cronograma!{col_siguiente}$2:{col_siguiente}${len(PUESTOS) + 1}=$A{fila_actual})*(Cronograma!$B$2:$B${len(PUESTOS) + 1}=FALSE))>0'
                                    cond2 = f'({dia_vacio})*NOT({trabajo_manana_diurno})'
                                    condiciones_descanso.append(cond2)
                                
                                if condiciones_descanso:
                                    if len(condiciones_descanso) == 1:
                                        partes_bloque.append(f'({condiciones_descanso[0]})')
                                    else:
                                        partes_bloque.append(f'(({condiciones_descanso[0]})+({condiciones_descanso[1]})>0)')
                                else:
                                    # No hay condiciones aplicables: NO es descanso
                                    partes_bloque.append('0')
                            
                            if partes_bloque:
                                bloques_formula_nuevos.append(f'SUM({",".join(partes_bloque)})')
                        
                        # Reconstruir formula_final
                        if len(bloques_formula_nuevos) == 0:
                            formula_final = '=0'
                            if index == 0:
                                print(f"DEBUG - No se generaron bloques nuevos, usando =0")
                        elif len(bloques_formula_nuevos) == 1:
                            formula_final = f'={bloques_formula_nuevos[0]}'
                        else:
                            formula_final = f'=SUM({",".join(bloques_formula_nuevos)})'
                        
                        if index == 0:
                            print(f"DEBUG - Fórmula reconstruida: {len(bloques_formula_nuevos)} bloques, tamaño: {len(formula_final)} chars")
                    
                    # Verificar que formula_final no esté vacío antes de escribir
                    if formula_final and len(formula_final) > 0 and len(formula_final) < 8192:  # Límite de Excel
                        # Debug: imprimir primera fórmula para verificar
                        if index == 0:
                            print(f"DEBUG - Escribiendo fórmula para {empleado['nombre']} (fila {fila_actual}):")
                            print(f"DEBUG - Longitud: {len(formula_final)}")
                            print(f"DEBUG - Primeros 500 chars: {formula_final[:500]}")
                        try:
                            worksheet.write_formula(f'D{fila_actual}', formula_final)
                            formula_exitosa = True
                            if index == 0:
                                print(f"DEBUG - Fórmula escrita exitosamente en D{fila_actual}")
                        except Exception as e:
                            if index == 0:
                                print(f"DEBUG - Error al escribir fórmula: {e}")
                            raise
                    elif not formula_final or len(formula_final) == 0:
                        # Si está vacío, usar fórmula por defecto
                        formula_final = '=0'
                        worksheet.write_formula(f'D{fila_actual}', formula_final)
                        formula_exitosa = True
                    else:
                        # Si aún es muy larga, reducir más el bloque
                        intento += 1
                        tamano_bloque_actual = max(5, tamano_bloque_actual // 2)
                        formula_final = ""  # Forzar reconstrucción en la siguiente iteración
                        
                except Exception as e:
                    if index == 0:
                        print(f"DEBUG - Excepción en intento {intento + 1}: {e}")
                    intento += 1
                    tamano_bloque_actual = max(5, tamano_bloque_actual // 2)
                    formula_final = ""  # Forzar reconstrucción en la siguiente iteración
            
            # Si después de todos los intentos falla, usar fórmula mínima pero SIEMPRE fórmula
            if not formula_exitosa:
                if index == 0:
                    print(f"DEBUG - No se pudo escribir fórmula después de {max_intentos} intentos, usando =0")
                try:
                    worksheet.write_formula(f'D{fila_actual}', '=0')
                except:
                    # Último recurso absoluto: fórmula que siempre retorna 0
                    if index == 0:
                        print(f"DEBUG - Error al escribir =0, usando =IF(TRUE,0,0)")
                    worksheet.write_formula(f'D{fila_actual}', '=IF(TRUE,0,0)')
        
        # Formato condicional para colorear celdas según tipo de turno
        # Amarillo: turnos diurnos
        worksheet.conditional_format(f'E2:{get_excel_column_name(len(cronograma) + 4)}{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=VLOOKUP(E2, Cronograma!$A$2:$B${len(PUESTOS) + 1}, 2, FALSE)=FALSE',
            'format':   format_yellow
        })
        worksheet.conditional_format(f'B2:B{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=TRUE',
            'format':   format_yellow
        })
        
        # Azul: turnos nocturnos
        worksheet.conditional_format(f'E2:{get_excel_column_name(len(cronograma) + 4)}{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=VLOOKUP(E2, Cronograma!$A$2:$B${len(PUESTOS) + 1}, 2, FALSE)=TRUE',
            'format':   format_blue
        })
        worksheet.conditional_format(f'C2:C{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=TRUE',
            'format':   format_blue
        })
        
        # Verde: descansos (celdas vacías que cumplen las condiciones de descanso)
        # La fórmula verifica para cada día si es descanso según las nuevas reglas
        num_dias = len(cronograma)
        for dia_idx in range(num_dias):
            col_dia = get_excel_column_name(5 + dia_idx)
            
            # Construir condición para formato verde (similar a la fórmula de descansos)
            # En formato condicional, las referencias son relativas a la celda que se está evaluando
            # Usar referencias relativas para que funcione en todas las filas
            
            # Verificar si el empleado NO trabajó ese día
            # Buscar directamente en Cronograma si el empleado está asignado a algún puesto ese día
            # Usar referencia relativa para el nombre del empleado (columna A, fila relativa)
            nombre_emp_ref = '$A2'  # Referencia mixta: columna absoluta, fila relativa
            col_dia_cronograma = get_excel_column_name(3 + dia_idx)  # Columna del día en hoja Cronograma
            
            # Verificar si el empleado trabajó ese día (buscar su nombre en la columna del día en Cronograma)
            dia_vacio = f'ISERROR(MATCH({nombre_emp_ref}, Cronograma!{col_dia_cronograma}$2:{col_dia_cronograma}${len(PUESTOS) + 1}, 0))'
            
            condiciones_verde = []
            
            # Condición 1: Día vacío Y día anterior no trabajó nocturno
            # NO aplicar al primer día (dia_idx > 0)
            if dia_idx > 0:
                col_anterior_cronograma = get_excel_column_name(3 + dia_idx - 1)
                # Verificar si día anterior trabajó nocturno usando INDEX/MATCH
                trabajo_ayer_nocturno = f'IF(ISERROR(MATCH({nombre_emp_ref},Cronograma!{col_anterior_cronograma}$2:{col_anterior_cronograma}${len(PUESTOS) + 1},0)),FALSE,INDEX(Cronograma!$B$2:$B${len(PUESTOS) + 1},MATCH({nombre_emp_ref},Cronograma!{col_anterior_cronograma}$2:{col_anterior_cronograma}${len(PUESTOS) + 1},0))=TRUE)'
                cond1 = f'AND({dia_vacio}, NOT({trabajo_ayer_nocturno}))'
                condiciones_verde.append(cond1)
            
            # Condición 2: Día vacío Y día siguiente no trabaja diurno
            # NO aplicar al último día (dia_idx < num_dias - 1)
            if dia_idx < num_dias - 1:
                col_siguiente_cronograma = get_excel_column_name(3 + dia_idx + 1)
                # Verificar si día siguiente trabaja diurno usando INDEX/MATCH
                trabajo_manana_diurno = f'IF(ISERROR(MATCH({nombre_emp_ref},Cronograma!{col_siguiente_cronograma}$2:{col_siguiente_cronograma}${len(PUESTOS) + 1},0)),FALSE,INDEX(Cronograma!$B$2:$B${len(PUESTOS) + 1},MATCH({nombre_emp_ref},Cronograma!{col_siguiente_cronograma}$2:{col_siguiente_cronograma}${len(PUESTOS) + 1},0))=FALSE)'
                cond2 = f'AND({dia_vacio}, NOT({trabajo_manana_diurno}))'
                condiciones_verde.append(cond2)
            
            # Si cumple cualquiera de las condiciones aplicables, es descanso (verde)
            if condiciones_verde:
                cond_verde = f'OR({",".join(condiciones_verde)})'
            else:
                # Si no hay condiciones aplicables, no es descanso
                cond_verde = 'FALSE'
            
            worksheet.conditional_format(f'{col_dia}2:{col_dia}{len(empleados) + 1}', {
                'type':     'formula',
                'criteria': cond_verde,
                'format':   format_green
            })
        
        # Columna de descansos también en verde
        worksheet.conditional_format(f'D2:D{len(empleados) + 1}', {
            'type':     'formula',
            'criteria': f'=TRUE',
            'format':   format_green
        })
        
        worksheet.autofit()
        worksheet.protect(options={'format_columns': True, 'format_rows': True})