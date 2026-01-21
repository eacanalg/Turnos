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

def actualizar_descansos (empleados, dia, dia_anterior, dia_siguiente, PUESTOS):
    """
    Actualiza descansos según la nueva lógica:
    - Un día se considera descanso si: ese día no hay turnos asignados Y el día anterior no se trabajó nocturno
    - También será descanso si: ese día no se trabajó Y al día siguiente no se trabaja diurno
    """
    actualizados = []
    puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
    puestos_diurnos = [item for item in PUESTOS if not item["nocturno"]]
    
    # Obtener empleados asignados a puestos del día actual
    empleados_dia = [dia[puesto["nombre"]] for puesto in PUESTOS if dia.get(puesto["nombre"]) is not None]
    
    # Obtener empleados que trabajaron en puestos nocturnos el día anterior
    empleados_dia_anterior_nocturno = []
    if dia_anterior:
        empleados_dia_anterior_nocturno = [dia_anterior[puesto["nombre"]] for puesto in puestos_nocturnos if dia_anterior.get(puesto["nombre"]) is not None]
    
    # Obtener empleados que trabajarán en puestos diurnos el día siguiente
    empleados_dia_siguiente_diurno = []
    if dia_siguiente:
        empleados_dia_siguiente_diurno = [dia_siguiente[puesto["nombre"]] for puesto in puestos_diurnos if dia_siguiente.get(puesto["nombre"]) is not None]
    
    for empleado in empleados:
        objeto_modificado = empleado.copy()
        nombre = empleado["nombre"]
        trabajo_hoy = nombre in empleados_dia
        trabajo_ayer_nocturno = nombre in empleados_dia_anterior_nocturno if dia_anterior else False
        trabajo_manana_diurno = nombre in empleados_dia_siguiente_diurno if dia_siguiente else False
        
        if trabajo_hoy:
            # El empleado trabajó hoy: no descansa, incrementa días sin descanso
            objeto_modificado['dias_sin_descanso'] += 1
        else:
            # El empleado no trabajó hoy: verificar si es descanso según las nuevas reglas
            es_descanso = False
            
            # Condición 1: Ese día no hay turnos asignados Y el día anterior no se trabajó nocturno
            # SOLO aplicar si hay día anterior (no es el primer día)
            if dia_anterior:
                condicion1 = not trabajo_ayer_nocturno
                if condicion1:
                    es_descanso = True
            
            # Condición 2: Ese día no se trabajó Y al día siguiente no se trabaja diurno
            # SOLO aplicar si hay día siguiente (no es el último día)
            if dia_siguiente:
                condicion2 = not trabajo_manana_diurno
                if condicion2:
                    es_descanso = True
            
            if es_descanso:
                objeto_modificado["descansos"] += 1
                objeto_modificado['dias_sin_descanso'] = 0
            else:
                # No es descanso según las reglas, pero incrementa días sin descanso
                objeto_modificado['dias_sin_descanso'] += 1
        
        actualizados.append(objeto_modificado)
    
    return actualizados

def formatear_fecha (fecha):
    return fecha.date()

def calcular_peso_persona (empleado, puesto, dias, promedio_turnos=None, ajuste_equidad=None, promedio_descansos=None, desviacion_descansos=None):
    # Peso base MUY alto para incentivar asignaciones (turno vacío es última opción)
    # Debe ser lo suficientemente alto para que cualquier asignación sea mejor que vacío
    peso_base = 10000
    dias_trabajados = empleado['turnos_dia'] + empleado['turnos_noche']
    nocturno = puesto['nocturno']
    
    # Incentivo base para asignar (siempre positivo para preferir asignación sobre vacío)
    incentivo_asignacion = 1000
    
    # PENALIZACIÓN CRÍTICA: Si nunca descansa (descansos = 0 y tiene turnos)
    # Esto es prioritario: evitar que alguien trabaje todos los días
    penalizacion_sin_descansos = 0
    if dias_trabajados > 0 and empleado['descansos'] == 0:
        # Nunca ha descansado: PENALIZACIÓN MUY ALTA para forzar descanso
        penalizacion_sin_descansos = -50000  # Penalización extrema
        # Penalización adicional por cada día trabajado sin descanso
        penalizacion_sin_descansos -= dias_trabajados * 10000
    
    # INCENTIVO MUY FUERTE: Si tiene más descansos que turnos trabajados
    # Este es un caso extremo de desequilibrio - debe trabajar MÁS, no menos
    # A MÁS DESCANSOS, MÁS DEBE SER ELEGIBLE PARA TRABAJAR
    incentivo_descansos_vs_turnos = 0
    if empleado['descansos'] > dias_trabajados:
        # Tiene más descansos que turnos: INCENTIVO MUY ALTO para trabajar
        diferencia_absoluta = empleado['descansos'] - dias_trabajados
        # Incentivo exponencial: cada descanso extra sobre turnos trabajados se incentiva fuertemente
        incentivo_descansos_vs_turnos = diferencia_absoluta * 20000  # MUY alto incentivo
        # Incentivo adicional si la diferencia es grande
        if diferencia_absoluta > 2:
            incentivo_descansos_vs_turnos += (diferencia_absoluta - 2) * 30000  # Incentivo extremo
        # Si tiene más del doble de descansos que turnos, incentivo aún mayor
        if empleado['descansos'] > dias_trabajados * 2:
            incentivo_descansos_vs_turnos += (empleado['descansos'] - dias_trabajados * 2) * 50000
    
    # Proporción ideal: 5 días trabajo / 2 días descanso = ratio 0.4 (2/5)
    # Si tiene muchos descansos, incentivar MUY fuertemente asignación
    # Penalizar MUY fuertemente el exceso de descansos
    if (dias_trabajados == 0):
        # Empleado sin turnos: MUY alta prioridad para asignar
        variable_por_descansos = 5000
    else:
        ratio_descansos = empleado['descansos'] / dias_trabajados
        ratio_ideal = 0.4  # 2 descansos por 5 trabajos
        
        if ratio_descansos > ratio_ideal:
            # Tiene demasiados descansos: incentivar MUY fuertemente asignación
            # Aumentar exponencialmente con el exceso de descansos
            exceso_descansos = ratio_descansos - ratio_ideal
            variable_por_descansos = exceso_descansos * 5000  # Aumentado de 2000 a 5000
            # Bonus adicional si el exceso es muy grande
            if exceso_descansos > 0.2:
                variable_por_descansos += (exceso_descansos - 0.2) * 10000  # Aumentado de 3000 a 10000
            # Bonus adicional por número absoluto de descansos excesivos
            descansos_excesivos = empleado['descansos'] - (dias_trabajados * ratio_ideal)
            if descansos_excesivos > 2:
                variable_por_descansos += descansos_excesivos * 3000  # Penalización adicional por cada descanso extra
        else:
            # Tiene pocos descansos: desincentivar (pero menos que antes para no penalizar demasiado)
            variable_por_descansos = (ratio_descansos - ratio_ideal) * 100
    
    # Penalización por trabajar más de 5 días seguidos (ALTA PENALIZACIÓN)
    # También penalizar si tiene muchos días seguidos mientras otros descansan
    dias_sin_descanso = empleado['dias_sin_descanso']
    if dias_sin_descanso >= 6:
        # Ya trabajó más de 5 días seguidos: penalización MUY alta
        dias_extra = dias_sin_descanso - 5  # Días por encima del límite de 5
        variable_por_trabajos = -dias_extra * 5000  # Penalización MUY alta
    elif dias_sin_descanso == 5:
        # Está en el límite (5 días): penalización MUY alta para evitar el día 6
        variable_por_trabajos = -3000
    elif dias_sin_descanso == 4:
        # Está cerca del límite: desincentivar fuertemente
        variable_por_trabajos = -500
    elif dias_sin_descanso == 3:
        # Está trabajando varios días: desincentivar ligeramente si otros descansan mucho
        variable_por_trabajos = -100
    else:
        # Puede trabajar más días: incentivar si tiene pocos turnos
        variable_por_trabajos = max(0, (4 - dias_sin_descanso) * 10)
    
    # Variable por disparidad entre jornadas (incentivar balance) - REDUCIDA para priorizar descansos
    if nocturno:
        variable_jornadas = (empleado['turnos_dia'] - empleado['turnos_noche']) * 0.5
    else:
        variable_jornadas = (empleado['turnos_noche'] - empleado['turnos_dia']) * 0.5
    
    # Penalización adicional si la proporción trabajo/descanso está muy desbalanceada
    if dias_trabajados > 0:
        ratio_actual = empleado['descansos'] / dias_trabajados
        desbalance = abs(ratio_actual - 0.4)
        if desbalance > 0.2:  # Si está muy desbalanceado (más de 20% de diferencia)
            penalizacion_desbalance = -desbalance * 200
        else:
            penalizacion_desbalance = 0
    else:
        penalizacion_desbalance = 0
    
    # Incentivo por equidad: si el empleado tiene menos turnos que el promedio, incentivar asignación
    # Penalización fuerte si tiene más turnos que el promedio
    # También considerar descansos en la equidad
    incentivo_equidad = 0
    if promedio_turnos is not None:
        diferencia_turnos = promedio_turnos - dias_trabajados
        
        # Calcular promedio de descansos para equidad
        # (esto se puede mejorar calculando el promedio real de descansos)
        promedio_descansos_estimado = promedio_turnos * 0.4  # Ratio ideal
        diferencia_descansos = promedio_descansos_estimado - empleado['descansos']
        
        if diferencia_turnos > 0.1:  # Tiene al menos 0.1 turnos menos que el promedio
            # Tiene menos turnos que el promedio: incentivar MUY fuertemente
            # Aumentar el incentivo exponencialmente con la diferencia
            incentivo_equidad = diferencia_turnos * 5000  # Aumentado de 2000 a 5000
            # Bonus adicional si la diferencia es grande
            if diferencia_turnos > 1:
                incentivo_equidad += (diferencia_turnos - 1) * 10000  # Aumentado de 3000 a 10000
            # Bonus adicional por descansos excesivos
            if diferencia_descansos < -2:  # Tiene más de 2 descansos que el promedio
                incentivo_equidad += abs(diferencia_descansos) * 5000
            # BONUS EXTREMO: Si tiene más descansos que turnos trabajados
            if empleado['descansos'] > dias_trabajados:
                diferencia_extrema = empleado['descansos'] - dias_trabajados
                incentivo_equidad += diferencia_extrema * 25000  # Incentivo MUY alto para trabajar
        elif diferencia_turnos < -0.1:
            # Tiene más turnos que el promedio: penalizar MUY fuertemente
            # Penalización exponencial para evitar desequilibrios
            incentivo_equidad = diferencia_turnos * 5000  # Aumentado de 3000 a 5000
            # Penalización adicional si la diferencia es grande
            if diferencia_turnos < -1:
                incentivo_equidad += (diferencia_turnos + 1) * 10000  # Aumentado de 5000 a 10000
        elif dias_trabajados == 0 and promedio_turnos == 0:
            # Al inicio cuando todos tienen 0, dar incentivo base para equidad
            incentivo_equidad = 1000
        elif dias_trabajados == 0 and promedio_turnos > 0:
            # Si no ha trabajado nada pero otros sí: incentivo muy alto
            incentivo_equidad = promedio_turnos * 8000  # Aumentado de 3000 a 8000
        elif diferencia_descansos < -3:  # Tiene más de 3 descansos que el promedio estimado
            # Penalización adicional por exceso de descansos incluso si los turnos están balanceados
            incentivo_equidad += abs(diferencia_descansos) * 4000
    
    # Ajuste basado en promedio y desviación de descansos (minimizar ambos)
    # PRIORIDAD ALTA: Asegurar que todos descansen la misma cantidad de días
    ajuste_descansos_balance = 0
    if promedio_descansos is not None and desviacion_descansos is not None:
        # Si el empleado tiene menos descansos que el promedio, incentivar asignación
        # (para que otros descansen más y reduzcamos el promedio)
        diferencia_descansos_promedio = promedio_descansos - empleado['descansos']
        
        if diferencia_descansos_promedio > 0:
            # Tiene menos descansos que el promedio: incentivar trabajar para que otros descansen más
            # Esto ayuda a reducir el promedio total de descansos y lograr equidad
            ajuste_descansos_balance = diferencia_descansos_promedio * 8000  # Aumentado de 2000 a 8000
            # Bonus exponencial si la diferencia es grande
            if diferencia_descansos_promedio > 1:
                ajuste_descansos_balance += (diferencia_descansos_promedio - 1) * 12000
        else:
            # Tiene más descansos que el promedio: INCENTIVAR MUY FUERTEMENTE trabajar
            # A MÁS DESCANSOS, MÁS DEBE TRABAJAR para balancear
            diferencia_absoluta = abs(diferencia_descansos_promedio)
            ajuste_descansos_balance = diferencia_absoluta * 10000  # INCENTIVO alto (antes era penalización)
            # Incentivo exponencial si la diferencia es grande
            if diferencia_absoluta > 1:
                ajuste_descansos_balance += (diferencia_absoluta - 1) * 15000  # Incentivo extremo
        
        # INCENTIVO por estar lejos del promedio si tiene muchos descansos
        # Si tiene muchos descansos y está lejos del promedio, incentivar trabajar más
        if diferencia_descansos_promedio < 0 and abs(diferencia_descansos_promedio) > desviacion_descansos:
            # Tiene más descansos que el promedio y está muy lejos: incentivar trabajar
            ajuste_descansos_balance += abs(diferencia_descansos_promedio - desviacion_descansos) * 10000
        
        # Incentivo adicional si tiene muchos descansos (equidad estricta)
        # A más descansos, más incentivo para trabajar
        if diferencia_descansos_promedio < -0.5:  # Si tiene más de medio día de descansos que el promedio
            ajuste_descansos_balance += abs(diferencia_descansos_promedio) * 8000
    
    # Ajuste adicional basado en desequilibrios de iteraciones anteriores
    # PRIORIDAD ALTA: Corregir desequilibrios de descansos detectados en iteraciones previas
    ajuste_iterativo = 0
    if ajuste_equidad is not None and empleado['nombre'] in ajuste_equidad:
        # Ajustar peso basándose en desequilibrios previos
        ajuste_info = ajuste_equidad[empleado['nombre']]
        diferencia_descansos_abs = ajuste_info.get('diferencia_descansos_abs', 0)
        
        # Si tenía muchos descansos o pocos turnos, aumentar incentivo MUY fuertemente
        if ajuste_info.get('muchos_descansos', False) or ajuste_info.get('pocos_turnos', False):
            ajuste_iterativo = ajuste_info.get('factor_ajuste', 0) * 5000  # Aumentado de 1000 a 5000
            # Bonus adicional si el problema es específicamente de descansos
            if ajuste_info.get('muchos_descansos', False):
                ajuste_iterativo += diferencia_descansos_abs * 8000  # Usar diferencia absoluta para ajuste más preciso
        # Si tenía pocos descansos o muchos turnos, reducir incentivo MUY fuertemente
        elif ajuste_info.get('pocos_descansos', False) or ajuste_info.get('muchos_turnos', False):
            ajuste_iterativo = -ajuste_info.get('factor_ajuste', 0) * 4000  # Aumentado de 500 a 4000
            # Penalización adicional si el problema es específicamente de descansos
            if ajuste_info.get('pocos_descansos', False):
                ajuste_iterativo -= diferencia_descansos_abs * 6000  # Usar diferencia absoluta para ajuste más preciso
    
    peso_total = (peso_base + incentivo_asignacion + variable_por_descansos + 
                   variable_por_trabajos + variable_jornadas + penalizacion_desbalance + 
                   incentivo_equidad + incentivo_descansos_vs_turnos + ajuste_iterativo + 
                   penalizacion_sin_descansos + ajuste_descansos_balance)
    
    return peso_total

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

    # Precalcular bloqueos estáticos para todos los días
    bloqueos_por_dia = []
    for i in range(len(cronograma)):
        dia = cronograma[i]
        bloqueos_dia = []
        bloqueos_noche = []
        
        # Bloqueos por cronograma
        for empleado in empleados:
            if dia['fecha'] in empleado['bloqueos_dia']:
                bloqueos_dia.append(empleado["nombre"])
            if dia['fecha'] in empleado['bloqueos_noche']:
                bloqueos_noche.append(empleado["nombre"])
        
        bloqueos_por_dia.append({
            'bloqueos_dia': bloqueos_dia,
            'bloqueos_noche': bloqueos_noche
        })
    
    # Sistema iterativo para mejorar equidad
    max_iteraciones_equidad = 3
    mejor_solucion = None
    mejor_equidad_score = float('inf')
    desequilibrios_previos = {}  # Guardar desequilibrios de iteraciones anteriores
    
    for iteracion_equidad in range(max_iteraciones_equidad):
        if iteracion_equidad > 0:
            
            # Reiniciar estados para nueva iteración
            empleados = [crear_clase_empleado(e) for e in EMPLEADOS]
            cronograma = [cronograma_diario_vacio(dia, PUESTOS) for dia in dias]
            
            # Aplicar ajustes de pesos basados en desequilibrios previos
            # Esto se hará modificando los pesos en calcular_peso_persona
            # mediante un factor de ajuste por empleado
        else:
            # Primera iteración: empezar desde cero
            empleados = [crear_clase_empleado(e) for e in EMPLEADOS]
            cronograma = [cronograma_diario_vacio(dia, PUESTOS) for dia in dias]
        
        # ========== CONSTRUIR GRAFO GLOBAL COMPLETO ==========
        G_global = nx.Graph()
        
        # Crear nodos para todos los (puesto, dia)
        for i in range(len(cronograma)):
            for puesto in PUESTOS:
                nodo_puesto_dia = (puesto['nombre'], i)
                G_global.add_node(nodo_puesto_dia, bipartite=0, puesto=puesto['nombre'], dia=i, nocturno=puesto['nocturno'])
        
        # Crear nodos para todos los empleados
        for empleado in empleados:
            G_global.add_node(empleado['nombre'], bipartite=1)
        
        # Calcular promedio inicial para equidad (se actualizará iterativamente)
        total_turnos = sum(e['turnos_dia'] + e['turnos_noche'] for e in empleados)
        promedio_turnos = total_turnos / len(empleados) if len(empleados) > 0 else 0
        
        # Calcular promedio y desviación de descansos para balance
        total_descansos = sum(e['descansos'] for e in empleados)
        promedio_descansos = total_descansos / len(empleados) if len(empleados) > 0 else 0
        # Calcular desviación estándar de descansos
        if len(empleados) > 0:
            varianza_descansos = sum((e['descansos'] - promedio_descansos) ** 2 for e in empleados) / len(empleados)
            desviacion_descansos = varianza_descansos ** 0.5
        else:
            desviacion_descansos = 0
        
        # Agregar aristas considerando TODAS las restricciones
        # Para la restricción noche->día, usamos un enfoque donde bloqueamos aristas
        # que violarían esta restricción basándose en posibles asignaciones nocturnas previas
        for j in range(len(cronograma)):
            bloqueos_dia = bloqueos_por_dia[j]['bloqueos_dia'].copy()
            bloqueos_noche = bloqueos_por_dia[j]['bloqueos_noche'].copy()
            
            for puesto in PUESTOS:
                nodo_puesto_dia = (puesto['nombre'], j)
                es_nocturno = puesto["nocturno"]
                
                # Bloqueo por puesto
                bloqueos_puesto = [e['nombre'] for e in empleados if puesto['nombre'] not in e['puestos_habilitados']]
                
                # Para restricción noche->día: si es turno diurno (j > 0), 
                # no podemos crear aristas para empleados que podrían trabajar noche en j-1
                # Como no sabemos quién trabajará noche, bloqueamos preventivamente
                # solo si hay una alta probabilidad de conflicto
                
                for empleado in empleados:
                    bloqueado = False
                    
                    # Bloqueo por puesto
                    if empleado['nombre'] in bloqueos_puesto:
                        bloqueado = True
                    
                    # Bloqueos por cronograma (dia/noche)
                    if not bloqueado:
                        if es_nocturno:
                            if empleado['nombre'] in bloqueos_noche:
                                bloqueado = True
                        else:
                            if empleado['nombre'] in bloqueos_dia:
                                bloqueado = True
                    
                    # Restricción noche->día: Para turnos diurnos, no crear arista si
                    # el empleado podría estar trabajando noche el día anterior
                    # Como no sabemos las asignaciones previas, usamos un enfoque conservador:
                    # Solo bloqueamos si el empleado tiene alta probabilidad de trabajar noche anterior
                    # basado en sus puestos habilitados
                    if not bloqueado and not es_nocturno and j > 0:
                        # Verificar si el empleado puede trabajar en puestos nocturnos
                        puede_trabajar_noche = any(p['nombre'] in empleado['puestos_habilitados'] 
                                                   for p in PUESTOS if p['nocturno'])
                        # Si puede trabajar noche y no está bloqueado para noche el día anterior,
                        # hay riesgo de conflicto. Usaremos pesos muy negativos en lugar de bloquear
                        # para que el matching evite estas asignaciones
                        if puede_trabajar_noche:
                            # No bloqueamos, pero usaremos un peso muy bajo para desincentivar
                            # Esta asignación se verificará después del matching
                            pass  # Continuamos, pero el peso se ajustará
                    
                    # Agregar arista si no está bloqueado
                    if not bloqueado:
                        # Calcular peso inicial (se ajustará considerando restricciones dinámicas)
                        # Pasar ajuste_equidad, promedio_descansos y desviacion_descansos
                        peso = calcular_peso_persona(empleado, puesto, len(cronograma), promedio_turnos, 
                                                     desequilibrios_previos, promedio_descansos, desviacion_descansos)
                        
                        # Ajuste de peso para restricción noche->día potencial
                        # Si es turno diurno y el empleado puede trabajar noche, reducir peso ligeramente
                        if not es_nocturno and j > 0:
                            puede_trabajar_noche = any(p['nombre'] in empleado['puestos_habilitados'] 
                                                       for p in PUESTOS if p['nocturno'])
                            if puede_trabajar_noche:
                                # Reducir peso para desincentivar (pero no bloquear completamente)
                                peso = peso * 0.8
                        
                        G_global.add_edge(nodo_puesto_dia, empleado['nombre'], weight=peso)
        
        # ========== RESOLVER GRAFO UNA SOLA VEZ ==========
        matching_global = nx.algorithms.matching.max_weight_matching(G_global, maxcardinality=True)
        
        # Procesar matching completo
        asignaciones_completas = {}
        for u, v in matching_global:
            if isinstance(u, tuple):
                asignaciones_completas[u] = v
            elif isinstance(v, tuple):
                asignaciones_completas[v] = u
        
        
        # ========== PROCESAR ASIGNACIONES DÍA POR DÍA ==========
        for i in range(len(cronograma)):
            dia = cronograma[i]
            
            # Bloqueos dinámicos para este día
            bloqueos_dia = bloqueos_por_dia[i]['bloqueos_dia'].copy()
            bloqueos_noche = bloqueos_por_dia[i]['bloqueos_noche'].copy()
            
            if i > 0:
                dia_anterior = cronograma[i - 1]
                puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
                for p in puestos_nocturnos:
                    empleado_noche_anterior = dia_anterior[p["nombre"]]
                    if empleado_noche_anterior:
                        bloqueos_dia.append(empleado_noche_anterior)
            
            # Procesar asignaciones del matching para este día
            for puesto in PUESTOS:
                nodo_puesto_dia = (puesto['nombre'], i)
                empleado_asignado = asignaciones_completas.get(nodo_puesto_dia)
                es_nocturno = puesto["nocturno"]
                
                asignado_exitosamente = False
                
                if empleado_asignado:
                    # Verificar restricciones finales (especialmente noche->día)
                    bloqueos_puesto = [e['nombre'] for e in empleados if puesto['nombre'] not in e['puestos_habilitados']]
                    
                    bloqueado = False
                    if es_nocturno:
                        if empleado_asignado in bloqueos_noche + bloqueos_puesto:
                            bloqueado = True
                    else:
                        if empleado_asignado in bloqueos_dia + bloqueos_puesto:
                            bloqueado = True
                    
                    # Verificar restricción noche->día
                    if not bloqueado and not es_nocturno and i > 0:
                        dia_anterior = cronograma[i - 1]
                        puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
                        for p in puestos_nocturnos:
                            if dia_anterior[p["nombre"]] == empleado_asignado:
                                bloqueado = True
                                break
                    
                    if not bloqueado:
                        empleado_obj = next((e for e in empleados if e['nombre'] == empleado_asignado), None)
                        if empleado_obj:
                            empleados = actualizar_empleados(empleados, empleado_asignado, es_nocturno)
                            dia[puesto["nombre"]] = empleado_asignado
                            asignado_exitosamente = True
                
                # Fallback si no se asignó
                if not asignado_exitosamente:
                    bloqueos_puesto = [e['nombre'] for e in empleados if puesto['nombre'] not in e['puestos_habilitados']]
                    
                    empleados_disponibles = []
                    if es_nocturno:
                        empleados_disponibles = [item for item in empleados 
                                               if item["nombre"] not in bloqueos_noche + bloqueos_puesto]
                    else:
                        empleados_disponibles = [item for item in empleados 
                                               if item["nombre"] not in bloqueos_dia + bloqueos_puesto]
                    
                    empleados_validos = []
                    for emp in empleados_disponibles:
                        if not es_nocturno and i > 0:
                            dia_anterior = cronograma[i - 1]
                            puestos_nocturnos = [item for item in PUESTOS if item["nocturno"]]
                            trabajo_noche_anterior = False
                            for p in puestos_nocturnos:
                                if dia_anterior[p["nombre"]] == emp["nombre"]:
                                    trabajo_noche_anterior = True
                                    break
                            if not trabajo_noche_anterior:
                                empleados_validos.append(emp)
                        else:
                            empleados_validos.append(emp)
                    
                    if empleados_validos:
                        for emp in empleados_validos:
                            emp['diferencia_promedio'] = promedio_turnos - (emp['turnos_dia'] + emp['turnos_noche'])
                        
                        # Recalcular promedio y desviación para fallback
                        total_descansos_fallback = sum(emp['descansos'] for emp in empleados)
                        promedio_descansos_fallback = total_descansos_fallback / len(empleados) if len(empleados) > 0 else 0
                        if len(empleados) > 0:
                            varianza_descansos_fallback = sum((emp['descansos'] - promedio_descansos_fallback) ** 2 for emp in empleados) / len(empleados)
                            desviacion_descansos_fallback = varianza_descansos_fallback ** 0.5
                        else:
                            desviacion_descansos_fallback = 0
                        
                        empleados_validos.sort(key=lambda e: (
                            -e['diferencia_promedio'],
                            e['turnos_dia'] + e['turnos_noche'],
                            e['dias_sin_descanso'],
                            -calcular_peso_persona(e, puesto, len(cronograma), promedio_turnos, 
                                                   desequilibrios_previos, promedio_descansos_fallback, desviacion_descansos_fallback)
                        ))
                        empleado_fallback = empleados_validos[0]
                        empleados = actualizar_empleados(empleados, empleado_fallback["nombre"], es_nocturno)
                        dia[puesto["nombre"]] = empleado_fallback["nombre"]
            
            cronograma[i] = dia
            
            if i > 0:
                # Obtener día siguiente si existe
                dia_siguiente = cronograma[i + 1] if i + 1 < len(cronograma) else None
                empleados = actualizar_descansos(empleados, dia, cronograma[i - 1], dia_siguiente, PUESTOS)
        
        # Al final de cada iteración, evaluar equidad y guardar mejor solución
        total_turnos_final = sum(e['turnos_dia'] + e['turnos_noche'] for e in empleados)
        promedio_final = total_turnos_final / len(empleados) if len(empleados) > 0 else 0
        total_descansos_final = sum(e['descansos'] for e in empleados)
        promedio_descansos_final = total_descansos_final / len(empleados) if len(empleados) > 0 else 0
        
        # Calcular score de equidad (menor es mejor)
        # Incluir: varianza de turnos, varianza de descansos, promedio de descansos, y desviación de descansos
        varianza_turnos = sum((e['turnos_dia'] + e['turnos_noche'] - promedio_final) ** 2 for e in empleados) / len(empleados) if len(empleados) > 0 else 0
        varianza_descansos = sum((e['descansos'] - promedio_descansos_final) ** 2 for e in empleados) / len(empleados) if len(empleados) > 0 else 0
        desviacion_descansos_final = varianza_descansos ** 0.5
        
        # Penalizar si hay empleados sin descansos
        empleados_sin_descansos = sum(1 for e in empleados if e['descansos'] == 0 and (e['turnos_dia'] + e['turnos_noche']) > 0)
        penalizacion_sin_descansos = empleados_sin_descansos * 1000  # Penalización alta por cada empleado sin descansos
        
        # Score: minimizar varianza, promedio de descansos, y desviación
        # PRIORIDAD ALTA: Equidad en descansos - dar más peso a varianza y desviación de descansos
        # Peso relativo: varianza de descansos es MUY importante para equidad
        equidad_score = (varianza_turnos + 
                         varianza_descansos * 3.0 +  # Aumentado de 1.0 a 3.0 - PRIORIDAD ALTA
                         promedio_descansos_final * 0.5 + 
                         desviacion_descansos_final * 2.0 +  # Aumentado de 0.3 a 2.0 - PRIORIDAD ALTA
                         penalizacion_sin_descansos)
        
        # Guardar mejor solución
        if mejor_solucion is None or equidad_score < mejor_equidad_score:
            mejor_solucion = {
                'empleados': [{'nombre': e['nombre'], 'turnos_dia': e['turnos_dia'], 
                              'turnos_noche': e['turnos_noche'], 'descansos': e['descansos'],
                              'dias_sin_descanso': e['dias_sin_descanso']} for e in empleados],
                'cronograma': [{**d} for d in cronograma]
            }
            mejor_equidad_score = equidad_score
        
        # Encontrar empleados con desequilibrios extremos para próxima iteración
        desequilibrios = []
        desequilibrios_previos = {}  # Reiniciar para nueva iteración
        
        for emp in empleados:
            turnos_emp = emp['turnos_dia'] + emp['turnos_noche']
            diferencia_turnos = promedio_final - turnos_emp  # Positivo = tiene menos turnos
            diferencia_descansos = emp['descansos'] - promedio_descansos_final  # Positivo = tiene más descansos
            
            # Verificar desequilibrio en turnos
            desequilibrio_turnos = abs(diferencia_turnos) > max(2, promedio_final * 0.25)
            # Verificar desequilibrio en descansos - MÁS ESTRICTO para lograr equidad
            # Cualquier diferencia mayor a 1 día o 20% del promedio se considera desequilibrio
            umbral_descansos = max(1, promedio_descansos_final * 0.2)  # Más estricto: 1 día o 20% (antes era 2 días o 40%)
            desequilibrio_descansos = abs(diferencia_descansos) > umbral_descansos
            
            if desequilibrio_turnos or desequilibrio_descansos:
                desequilibrios.append({
                    'empleado': emp,
                    'diferencia_turnos': abs(diferencia_turnos),
                    'diferencia_descansos': abs(diferencia_descansos),
                    'turnos': turnos_emp,
                    'descansos': emp['descansos']
                })
                
                # Guardar información para ajuste en próxima iteración
                # Dar más peso a desequilibrios de descansos para equidad
                factor_ajuste = max(abs(diferencia_turnos), abs(diferencia_descansos) * 1.5)  # Descansos tienen 1.5x más peso
                desequilibrios_previos[emp['nombre']] = {
                    'muchos_descansos': diferencia_descansos > 1,  # Más estricto: antes era > 2
                    'pocos_descansos': diferencia_descansos < -1,  # Más estricto: antes era < -2
                    'pocos_turnos': diferencia_turnos > 2,
                    'muchos_turnos': diferencia_turnos < -2,
                    'factor_ajuste': factor_ajuste,
                    'diferencia_descansos_abs': abs(diferencia_descansos)  # Guardar diferencia absoluta para ajustes más precisos
                }
        
        # Si hay desequilibrios significativos, continuar iterando
        # PRIORIDAD: Ser más estricto con desequilibrios de descansos
        if iteracion_equidad < max_iteraciones_equidad - 1:
            # Casos extremos: más estricto para descansos (diferencia > 2 días)
            casos_extremos_descansos = sum(1 for d in desequilibrios if d['diferencia_descansos'] > 2)
            casos_extremos_turnos = sum(1 for d in desequilibrios if d['diferencia_turnos'] > 4)
            casos_extremos = casos_extremos_descansos + casos_extremos_turnos
            
            # Continuar si hay muchos desequilibrios o casos extremos de descansos
            # Ser más estricto: continuar si hay desequilibrios de descansos incluso si son pocos
            desequilibrios_descansos = sum(1 for d in desequilibrios if abs(d['empleado']['descansos'] - promedio_descansos_final) > 1)
            
            if (len(desequilibrios) > len(empleados) * 0.15 or 
                casos_extremos > 0 or 
                desequilibrios_descansos > len(empleados) * 0.1):  # Continuar si >10% tiene desequilibrio de descansos
                continue
            else:
                # Equidad aceptable, salir del bucle
                break
    
    # Usar la mejor solución encontrada
    if mejor_solucion is not None:
        # Restaurar mejor solución
        for i, emp_sol in enumerate(mejor_solucion['empleados']):
            emp_actual = next((e for e in empleados if e['nombre'] == emp_sol['nombre']), None)
            if emp_actual:
                emp_actual['turnos_dia'] = emp_sol['turnos_dia']
                emp_actual['turnos_noche'] = emp_sol['turnos_noche']
                emp_actual['descansos'] = emp_sol['descansos']
                emp_actual['dias_sin_descanso'] = emp_sol['dias_sin_descanso']
        
        cronograma = mejor_solucion['cronograma']
        
        # Mostrar resumen de equidad en descansos
        total_descansos_final = sum(e['descansos'] for e in empleados)
        promedio_descansos_final = total_descansos_final / len(empleados) if len(empleados) > 0 else 0
        varianza_descansos_final = sum((e['descansos'] - promedio_descansos_final) ** 2 for e in empleados) / len(empleados) if len(empleados) > 0 else 0
        desviacion_descansos_final = varianza_descansos_final ** 0.5
        min_descansos = min(e['descansos'] for e in empleados)
        max_descansos = max(e['descansos'] for e in empleados)
        rango_descansos = max_descansos - min_descansos
        

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
        for index, empleado in enumerate(empleados):
            nombre_emp = empleado["nombre"]
            num_dias = len(cronograma)
            
            # Construir fórmula día por día usando SUM
            partes_formula = []
            for dia_idx in range(num_dias):
                # Verificar si el empleado NO trabajó ese día
                # Buscar directamente en Cronograma si el empleado está asignado a algún puesto ese día
                nombre_emp_excel = f'A{index+2}'  # Columna A tiene el nombre del empleado
                col_dia_cronograma = get_excel_column_name(3 + dia_idx)  # Columna del día en hoja Cronograma
                
                # Verificar si el empleado trabajó ese día (buscar su nombre en la columna del día en Cronograma)
                dia_vacio = f'ISERROR(MATCH({nombre_emp_excel}, Cronograma!{col_dia_cronograma}$2:{col_dia_cronograma}${len(PUESTOS) + 1}, 0))'
                
                condiciones_descanso = []
                
                # Condición 1: Día vacío Y día anterior no trabajó nocturno
                # NO aplicar al primer día (dia_idx > 0)
                if dia_idx > 0:
                    # Verificar si día anterior trabajó nocturno
                    # Buscar si el empleado está en algún puesto nocturno el día anterior en Cronograma
                    col_anterior_cronograma = get_excel_column_name(3 + dia_idx - 1)
                    # Usar SUMPRODUCT para verificar si el empleado trabajó en un puesto nocturno el día anterior
                    # Buscar el empleado en la columna del día anterior, y si lo encuentra, verificar si ese puesto es nocturno
                    trabajo_ayer_nocturno = f'IF(ISERROR(MATCH({nombre_emp_excel},Cronograma!{col_anterior_cronograma}$2:{col_anterior_cronograma}${len(PUESTOS) + 1},0)),FALSE,INDEX(Cronograma!$B$2:$B${len(PUESTOS) + 1},MATCH({nombre_emp_excel},Cronograma!{col_anterior_cronograma}$2:{col_anterior_cronograma}${len(PUESTOS) + 1},0))=TRUE)'
                    cond1 = f'AND({dia_vacio}, NOT({trabajo_ayer_nocturno}))'
                    condiciones_descanso.append(cond1)
                
                # Condición 2: Día vacío Y día siguiente no trabaja diurno
                # NO aplicar al último día (dia_idx < num_dias - 1)
                if dia_idx < num_dias - 1:
                    # Verificar si día siguiente trabaja diurno
                    # Buscar si el empleado está en algún puesto diurno el día siguiente en Cronograma
                    col_siguiente_cronograma = get_excel_column_name(3 + dia_idx + 1)
                    # Usar SUMPRODUCT para verificar si el empleado trabajará en un puesto diurno el día siguiente
                    trabajo_manana_diurno = f'IF(ISERROR(MATCH({nombre_emp_excel},Cronograma!{col_siguiente_cronograma}$2:{col_siguiente_cronograma}${len(PUESTOS) + 1},0)),FALSE,INDEX(Cronograma!$B$2:$B${len(PUESTOS) + 1},MATCH({nombre_emp_excel},Cronograma!{col_siguiente_cronograma}$2:{col_siguiente_cronograma}${len(PUESTOS) + 1},0))=FALSE)'
                    cond2 = f'AND({dia_vacio}, NOT({trabajo_manana_diurno}))'
                    condiciones_descanso.append(cond2)
                
                # Si cumple cualquiera de las condiciones aplicables, es descanso
                # Un día es descanso si cumple AL MENOS UNA de las dos condiciones
                if condiciones_descanso:
                    if len(condiciones_descanso) == 1:
                        # Solo una condición aplicable
                        partes_formula.append(f'IF({condiciones_descanso[0]}, 1, 0)')
                    else:
                        # Dos condiciones aplicables: usar OR
                        partes_formula.append(f'IF(OR({",".join(condiciones_descanso)}), 1, 0)')
                else:
                    # Si no hay condiciones aplicables (solo pasa si hay 1 día), no es descanso
                    partes_formula.append('0')
            
            # Sumar todos los días que son descanso
            formula_final = f'''=SUM({",".join(partes_formula)})'''
            
            
            # Escribir la fórmula
            try:
                worksheet.write_formula(f'D{index+2}', formula_final)
            except Exception as e:
                # Si falla, usar valor calculado en Python como respaldo
                worksheet.write(f'D{index+2}', empleado['descansos'])
        
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