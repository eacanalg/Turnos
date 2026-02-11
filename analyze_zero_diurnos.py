import glob
import pandas as pd
from collections import Counter

files = glob.glob('Cronograma-*.xlsx')
if not files:
    print('NO_EXCEL_FOUND')
    raise SystemExit(1)
# pick latest by modified time
files.sort(key=lambda p: __import__('os').path.getmtime(p), reverse=True)
path = files[0]
print('USING', path)

df = pd.read_excel(path, sheet_name='Cronograma')
# columns: Puestos, Nocturno, <dates...>
if 'Nocturno' not in df.columns:
    print('FORMAT_ERROR')
    raise SystemExit(1)

day_cols = [c for c in df.columns if c not in ('Puestos', 'Nocturno')]

# Count diurnal / nocturnal per employee
diurno_counts = Counter()
nocturno_counts = Counter()
all_names = set()
for c in day_cols:
    col = df[c]
    for idx, val in col.items():
        if pd.isna(val):
            continue
        name = str(val)
        all_names.add(name)
        if df.at[idx, 'Nocturno'] == True:
            nocturno_counts[name] += 1
        else:
            diurno_counts[name] += 1

# Ensure all employees from sheet Empleados also included if present
try:
    df_emp = pd.read_excel('Entradas.xlsm', sheet_name='Empleados', header=None, names=['nombre', 'puestos_habilitados'])
    for n in df_emp['nombre'].tolist():
        if isinstance(n, str):
            all_names.add(n)
            diurno_counts.setdefault(n, 0)
            nocturno_counts.setdefault(n, 0)
except Exception:
    pass

zero_diurnos = [n for n in sorted(all_names) if diurno_counts.get(n,0) == 0 and nocturno_counts.get(n,0) > 0]
zero_nocturnos = [n for n in sorted(all_names) if nocturno_counts.get(n,0) == 0 and diurno_counts.get(n,0) > 0]

print('\nEmployees with 0 diurnos (but >0 nocturnos):')
for n in zero_diurnos:
    print(f'  {n}: diurno=0, nocturno={nocturno_counts.get(n,0)}')

print('\nEmployees with 0 nocturnos (but >0 diurnos):')
for n in zero_nocturnos:
    print(f'  {n}: nocturno=0, diurno={diurno_counts.get(n,0)}')

print('\nSummary counts sample:')
for n in sorted(list(all_names))[:30]:
    print(f'  {n}: diurno={diurno_counts.get(n,0)}, nocturno={nocturno_counts.get(n,0)}')
