import glob
import pandas as pd
import os

files = glob.glob('Cronograma-*.xlsx')
if not files:
    print('No Cronograma file found')
    raise SystemExit(1)

# Prefer *_updated if present
files_updated = [f for f in files if '_updated' in f]
file = max(files_updated or files, key=os.path.getmtime)
print('Using', file)

df = pd.read_excel(file, sheet_name='Cronograma')
# Cronograma structure: first columns are Puestos, Nocturno, then dates
num_puestos = len(df)
num_cols = len(df.columns)

# Count empty cells excluding first two columns
date_cols = df.columns[2:]
empty_counts = df[date_cols].isna().sum()

total_empty = int(empty_counts.sum())
max_empty = int(empty_counts.max())
worst_days = [str(c.date()) if hasattr(c, 'date') else str(c) for c,v in empty_counts.items() if v==max_empty]

print('Total puestos:', num_puestos)
print('Total dias:', len(date_cols))
print('Total celdas sin asignar:', total_empty)
print('Dia(s) con más sin asignar (', max_empty, '):', ', '.join(worst_days))

# List puestos with most unassigned across all days
rows_empty = df[date_cols].isna().sum(axis=1)
rows = df.iloc[:,0].tolist()
max_row_empty = int(rows_empty.max())
worst_puestos = [rows[i] for i,v in enumerate(rows_empty) if v==max_row_empty]
print('Puesto(s) con más sin asignar (por puesto):', worst_puestos, '->', max_row_empty)

# Show top 10 puestos by unassigned
top = sorted([(rows[i], int(v)) for i,v in enumerate(rows_empty)], key=lambda x: -x[1])[:10]
print('\nTop 10 puestos por celdas sin asignar:')
for p,c in top:
    print(f'- {p}: {c}')
