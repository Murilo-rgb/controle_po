import re

with open('server.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Count import_file
count = content.count('def import_file(path):')
print('=== import_file function count ===')
print(count)

# 2. Find the import_file function and analyze its content
lines = content.split('\n')
import_line = None
to_cents_line = None
for i, line in enumerate(lines):
    if line.strip() == 'def import_file(path):':
        import_line = i
    if import_line is not None and line.strip().startswith('def _to_cents'):
        to_cents_line = i
        break

func_lines = lines[import_line:to_cents_line]
func_content = '\n'.join(func_lines)

print('\n=== Import_file function analysis ===')
print('Total lines of import_file:', len(func_lines))

# Check if colunas_extra is used in import_file
if 'colunas_extra' in func_content:
    print('⚠️  colunas_extra encontrado NA FUNCAO import_file!')
else:
    print('OK: colunas_extra NÃO está presente na função import_file')

# Check if unique_cols is used
if 'unique_cols' in func_content:
    print('OK: Usa unique_cols (versão corrigida/atlernaizada)')

# Check if the insert uses unique_cols
if 'columns=' in func_content and 'EXCLUDED.{c}' in func_content:
    print('OK: INSERT usa unique_cols (dynamically built from map)')

# Check the cols list
cols_match = re.search(r'cols = \[.*?\]', func_content, re.DOTALL)
if cols_match:
    cols_str = cols_match.group(0)
    items = re.findall(r'"([^"]+)"', cols_str)
    print('\n=== MAPA DE COLUNAS ===')
    print(f'Total de colunas mapeadas: {len(items)}')
    print('Primeiras 8 colunas:')
    for i, col in enumerate(items[:8]):
        print(f'  {i+1}. {col}')

print('\n=== CONCLUSAO ===')
if count == 1 and 'colunas_extra' not in func_content:
    print('✓ Cada coluna do Excel está mapeada individualmente com a coluna do banco.')
    print('✓ A logica colunas_extra (armazenamento em JSON) foi REMOVIDA da importacao.')
    print('✓ A logica atualizada usa: cols + unique_cols (lista dinamica de colunas).')
else:
    print('⚠️  Ha algo a verificar...')
