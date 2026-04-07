import re

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Znajdź funkcję send_formatted_mail
start_line = -1
for i, line in enumerate(lines):
    if line.strip().startswith('def send_formatted_mail'):
        start_line = i
        break

if start_line == -1:
    print('Function send_formatted_mail not found')
    sys.exit(1)

# Znajdź koniec funkcji (następna funkcja lub koniec pliku)
end_line = len(lines)
for i in range(start_line + 1, len(lines)):
    if lines[i].strip().startswith('def ') and i > start_line:
        end_line = i
        break

print(f'Function lines {start_line} to {end_line}')

# W obrębie tych linii szukamy except Exception as e: z błędnym wcięciem
for i in range(start_line, end_line):
    if 'except Exception as e:' in lines[i]:
        # Sprawdźmy kontekst: czy poprzednia linia zawiera 'try:' z wcięciem 8 spacji?
        # Poprawimy tylko jeśli jest w bloku if cache_data...
        # Po prostu poprawmy wcięcie na 8 spacji, a następną linię na 12
        # Ale upewnijmy się, że to ten fragment
        if i > 0 and 'cache_data' in lines[i-1]:
            print(f'Found target except at line {i}: {repr(lines[i])}')
            # Wcięcie powinno być 8 spacji (bo try ma 8)
            lines[i] = '        except Exception as e:\n'
            # Następna linia to log_event
            if i+1 < len(lines) and 'log_event("WARNING"' in lines[i+1]:
                lines[i+1] = '            log_event("WARNING", f"Błąd parsowania czasu cache: {e}")\n'
            print('Fixed')
            break

# Zapisz
with open('app.py', 'w', encoding='utf-8') as f:
    f.writelines(lines)