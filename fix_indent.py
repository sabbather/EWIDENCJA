import sys

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

changed = False
for i, line in enumerate(lines):
    if 'except Exception as e:' in line and line.count(' ') > 10:
        print(f'Found at line {i}: {repr(line)}')
        lines[i] = '        except Exception as e:\n'
        if i+1 < len(lines) and 'log_event("WARNING"' in lines[i+1]:
            lines[i+1] = '            log_event("WARNING", f"Błąd parsowania czasu cache: {e}")\n'
        changed = True
        break

if changed:
    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    print('Fixed indentation')
else:
    print('No issues found')