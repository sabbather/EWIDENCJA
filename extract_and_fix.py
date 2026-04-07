import re

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the function
start_line = -1
end_line = -1

for i, line in enumerate(lines):
    if line.strip().startswith('def send_formatted_mail'):
        start_line = i
        break

if start_line == -1:
    print('Function not found')
    exit(1)

# Find the end of function (next function or end of file)
for i in range(start_line + 1, len(lines)):
    if lines[i].strip().startswith('def ') and i > start_line:
        end_line = i
        break

if end_line == -1:
    end_line = len(lines)

print(f'Function from line {start_line} to {end_line}')

# Extract function lines
func_lines = lines[start_line:end_line]

# Fix the indentation issue
# Find the problematic except block
for i, line in enumerate(func_lines):
    if 'except Exception as e:' in line and line.count(' ') > 10:
        print(f'Found problematic except at relative line {i}: {repr(line)}')
        # Fix this line and the next line
        func_lines[i] = '        except Exception as e:\n'
        if i+1 < len(func_lines):
            func_lines[i+1] = '            log_event("WARNING", f"Błąd parsowania czasu cache: {e}")\n'
        break

# Also fix the "Jeśli cache nie jest ważny" comment indentation
# Look for the comment
for i, line in enumerate(func_lines):
    if 'Jeśli cache nie jest ważny' in line:
        print(f'Found comment at relative line {i}: {repr(line)}')
        # This should be indented with 4 spaces (same as surrounding code)
        # The line is: '    # Je榣i cache nie jest wa緉y, zaktualizuj go'
        # It should be: '    # Je渓i cache nie jest wa縩y, zaktualizuj go'
        # Actually just fix indentation to 4 spaces
        func_lines[i] = '    # Jeśli cache nie jest ważny, zaktualizuj go\n'
        break

# Fix the second "CZĘŚĆ 2: WYSYŁKA E-MAIL" comment
for i, line in enumerate(func_lines):
    if 'CZĘŚĆ 2: WYSYŁKA E-MAIL' in line and '#' in line:
        print(f'Found second CZĘŚĆ 2 at relative line {i}: {repr(line)}')
        # Should be 4 spaces indentation
        func_lines[i] = '    # --- CZĘŚĆ 2: WYSYŁKA E-MAIL ---\n'
        break

# Also fix the line with outlook = None (should be 4 spaces)
for i, line in enumerate(func_lines):
    if 'outlook = None' in line and '#' not in line:
        # Check if it's the right line (should be after the comment)
        if i > 0 and 'CZĘŚĆ 2: WYSYŁKA E-MAIL' in func_lines[i-1]:
            print(f'Found outlook = None at relative line {i}: {repr(line)}')
            func_lines[i] = '    outlook = None\n'
            break

# Write back
new_lines = lines[:start_line] + func_lines + lines[end_line:]

with open('app.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print('Function fixed')