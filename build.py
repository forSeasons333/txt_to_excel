import os
import subprocess
with open('resources.qrc', 'w') as f:
    f.write('<!DOCTYPE RCC><RCC version="1.0">\n<qresource>\n')
    for file in os.listdir('icons'):
        f.write(f'<file alias="icons/{file}">icons/{file}</file>\n')
    f.write('</qresource>\n</RCC>')
subprocess.run(['pyrcc5', '-o', 'resources.py', 'resources.qrc'])