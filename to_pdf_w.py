import os
import subprocess

input_folder = os.path.join(os.path.expanduser('~'), 'Desktop', 'willConvert')
output_folder = os.path.join(os.path.expanduser('~'), 'Desktop', 'Converted')

for root, dirs, files in os.walk(input_folder):
    for file in files:
        if file.endswith('.pptx'):
            input_path = os.path.join(root, file)
            output_path = os.path.join(output_folder, os.path.splitext(file)[0] + '.pdf')
            process = subprocess.Popen(['C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE', '/pt', input_path, output_path])
            stdout, stderr = process.communicate()
