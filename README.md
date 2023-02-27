# pptx_to_pdf
Very basic app for converting *.pptx files to *.pdf files

### For Ubuntu
- *_u.py is ubuntu version.

### For Windows
- "*_w.py" is Windows version.
- Using the 'subprocess' library, pptx files are converted to PDF through the command-line interface of PowerPoint. The process.communicate() function is used to ensure that PowerPoint runs in the background during the conversion process.

#### Note: 
The stdout and stderr arguments of the subprocess. Popen() function can be used to capture output and error messages. However, they are not used in this example. Additionally, the folder name where the files are located is assumed to be "willConvert" and the folder where PowerPoint is installed is assumed to be "C:\Program Files\Microsoft Office\root\Office16". Adjust these values to match your own system configuration.
