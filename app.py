
from flask import Flask, render_template, request, send_file
from zipfile import ZipFile
from docx_Play import *
from io import BytesIO

import subprocess
import zipfile
import sys
import os

subprocess.check_call([sys.executable, "-m", "pip", "install", 'openpyxl'])

app = Flask(__name__)

if not os.path.isfile('lead_Pit/LRA/finished_Docs'):
    os.makedirs('lead_Pit/LRA/finished_Docs', exist_ok=True)
if not os.path.isfile('uploads'):
    os.makedirs('uploads', exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download', methods=['GET'])
def download():
    if request.method == 'GET':
        zip_buffer = BytesIO()
        zipfolder = ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_STORED)
        app_name = 'sample_Folder/' + str(os.listdir('sample_Folder')[0])
        for (root, dirs, files) in os.walk(app_name, topdown=True):
            if files:
                for file in files:
                    zipfolder.write(root + '/' + file)
        zipfolder.close()
        zip_buffer.seek(0)

        return send_file(zip_buffer, mimetype='zip', attachment_filename='sample_Folder.zip', as_attachment=True)


@app.route('/', methods=['POST'])
def upload_file():
    if not os.path.isfile('lead_Pit/LRA/finished_Docs'):
        os.makedirs('lead_Pit/LRA/finished_Docs', exist_ok=True)
    if not os.path.isfile('uploads'):
        os.makedirs('uploads', exist_ok=True)

    for file in request.files.getlist('file'):  # temporarily save import folder in uploads folder
        path_conc = ''
        for x in range(1, len(str(file.filename).split('/'))-1):
            if not os.path.exists('uploads' + path_conc + '/' + str(file.filename).split('/')[x]):
                os.makedirs('uploads' + path_conc + '/' + str(file.filename).split('/')[x])
            path_conc += '/' + str(file.filename).split('/')[x]
        file.save('uploads' + path_conc + '/' + str(file.filename).split('/')[x+1])

    make_it()  # create and temporarily save all report documents

    # save all report documents on ephemeral zipfile
    zip_buffer = BytesIO()
    zipfolder = ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_STORED)
    for (root, dirs, files) in os.walk('lead_Pit/LRA/finished_Docs', topdown=True):
        if files:
            print('root', root)
            for file in files:
                zipfolder.write(root + '/' + file, 'output/' + str(root).split('\\')[-1] + '/' + file)
    zipfolder.close()
    zip_buffer.seek(0)

    # after zipfile is written, clear all temporary data from server
    for (root, dirs, files) in os.walk('uploads', topdown=True):
        if files:
            for file in files:
                os.remove(root + '/' + file)
    for (root, dirs, files) in os.walk('uploads', topdown=False):
        os.rmdir(root)

    for (root, dirs, files) in os.walk('lead_Pit/LRA/finished_Docs', topdown=True):
        if files:
            for file in files:
                os.remove(root + '/' + file)
    for (root, dirs, files) in os.walk('lead_Pit/LRA/finished_Docs', topdown=False):
        os.rmdir(root)

    # send zipfile to user
    return send_file(zip_buffer, mimetype='zip', attachment_filename='test.zip', as_attachment=True)


if __name__ == '__main__':
    app.run(debug=False, threaded=True)
