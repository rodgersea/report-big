
from flask import Flask, render_template, request, send_file
from zipfile import ZipFile
from docx_Play import *
from io import BytesIO

import subprocess
import zipfile
import sys
import os

app = Flask(__name__)

if not os.path.isfile(os.path.join(app.root_path, 'finished_Docs')):  # create directory "finished_Docs" in root
    os.makedirs(os.path.join(app.root_path, 'finished_Docs'), exist_ok=True)
if not os.path.isfile(os.path.join(app.root_path, 'uploads')):  # create directory "uploads" in root
    os.makedirs(os.path.join(app.root_path, 'uploads'), exist_ok=True)


@app.route('/')  # on load page, display index.html located in the "templates" directory in root
def index():
    return render_template('index.html')


# user requests download of sample folder, serves samples folder as zipfile
@app.route('/download', methods=['GET'])
def download():
    if request.method == 'GET':
        zip_buffer = BytesIO()  # we have no reason to store zip file changes to the server, so we modify zip archives in memory with bitesIO()
        zipfolder = ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_STORED)  # Zipfile(file, mode, compression)
        app_name = os.join.path(app.root_path, 'sample_Folder', str(os.listdir('sample_Folder')[0]))  # the sample folder is a permanent feature on the server and is served on buttom click as a zipfile
        for (root, dirs, files) in os.walk(app_name, topdown=True):  # step function to construct the zipfile
            if files:
                for file in files:
                    zipfolder.write(os.path.join(root, file))
        zipfolder.close()
        zip_buffer.seek(0)  # set reference point at the beginning of the file

        return send_file(zip_buffer, mimetype='zip', attachment_filename='sample_Folder.zip', as_attachment=True)  # return zipfile of sample folder to user


# user submits data set to report writer
# app returns zipfile with lead based paint report as pdf
@app.route('/', methods=['POST'])
def upload_file():
    dispp(request.files.getlist('file')[0])
    app_str = str(request.files.getlist('file')[0]).split('/')[1].split(' - ')[0]
    # temporarily save import folder in uploads folder
    for file in request.files.getlist('file'):  # add every folder and file
        path_conc = os.path.join(app.root_path, 'uploads')  # path_concatenate, compounding path for every file's path
        for x in range(1, len(str(file.filename).split('/'))-1):  # find length of path for file from root
            file_hold = os.path.join(path_conc, str(file.filename).split('/')[x])  # create path for current directory creation
            if not os.path.exists(file_hold):  # checks for duplicate folders and files
                os.makedirs(file_hold)
            path_conc = file_hold  # effectively concatenates last folder/file in loop to beginning of next loop run
        file.save(os.path.join(path_conc, str(file.filename).split('/')[x+1]))  # save folder/file

    make_it()  # run func_repo; create all report documents on disk/server, return documents to user, then erase documents from disk/server

    # save all report documents on ephemeral zipfile
    zip_buffer = BytesIO()  # previous mention
    zipfolder = ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_STORED)  # previous mention
    for (root, dirs, files) in os.walk(os.path.join(app.root_path, 'finished_Docs'), topdown=True):  # creates the zipfile the same way we served the sample folder
        if files:
            print('root', root)
            for file in files:
                zipfolder.write(os.path.join(root, file), os.path.join('output', str(root).split('\\')[-1], file))
    zipfolder.close()
    zip_buffer.seek(0)

    # after zipfile is written, clear all temporary data from server, from the uploads and the finished_Docs directories
    for (root, dirs, files) in os.walk(os.path.join(app.root_path, 'uploads'), topdown=True):  # same method of looping through directory for zipfile, but with os.remove
        if files:
            for file in files:
                os.remove(os.path.join(root, file))
    for (root, dirs, files) in os.walk(os.path.join(app.root_path, 'uploads'), topdown=False):
        os.rmdir(root)
    for (root, dirs, files) in os.walk(os.path.join(app.root_path, 'finished_Docs'), topdown=True):
        if files:
            for file in files:
                os.remove(os.path.join(root, file))
    for (root, dirs, files) in os.walk(os.path.join(app.root_path, 'finished_Docs'), topdown=False):
        os.rmdir(root)

    return send_file(zip_buffer, mimetype='zip', attachment_filename=app_str + '_LBP_Report.zip', as_attachment=True)  # send output zipfile to user


if __name__ == '__main__':  # not sure what this does, let's delete it at the very end
    app.run(debug=False, threaded=True)
