import tkinter as tk
from tkinter import filedialog
import win32com.client
import os
import tempfile
import sys
import time
import shutil

input_file = ''
ext_code = 18 #PNG
savePath = ''
delete = False # to delete temp file

mode = sys.argv[1]
if mode == 'drag':
    input_file = sys.argv[2]
elif mode == 'save':
    savePath = sys.argv[2]
    selected_ext = sys.argv[4]
    input_file = sys.argv[3]

    if selected_ext == 'pdf':
        ext_code = 32
    elif selected_ext == 'jpg':
        ext_code = 17
else:
    root = tk.Tk()
    root.withdraw()
    try:
        input_file = filedialog.askopenfilename(title = "Select file", filetypes = (("PPT files","*.ppt *.pptx"),("all files","*.*")))
    except Exception as error:
        e = error
        print('error')

if input_file != '' and input_file != None:


    if mode == 'drag' or mode == 'open':
        temp_directory = "pptviewer"
        parent_directory = tempfile.gettempdir()
        path = os.path.join(parent_directory, temp_directory)
        try:
            os.mkdir(path)
        except OSError as error:  
            e = error

        modified_time = time.ctime(os.path.getmtime(input_file)).replace(":", "").replace(" ", "")
        created_time = time.ctime(os.path.getctime(input_file)).replace(":", "").replace(" ", "")

        filename = os.path.basename(input_file).replace(".", "") + modified_time + created_time
        
        output_file = path + '\\' + filename

        stored_files = os.listdir(path)
        baseName = os.path.basename(input_file).split(".")[0]
        if stored_files.__contains__(filename):
            fileCount = len(os.listdir(output_file))
            print(output_file + '==**==' + str(fileCount) + '==**==' + baseName + '==**==' + input_file)
            sys.stdout.flush()
        else:
            originalInputFile = input_file
            
            desktopPath = os.path.join(os.environ["HOMEPATH"], "Desktop")
            try:
                shutil.copy(input_file, desktopPath)
                input_file = desktopPath + '\\' + os.path.basename(input_file)
                delete = True
            except shutil.SameFileError:
                pass
            try:
                powerpoint = win32com.client.Dispatch('Powerpoint.Application')
                pdf = powerpoint.Presentations.Open(input_file, WithWindow=False)

                pdf.SaveAs(output_file, ext_code)

                fileCount = len(os.listdir(output_file))

            except Exception as error:
                print('error' + '==**==' + str(error))
                sys.stdout.flush()
            else:
                print(output_file + '==**==' + str(fileCount) + '==**==' + baseName + '==**==' + originalInputFile)
                sys.stdout.flush()
                pdf.Close()
                powerpoint.Quit()
                if delete:
                    os.remove(input_file)
    else:
        try:
            originalInputFile = input_file
            desktopPath = os.path.join(os.environ["HOMEPATH"], "Desktop")
            try:
                shutil.copy(input_file, desktopPath)
                input_file = desktopPath + '\\' + os.path.basename(input_file)
                delete = True
            except shutil.SameFileError:
                pass
            baseName = os.path.basename(input_file).split(".")[0]
            output_file = savePath + '\\' + baseName
            powerpoint = win32com.client.Dispatch('Powerpoint.Application')
            pdf = powerpoint.Presentations.Open(input_file, WithWindow=False)
            pdf.SaveAs(output_file, ext_code)
        except Exception as error:
            print('error' + '==**==' + str(error))
            sys.stdout.flush()
        else:
            print('success')
            sys.stdout.flush()
            pdf.Close()
            powerpoint.Quit()
            if delete:
                os.remove(input_file)

    




