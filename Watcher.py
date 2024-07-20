from tkinter import filedialog, messagebox
import shutil as sh
from watchdog.observers import Observer
from watchdog.events import FileSystemEvent, FileSystemEventHandler
import pandas as pa
import threading
import os
import shutil
import openpyxl

PathGlobal= None

#Any events that occur with the files are done in this class
#class has parent FileSystemEventHandler , responsible for managing events
class FileManagment(FileSystemEventHandler):

    #When file is created/inserted inside the folder
    def on_created(self, event):
        if not event.is_directory:
            self.process_file(event.src_path)

    #Method for verifying file extension and merging with master file
    def process_file(self, file_path):
        global PathGlobal
        #Folders paths
        path_Proc = PathGlobal + "/Processed"
        path_NotA= PathGlobal +"/NotApplicable"

        #Check extension
        if file_path.endswith(('.xls', '.xlsx')):
            if not file_path.endswith(('MasterExcel.xls', 'MasterExcel.xlsx')):
                self.Check_folders()
                #consolidate new file with master
                self.Consolidate_files(file_path)
                #Creates new name just in case there is a file that already exist with said name
                newname=self.Generate_Name(os.path.basename(file_path),path_Proc)
                #Moves file into processed folder
                sh.move(file_path, os.path.join(path_Proc, newname))
        else:
            #Process for wrong extension files
            self.Check_folders()
            newname=self.Generate_Name(os.path.basename(file_path),path_NotA)
            sh.move(file_path, os.path.join(path_NotA, newname))

    #Simple method that makes sure both folders are created
    def Check_folders(self):
        global PathGlobal
        path_NotA= PathGlobal +"/NotApplicable"
        path_Proc = PathGlobal + "/Processed"
        os.makedirs(path_NotA, exist_ok=True)
        os.makedirs(path_Proc, exist_ok=True)
    #Method to make unique name
    def Generate_Name(self,File_name,Folder):
        base, extension = os.path.splitext(File_name)
        counter = 1
        while os.path.exists(os.path.join(Folder, File_name)):
            File_name = f"{base}_{counter}{extension}"
            counter += 1
        return File_name


    def Consolidate_files(self,EFile_path):
        global PathGlobal
        #attemps before the consolidation stops
        attempts=3
        for attempt in range(attempts):
            #Try for writing erros
            try:
                #loads file
                with pa.ExcelFile(EFile_path) as EFile:
                    Path_main= PathGlobal+ "/MasterExcel.xlsx"

                    #Make sures master file exists
                    if not os.path.exists(Path_main):
                        with pa.ExcelWriter(Path_main, engine='openpyxl') as writer:
                            pa.DataFrame().to_excel(writer, sheet_name='Sheet',index=False)
                            
                    #Starts writing on masterfile    
                    with pa.ExcelWriter(Path_main, engine='openpyxl',mode = 'a',if_sheet_exists='replace') as writer:
                        for name in EFile.sheet_names:
                            Sheet= pa.read_excel(EFile_path,sheet_name=name,header=None)
                            Sheet.to_excel(writer,sheet_name=name,index=False,header=False)
                break
            except PermissionError:
                messagebox.showerror("Error", f"Failed to write. Please close the Excel file")
        else:
            pass
    
        

    




#Has all the methods to monitor folder

class Watcher:
    #constructor, atributes
    def __init__(self):
        self.path= None
        self.observer =None

    #Method that start thread to monitor folder
    def Start(self):
        event_handler = FileManagment()
        self.observer  = Observer()
        self.observer .schedule(event_handler, self.path, recursive=False)

        threadObserver =threading.Thread(target=self.observer.start)
        threadObserver.daemon = True
        threadObserver.start()

    #Method that stops observer 
    def Stop(self):
        if self.observer is not None:
            self.observer.stop()
            #stop just turns on a flag but needs to be called also, that is why join is necessary 
            self.observer.join()

    #Simple method so that interface.py can change folder path
    def Change_path(self,path):
        global PathGlobal
        self.path=  path
        PathGlobal= path

    