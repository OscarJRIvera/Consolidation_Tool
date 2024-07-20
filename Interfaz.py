#Libraries
from customtkinter import *
from tkinter import Label, PhotoImage, filedialog
import customtkinter
import Watcher as w
import shutil as sh
from PIL import Image,ImageTk

customtkinter.set_appearance_mode("dark")
app = CTk()

#Watcher class object used to monitor folder
Excelwatcher= w.Watcher() 


#Method for user to choose a folder to monitor
def Open_Folder():
    Folder_path = filedialog.askdirectory()
    if Folder_path:
        folder_name = os.path.basename(Folder_path)
        file_label.configure(text=folder_name)
        Excelwatcher.Change_path(Folder_path)

        #Make sure the watcher stops running if it is
        Excelwatcher.Stop()
        #Starts watcher
        Excelwatcher.Start()

        #Uploading file is now possible
        Button2.configure(state="normal")
        
        log_message(f"Folder selected: {folder_name}")


#Method for user to upload a file
def Upload_File():
    file_path = filedialog.askopenfilename()
    if file_path:
        sh.move(file_path, Excelwatcher.path)
        #Moves file into the folder that is selected
        file_name = os.path.basename(file_path)
        log_message(f"File uploaded: {file_name}")

def log_message(message):
    log_textbox.insert(END, message + "\n")
    log_textbox.yview(END) 




#--------------UI-------------------


#Background images used for the UI
image = Image.open("Images/background.jfif")
background_image = customtkinter.CTkImage(image, size=(600, 400))
image2 = Image.open("Images/Genpact.png")
background_image2 = customtkinter.CTkImage(image2, size=(80, 60))

#app size
app.geometry("600x400")

app.title("Folder Watcher Application")
app.iconbitmap('images/Genpact.ico')



img1 = Image.open('Images/upload.png')
img2 = Image.open('Images/folder.png')


#Widgets used 

bg_lbl = customtkinter.CTkLabel(app, text="", image=background_image)
bg_lbl.place(x=0, y=0)

bg_lbl2 = customtkinter.CTkLabel(app, text="", image=background_image2)
bg_lbl2.place(x=0.0, y=0.0)


Button1= CTkButton(master=app,text="Choose a folder to watch", command=Open_Folder,text_color="white",corner_radius=10,image=CTkImage(img2))
Button1.place(relx=0.80, rely=0.25, anchor="center")

Button2= CTkButton(master=app,text="Upload file", command=Upload_File,state="disabled",text_color="black",corner_radius=10,image=CTkImage(img1))
Button2.place(relx=0.80, rely=0.1, anchor="center")

file_label = CTkLabel(master=app, text="No folder selected",bg_color="#333333")
file_label.place(relx=0.8, rely=0.35, anchor="center")


log_textbox = CTkTextbox(master=app, width=200, height=300,fg_color="#333333",bg_color="#333333")
log_textbox.place(relx=0.40, rely=0.45, anchor="center")



app.resizable(False, False)




app.mainloop()