from tkinter import *
import tkinter
from tkinter import ttk
import openpyxl
from tkinter import filedialog
from PIL import Image, ImageTk


def open_halls():
    if not window_open_rooms:  # Check if rooms window is not already open
        open_rooms()

def open_rooms():
    global window_open_rooms
    window_open_rooms = True  # Set flag to indicate rooms window is open
    exec(open('C:/Users/DELL/Documents/intellipaat/Projects/mini project/mini project/halls1.py').read(), globals())

def open_summary_hall_sheet():
    exec(open('C:/Users/DELL/Documents/intellipaat/Projects/mini project/mini project/gui_summary_and_hall.py').read(), globals())

def exit_application():
    if root:
        root.destroy()

root = Tk()
root.title("SEATING ARRANGEMENT FOR END SEMESTER EXAMINATIONS")
screen_width = int(root.winfo_screenwidth())
screen_height = int(root.winfo_screenheight())
root.geometry(f"{screen_width}x{screen_height}")
#root.configure(background='#E6E6FA', bd=15, highlightthickness=10, highlightbackground="#800080",
#               highlightcolor="#E6E6FA", relief="groove")
bg_image = Image.open("C:/Users/DELL/Documents/intellipaat/Projects/mini project/mini project/5.jpg")
bg_image = bg_image.resize((screen_width, screen_height), Image.ADAPTIVE)
bg_image = ImageTk.PhotoImage(bg_image)
bg_label = Label(root, image=bg_image)#87CEEB#FFFFFF
bg_label.place(relwidth=1, relheight=1)

label = Label(root, text="Mahaveer Institute of Science and Technology,Hyderabad",fg='#87CEEB',font=("Comic Sans MS", 30, 'bold', 'italic', "underline"),underline=0)
label.place(x=130, y=5)

label = Label(root, text="Seating Arrangement For External Examinations",fg='#87CEEB',
                  font=("Comic Sans MS", 25, 'bold', 'italic', "underline"),underline=0)
label.place(x=300,y=80)
window_open_rooms = False

btn1 = Button(root,text='Halls',bg='#87CEEB',fg='#FFFFFF',font=("Comic Sans MS", 30, 'bold', 'italic')
              ,command=open_halls)
btn1.place(x=350,y=280)

btn2 = Button(root,text='Hall Plan &\nSummary Sheet',bg='#87CEEB',fg='#FFFFFF',
              font=("Comic Sans MS", 20, 'bold', 'italic'),command=open_summary_hall_sheet)
btn2.place(x=650,y=280)

btn3 = Button(root,text='Exit',bg='#87CEEB',fg='#FFFFFF',
              font=("Comic Sans MS", 30, 'bold', 'italic'),command=exit_application)
btn3.place(x=500,y=480)
root.protocol("WM_DELETE_WINDOW", exit_application)
root.mainloop()
