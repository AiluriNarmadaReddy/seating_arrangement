import tkinter as tk
from tkinter import filedialog
import subprocess
from PIL import Image, ImageTk
def browse_file():
    file_path = filedialog.askopenfilename(parent=frame1)  # Set parent to root window
    if file_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, file_path)

def get_output():
    input_path = entry_path.get()
    output = subprocess.check_output(['python', 'C:/Users/DELL/Documents/intellipaat/Projects/mini project/mini project/all_roll.py',input_path])
    output_label.configure(text=output.decode('utf-8'))

def exit_application():
    frame1.destroy()
    
    
def run_console_app():
    input_path = entry_path.get()
    subprocess.run(["python",'C:/Users/DELL/Documents/intellipaat/Projects/mini project/mini project/out_sheet.py',input_path])
    
 
frame1 = tk.Tk()
frame1.geometry("900x600+10+10")
frame1.protocol("WM_DELETE_WINDOW", exit_application)
frame1.configure(background='#E6E6FA', bd=15, highlightthickness=10, highlightbackground="#800080",
               highlightcolor="#E6E6FA", relief="groove")


frame1.title("Hall Sheet & Summary Sheet")

label = tk.Label(frame1, text="Browse the path for Excel Sheet", fg='#87CEEB', bg='white',
                  font=("Comic Sans MS", 15, 'bold', 'italic', "underline"), underline=0)
label.pack(pady=10)

frame_browse = tk.Frame(frame1, bg='#E6E6FA')
frame_browse.pack(pady=50)

entry_path = tk.Entry(frame_browse, width=40, font=('Arial', 16))
entry_path.grid(row=0, column=0, padx=5, pady=5)

button_browse = tk.Button(frame_browse, text="Browse", command=browse_file, bg='white', fg='#87CEEB',
                          font=("Comic Sans MS", 16, 'bold', 'italic'))
button_browse.grid(row=0, column=1, padx=5, pady=5)
run_button = tk.Button(frame_browse, text='Check if Rooms\n are Sufficient', command=get_output,bg='white',fg='#87CEEB',font=("Comic Sans MS", 16, 'bold', 'italic'))
run_button.grid(row=1, column=0, padx=5, pady=5)

button_run = tk.Button(frame_browse, text="Halls And \nSummary Sheet", command=run_console_app,bg='white',fg='#87CEEB',font=("Comic Sans MS", 16, 'bold', 'italic'))
button_run.grid(row=1, column=1, padx=5, pady=5)

output_label = tk.Label(frame_browse, text='', font=("Comic Sans MS", 16, 'bold', 'italic'),bg='#E6E6FA')
output_label.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

button_run = tk.Button(frame_browse, text="Exit", command=exit_application,bg='white',fg='#87CEEB',font=("Comic Sans MS", 16, 'bold', 'italic'))
button_run.grid(row=3, column=0, padx=5, pady=5)

frame1.mainloop()

