# -*- coding: utf-8 -*-
"""
Created on Mon Apr 17 18:08:44 2023

@author: Strasshofer
"""

import tkinter as tk
import openpyxl
import getpass
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import Calendar
from datetime import date
from docxtpl import DocxTemplate
import os
import shutil



# creating a window
root = tk.Tk()
#root.iconphoto(False, tk.PhotoImage(file='TU.png'))
user = getpass.getuser()

#------------------------------------------------------------------------------
#root.geometry ("800x500")  #Size of window

root.title("Quality Awareness")# Naming of program
root.wm_attributes("-topmost", 1) # Make stay on top
root.eval('tk::PlaceWindow . center') # Center when opened


style = ttk.Style(root) # Call Theme
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")


# define options in comboboxs
#------------------------------------------------------------------------------
dept_list = ["Shipping","Receiving","Sortation","ORC", "OBE/PPS", "Triage","Service","Quality Assurance"]
reason_list = {
    'Quality Assurance' : ['Safety','SOP Deviation','Transfer Audits','Incorrect IA','Over-Receive D/R'],
    'Shipping' : ['Safety', 'SOP Deviation','QA Found','Mis-Placed Pallet'],
    'Receiving' : ['Safety', 'SOP Deviation', 'Accuracy','Over Received','Miss Labeled'],
    'Sortation' : ['Safety', 'SOP Deviation','QA Found'],
    'Returns' : ['Safety', 'SOP Deviation','Transfer Creation','Incorrect IA','Incorrect IA request','Incorrect SKU '],
    'OBE/PPS' : ['Safety', 'SOP Deviation','Transfer Creation'],
    'Testing' : ['Safety', 'SOP Deviation','Boxing'],
    'Service' : ['Safety', 'SOP Deviation',],
    'External' : ['Safety', 'SOP Deviation',]
}

#Create and plce file in your documents folder
#------------------------------------------------------------------------------


qaf_master = "C:\\Users\\Strasshofer\\Documents\\Programming\\Python Scripts\\Incident Tracker\\QAF.xlsx"
qaf_doc = 'QAF.docx'
store = 'store.txt'
file_path = os.path.join(os.path.expanduser('~'), 'Documents/QAF', qaf_doc)
store_path = os.path.join(os.path.expanduser('~'), 'Documents/QAF', store)

if os.path.exists(file_path):
    pass
else:
    # Copy the file from the project folder to the Documents folder
    project_file_path = qaf_doc
    shutil.copyfile(project_file_path, file_path)
    messagebox.showinfo("New File", "A new Word Doc file named QAF.docx has been added to your Documents folder.")
    
if os.path.exists(store_path):
    pass
else:
    # Copy the file from the project folder to the Documents folder
    store_file_path = store
    shutil.copyfile(store_file_path, store_path)
    messagebox.showinfo("New File", "A new file named store.txt has been added to your Documents folder. /n This will store your location number for tracking.")

def open_store_file():
    os.startfile(store_path)
    messagebox.showinfo("Friendly Reminder","Dont forget to save! \nDo not move the file.")

def open_qaf_file():
    os.startfile(file_path)
    messagebox.showinfo("Friendly Reminder","Dont forget to save! \nDo not move the file.")
    


# Update the options of the reason_combobox based on the selected department
def update_reason_options(event):
    selected_dept = dept_combobox.get()
    if selected_dept in reason_list:
        reasons = reason_list[selected_dept]
        reason_combobox.config(values=reasons)
        reason_combobox.current(0)
    else:
        reason_combobox.config(values=[])

def load_data():
    path = qaf_master
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    print(list_values)

    for col_name in list_values[0]:         # display column names in Treeview
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:     # populate treeview with data
        treeview.insert('', tk.END, values=value_tuple)

def filter_by_user(user):
    for row in treeview.get_children():
        if treeview.item(row)["values"][1] != user:
            treeview.delete(row)


#------------------------------------------------------------------------------
def update_treeview():
    # Clear treeview
    treeview.delete(*treeview.get_children())
    df = pd.read_excel("QAF.xlsx") # Read data from Excel file
    df = df.loc[df["user"] == user]  # Filter data based on current user
    for index, row in df.iterrows():  # Insert data into treeview
        treeview.insert("", "end", values=row.to_list())
#------------------------------------------------------------------------------

   

def clear_inputs():
    associate_entry.delete(0, "end")
    associate_entry.insert(0, "Associate")
    dept_combobox.set(dept_list[0])
    reason_combobox.update()
    issue_date.place_forget()
    cost_entry.delete(0,"end")
    cost_entry.insert(0, "Cost")
    research_entry.delete(0, "end")
    research_entry.insert(0, "Research")
    correction_entry.delete(0, "end")
    correction_entry.insert(0, "Coaching/Suggestions")
    shrinkcheck.state(["!selected"])


def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")
        
def show_selected_date():
    issue_date.config(text="Selected date: " + cal.get_date())
# Define a function to format dates

def get_date():
    selected_date = cal.get_date().strftime("%d-%m-%y")
    return selected_date


def delete_row():
    selection = treeview.selection()
    for item in selection:
        treeview.delete(item)


        
def render_inputs():
    issue_date = cal.get_date()
    formatted_date = issue_date.strftime("%d-%m-%y")
    doc_filename = "QAF.docx"
    doc_path = os.path.join(os.path.expanduser("~"), "Documents", "QAF", doc_filename)
    save_path = os.path.join(os.path.expanduser("~"), "Documents", "QAF", 'output.docx')
    doc = DocxTemplate(doc_path)
    submit_date = date.today().strftime("%d-%m-%y")
    date_of_issue = formatted_date
    associate = associate_entry.get()   
    department = dept_combobox.get()
    reason = reason_combobox.get()
    research = research_entry.get()
    correction = correction_entry.get()
    
    context = {
    'submit_date': submit_date,
    'date_of_issue': date_of_issue,
    'associate': associate,
    'department': department,
    'reason': reason,
    'research': research,
    'correction': correction
    }
    
    doc.render(context)
    doc.save(save_path)
    os.startfile(save_path)

def insert_row():
    issue_date = cal.get_date()
    formatted_date = issue_date.strftime("%d-%m-%y")
    
    with open(store_path, 'r') as f:
        location = f.read()
    
    user = getpass.getuser()
    submit_date = date.today().strftime("%d-%m-%y")
    date_of_issue = formatted_date
    associate = associate_entry.get()   
    department = dept_combobox.get()
    reason = reason_combobox.get()
    shrink_taken = "yes" if a.get() else "no"
    cost = cost_entry.get()
    research = research_entry.get()
    correction = correction_entry.get()
    
    print(location, user, submit_date, date_of_issue, associate, department, reason, shrink_taken, cost, research, correction)

    # Insert row into Excel sheet
    path = qaf_master
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [location, user, submit_date, date_of_issue, associate, department, reason, shrink_taken, cost, research, correction]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)



#Starting the Gui
#------------------------------------------------------------------------------
frame = ttk.Frame(root)
frame.pack()


#------------------------------------------------------------------------------
menu_bar = tk.Menu(root)

# create a file menu
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="New")
file_menu.add_command(label="Open")
file_menu.add_command(label="Save")
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

# create a settings menu
settings_menu = tk.Menu(menu_bar, tearoff=0)
settings_menu.add_command(label="Update Location", command=open_store_file)
settings_menu.add_command(label="Update QAF File", command=open_qaf_file)

# add the menus to the menu bar
menu_bar.add_cascade(label="File", menu=file_menu)
menu_bar.add_cascade(label="Update Files", menu=settings_menu)

# set the menu bar for the root window
root.config(menu=menu_bar)




#Inputs 
#------------------------------------------------------------------------------

# outlines left first column
widgets_frame = ttk.LabelFrame(frame, text = "Insert Row")
widgets_frame.grid(row = 0, column = 0, padx=20, pady=10)

associate_entry = ttk.Entry(widgets_frame, font = 'dokchampa') # Creates Text box, selects font
associate_entry.insert(0, "Associate")  # Places Text in textbox 
associate_entry.bind("<FocusIn>", lambda e: associate_entry.selection_range(0, tk.END)) #highlights textbox contents when clicked in 
#associate_entry.bind("<FocusIn>", lambda e: associate_entry.delete('0', 'end')) # deletes contents 
associate_entry.grid(row=1, column = 0, padx=5, pady=5, sticky="ew") # displays textbox 

dept_combobox = ttk.Combobox(widgets_frame, values=dept_list, font = 'dokchampa')
dept_combobox.current(0)
dept_combobox.grid(row=2, column=0, padx=5, pady=5,  sticky="ew")
dept_combobox.bind("<<ComboboxSelected>>", update_reason_options)

reason_combobox = ttk.Combobox(widgets_frame, values="Select_Dept.", font = 'dokchampa')
reason_combobox.current(0)
reason_combobox.grid(row=3, column=0, padx=5, pady=5,  sticky="ew")

#Create a button to open the calendar
select_date_button = tk.Button(widgets_frame, text="Date of Issue", 
                            command=lambda: cal.place(relx=0.5, 
                            rely=0.5, anchor=tk.CENTER), font = 'dokchampa')    # Create a label to display the selected date
issue_date = tk.Label(widgets_frame, text="Selected date: None")
cal = Calendar(widgets_frame, selectmode="day", date_pattern="mm-dd-yy")# Create the calendar widget
cal.bind("<<CalendarSelected>>", lambda event: [cal.place_forget(), show_selected_date()])# Attach the calendar to the button and hide it initially
tk.Widget.lift(cal)
select_date_button.grid(row=4, column=0, pady=10)# Add the widgets to the window
issue_date.grid(row=5, column=0)

a = tk.BooleanVar()
shrinkcheck = ttk.Checkbutton(widgets_frame, text="Shrink taken?", variable=a)
shrinkcheck.grid(row=6, column=0, padx=5, pady=5, sticky="nsew")
shrinkcheck.lower()

cost_entry = ttk.Entry(widgets_frame, font = 'dokchampa')
cost_entry.insert(0, "Unit(s) Cost")
cost_entry.bind("<FocusIn>", lambda e: cost_entry.selection_range(0, tk.END))
cost_entry.grid(row=7, column = 0, padx=5, pady=5, sticky="ew")
cost_entry.lower()

research_entry = ttk.Entry(widgets_frame, font = 'dokchampa')
research_entry.insert(0, "Research")
research_entry.bind("<FocusIn>", lambda e: research_entry.selection_range(0, tk.END) )
research_entry.grid(row=8, column = 0, padx=5, pady=5, sticky="ew")
research_entry.lower()


correction_entry = ttk.Entry(widgets_frame, font = 'dokchampa')
correction_entry.insert(0, "Coaching/Suggestions")
correction_entry.bind("<FocusIn>", lambda e: correction_entry.selection_range(0, tk.END))
correction_entry.grid(row=11, column = 0, padx=5, pady=5, sticky="ew")
correction_entry.lower()


#Buttons
#------------------------------------------------------------------------------

sub_button = ttk.Button(widgets_frame, text = "Submit", command = insert_row, width = 11)
sub_button.grid(row=12, column=0, padx=5, pady=5, sticky="w")
clear_button = ttk.Button(widgets_frame, text = "Clear Inputs", command = clear_inputs, width = 11)
clear_button.grid(row=12, column=0, padx=5, pady=5, sticky="e")

del_button = ttk.Button(widgets_frame, text = "Delete", command = delete_row, width = 11)
del_button.grid(row=13, column=0, padx=5, pady=5, sticky="e")


print_button = ttk.Button(widgets_frame, text = "Render", command = render_inputs, width = 11)
print_button.grid(row=13, column=0, padx=5, pady=5, sticky="w")

mode_switch = ttk.Checkbutton(
    widgets_frame, text = "Theme", style = "Switch", command = toggle_mode)
mode_switch.grid(row=14, column=0, padx=5, pady=10, sticky="nsew")


#Excel Viewer
#------------------------------------------------------------------------------
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("location", "user", "submit_date", "date_of_issue","associate","department","reason","shrink_taken","cost","research","correction")

treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)

treeview.column("location", minwidth=0, width=0)
treeview.column("user", minwidth=0, width=0)
treeview.column("submit_date", minwidth=0, width=0)
treeview.column("date_of_issue", width=80)
treeview.column("associate", width=90)
treeview.column("department", width=80)
treeview.column("reason", width=80)
treeview.column("shrink_taken", minwidth=0, width=0)
treeview.column("cost", minwidth=0, width=0)
treeview.column("research", width=150)
treeview.column("correction", width=150)

treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()







root.mainloop()
