from tkinter import Tk, filedialog, StringVar
from tkinter.ttk import Frame, Label, Button, Style
import main
import time


# Initialize Tk root window
root = Tk()
root.title('Work Order Generator')
root.geometry('850x800')

# Variables to store the file paths
master_sheet_path = StringVar()
dpp_sheet_path = StringVar()
isMasterSelected = False
isDppSelected = False
master_sheet_var = StringVar()
dpp_sheet_var = StringVar()
user_message = StringVar()

# Upload master excel sheet
def UploadMaster(event=None):
    global master_sheet_path, isMasterSelected 
    filename = filedialog.askopenfilename()
    print('Selected:', filename)
    #master_sheet_path = filename
    master_sheet_path.set(filename)
    isMasterSelected = True
    master_sheet_var.set(filename) 
    print(f'Master sheet status: {isMasterSelected}')

# Upload dpp sheet
def UploadDPP(event=None):
    global dpp_sheet_path, isDppSelected
    filename = filedialog.askopenfilename()
    print('Selected:', filename)
    #dpp_sheet_path = filename
    dpp_sheet_path.set(filename)
    isDppSelected = True
    dpp_sheet_var.set(filename) 
    print(f'DPP sheet status: {isDppSelected}')

# Executes the main script
# Formats, extracts, merges and saves new work orders
def ExecuteScript(event=None):
    global master_sheet_path, isMasterSelected
    global dpp_sheet_path, isDppSelected
    if isMasterSelected and isDppSelected:
        user_message.set('Executing script...')
        master_filepath = master_sheet_path.get()
        dpp_filepath = dpp_sheet_path.get()
        main.execute_script(master_filepath, dpp_filepath)
        user_message.set('Script execution complete. Please check the master sheet.')
        master_sheet_path.set('')
        master_sheet_var.set('')
        dpp_sheet_path.set('')
        dpp_sheet_var.set('')
        isMasterSelected = False
        isDppSelected = False
    else:
        user_message.set('Please select both files before submitting.')
        print('Please select both files before submitting')

# Setup style for the application
style = Style()
style.theme_use('default')

# Configure Button styles
style.configure('TButton', font=('Helvetica', 12), padding=(10, 5))
style.configure('Accent.TButton', font=('Helvetica', 12, 'bold'), padding=(10, 5))

# Create Labels
title = Label(root, text='Work Order Generator', font=('Arial', 24))
description = Label(root, text="Extract and add new works order to your Excel sheet", font=('Arial', 12))

# Create Labels and Buttons for file selection
file_frame = Frame(root)

master_sheet_label = Label(file_frame, text='Master Excel Sheet', font=('Arial', 10))
select_file_btn_1 = Button(file_frame, text="Select File", command=UploadMaster, width=20)
master_sheet_display = Label(file_frame, textvariable=master_sheet_var, font=('Arial', 10))

dpp_sheet_label = Label(file_frame, text='Exported DPP Sheet', font=('Arial', 10))
select_file_btn_2 = Button(file_frame, text="Select File", command=UploadDPP, width=20)
dpp_sheet_display = Label(file_frame, textvariable=dpp_sheet_var, font=('Arial', 10))

# Create Label for displaying user messages
message_display = Label(root, textvariable=user_message, font=('Arial', 10))

# Create Submit Button
submit_btn = Button(root, text="Submit", style='Accent.TButton', command=ExecuteScript, width=20)

# Layout with grid
title.grid(row=0, column=0, columnspan=3, sticky='w', padx=20, pady=(20, 10))
description.grid(row=1, column=0, columnspan=3, sticky='w', padx=20, pady=(0, 20))

file_frame.grid(row=2, column=0, columnspan=3, padx=20, pady=(5, 10), sticky='w')

master_sheet_label.grid(row=0, column=0, sticky='w')
select_file_btn_1.grid(row=0, column=1, padx=(10, 20), pady=(5, 10), sticky='w')
master_sheet_display.grid(row=0, column=2, padx=(10, 20), pady=(5, 10), sticky='w')

dpp_sheet_label.grid(row=1, column=0, sticky='w')
select_file_btn_2.grid(row=1, column=1, padx=(10, 20), pady=(5, 10), sticky='w')
dpp_sheet_display.grid(row=1, column=2, padx=(10, 20), pady=(5, 10), sticky='w')

submit_btn.grid(row=3, column=0, columnspan=3, pady=(20, 10), padx=20, sticky='w')
message_display.grid(row=4, column=0, columnspan=3, pady=(10, 20), padx=20, sticky='w')

# Explanation and FAQ lists
explanation_title = "After uploading both files and submitting them, the program will do the following: "
explanation_points = [
    "Creates a temporary copy of the Master Excel sheet, preventing any data loss.",
    "Format the raw Dell Partner Portal Excel sheet to match the Master Excel sheet template.",
    "Compares the dispatch numbers of both sheets, only selecting not present in your Master Excel sheet.",
    "Selected dispatch number rows are then analyzed; if there is any missing information, a red fill is applied to those cells.",
    "A green fill is applied to all other cells, for easy identification of the new entries.",
    "Dispatch number rows are appended to the bottom of the Master Excel sheet.",
    "The program successfully completes, and the temporary copy is removed.",
    "The changes to the Master Excel sheet are saved.",
    "Open your Master Excel sheet, add filters, and sort by Dispatch Number in descending order"
]

faq_list = {
    "What happens to previous Master Excel sheet entries?": "Nothing. Only new additions are styled; all other data is not modified.",
    "Where do the new entries (work orders) go?": "New work orders are added to the bottom of the Master Excel sheet by default. Filter the Dispatch Numbers in descending order to view the new additions."
}

# Create Frames for Explanation and FAQ
explanation_frame = Frame(root)
explanation_frame.grid(row=5, column=0, columnspan=3, padx=20, pady=(10, 10), sticky='w')

faq_frame = Frame(root)
faq_frame.grid(row=6, column=0, columnspan=3, padx=20, pady=(10, 10), sticky='w')

def CreateSteps(frame, title, points):
    steps_title_label = Label(frame, text=title, font=('Arial', 14, 'bold'))
    steps_title_label.pack(anchor='w', pady=(10, 5))
    for index, point in enumerate(points, start=1):
        step_label = Label(frame, text=f"{index}. {point}", font=('Arial', 10))
        step_label.pack(anchor='w')

def CreateFAQ(frame, faq_dict):
    faq_title_label = Label(frame, text="FAQ", font=('Arial', 14, 'bold'))
    faq_title_label.pack(anchor='w', pady=(10, 5))
    for index, (question, answer) in enumerate(faq_dict.items(), start=1):
        question_label = Label(frame, text=f"{index}. {question}", font=('Arial', 10, 'bold'))
        question_label.pack(anchor='w')
        answer_label = Label(frame, text=f"   {answer}", font=('Arial', 10))
        answer_label.pack(anchor='w', pady=(0, 5))


CreateSteps(explanation_frame, explanation_title, explanation_points)
CreateFAQ(faq_frame, faq_list)

root.mainloop()
