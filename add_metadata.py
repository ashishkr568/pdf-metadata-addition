"""
Code to Update Properties (Metadata) of a PDF file
App Name: PDF Metadata Addition Tool
Version: 1.0
Author: Ashish Kumar
Date: 13-Jan-2022
LinkedIn: : https://www.linkedin.com/in/ashish568/
Medium: https://medium.com/@ashish.568
GitHub: https://github.com/ashishkr568

.exe Creation Guide:
Run the Below command in Powershell to create .exe for this Application.
"pyinstaller add_metadata.py"

Note: The .exe file should always be with the img folder else the application won't work. This is a limitation of
tKinter as there is no way to map image without giving path"

"""

# Import Packages
import tkinter as tk
import tkinter.filedialog
# from PyPDF2 import PdfFileReader
import webbrowser

from PyPDF2 import PdfFileMerger

# from tkinter import messagebox

# Create a Tkinter App, set canvas size and set its icon and title
app = tk.Tk()
app.geometry("650x600")
#app.configure(background='#0077b5')
app.iconbitmap("img/appIcon.ico")
app.title("PDF Metadata Addition Tool")
# app.eval('tk::PlaceWindow . center')
app.resizable(False, False)


# Function to save output pdf file
def add_metadata(val_dict):
    # Read File from the Directory
    file_in = open(val_dict["inp_path"], 'rb')
    # pdf_reader = PdfFileReader(file_in)
    # metadata = pdf_reader.getDocumentInfo()

    # Add Metadata
    pdf_merger = PdfFileMerger()
    pdf_merger.append(file_in)
    pdf_merger.addMetadata({
        '/Title': val_dict["title"],
        '/Author': val_dict["author"],
        '/Subject': val_dict["subject"],
        '/Created': val_dict["created"],
        '/Modified': val_dict["modified"],
        '/Application': val_dict["application"],
        '/Keywords': val_dict["keywords"]
    })

    # Save File at output location

    try:
        file_out = open(val_dict["out_dir"] + '/' + val_dict["new_f_name"] + '.pdf', 'wb')
        pdf_merger.write(file_out)
        # Close Files
        file_in.close()
        file_out.close()

    except PermissionError:
        tk.messagebox.showerror('Error',
                                'Error: There is already an existing file with the same name and its open,'
                                ' Please Close it and try again..!!')
        val_dict = fetch_values()
        add_metadata(val_dict)
        tk.messagebox.showinfo("Message", "PDF Metadata Added..!!")

        # Delete all filled values from form
        e1.delete("0", "end")
        e2.delete("0", "end")
        e3.delete("0", "end")
        e5.delete("0", "end")
        e6.delete("0", "end")
        e7.delete("0", "end")
        e8.delete("0", "end")
        e9.delete("0", "end")
        e10.delete("0", "end")
        e11.delete("1.0", "end")


# On Click Message box popup
def on_add_metadata_click():
    val_dict = fetch_values()
    if val_dict["inp_path"] == "" or val_dict["out_dir"] == "" or val_dict["new_f_name"] == "":
        if val_dict["inp_path"] == "":
            tk.messagebox.showerror('Error', 'Error: Please Enter value for Input PDF Path')
        # Check if Output Path is empty
        if val_dict["out_dir"] == "":
            tk.messagebox.showerror('Error', 'Error: Please Enter value for Output Path..!!')

        # Check if Output File Name is empty
        if val_dict["new_f_name"] == "":
            tk.messagebox.showerror('Error', 'Error: Please Enter value for Output File Name..!!')
    else:
        val_dict = fetch_values()
        add_metadata(val_dict)
        tk.messagebox.showinfo("Message", "PDF Metadata Added..!!")

        # Delete all filled values from form
        e1.delete("0", "end")
        e2.delete("0", "end")
        e3.delete("0", "end")
        e5.delete("0", "end")
        e6.delete("0", "end")
        e7.delete("0", "end")
        e8.delete("0", "end")
        e9.delete("0", "end")
        e10.delete("0", "end")
        e11.delete("1.0", "end")


# Function to update properties(Metadata) of a pdf file
def fetch_values():
    # Fetch values from Form filled on Tkinter app
    inp_path = str(e1.get())
    out_dir = str(e2.get())

    # Check if User has provided .pdf in file name, If yes then strip it
    new_f_name = e3.get()
    new_f_name = new_f_name.replace(".pdf", "")

    title = str(e5.get())
    author = str(e6.get())
    subject = str(e7.get())
    created = str(e8.get())
    modified = str(e9.get())
    application = str(e10.get())
    keywords = str(e11.get("1.0", 'end-1c'))

    # Add these values to a Dictionary
    val_dict = {"inp_path": inp_path,
                "out_dir": out_dir,
                "new_f_name": new_f_name,
                "title": title,
                "author": author,
                "subject": subject,
                "created": created,
                "modified": modified,
                "application": application,
                "keywords": keywords}
    return val_dict


# Create Form in Tkinter App using Grid

# Row = 0

# Create a Frame to enter App Description -- Column = All
app_intro = "This tool will update Properties (Metadata) of PDF File"

# Create App Description Label
app_intro_frame = tk.Frame(app, width=600, height=50)
app_intro_frame.grid(row=0, column=0, columnspan=3)
app_intro_frame.grid_propagate(False)
app_intro_frame.update()
app_intro_label = tk.Label(app_intro_frame, text=app_intro, wraplength=400, font=("Arial Bold", 10, "italic"))
app_intro_label.place(x=325, y=20, anchor="center")

# Take File Information from User

# Row = 1

# Input File Location

# Define Label -- Column = 0
tk.Label(app, text="Input PDF Path").grid(row=1, column=0, ipadx=35, ipady=5, sticky=tk.W)

# Define TextBox -- Column = 1
e1 = tk.Entry(app, state='disabled')
e1.grid(row=1, column=1, padx=4, ipadx=111, ipady=3)


# Define a Browse Button for Input File Selection -- column = 2
# Function to open a new dialog and select pdf file
def inp_browse():
    inp_filename = tkinter.filedialog.askopenfiles(title="Select a File", filetypes=(("pdf Files", "*.pdf"),))
    # print(str(inp_filename[0].name))
    e1.config(state='normal')
    e1.insert(0, str(inp_filename[0].name))
    e1.config(state='disabled')


# Function to Update color on hover for input browse button
def inp_button_on_enter(e):
    inp_browse_button['background'] = 'Black'
    inp_browse_button['foreground'] = 'white'


def inp_button_on_leave(e):
    inp_browse_button['background'] = 'SystemButtonFace'
    inp_browse_button['foreground'] = 'black'


# Create Input Browse Button
inp_browse_button = tkinter.Button(app, text="Browse", command=inp_browse)
inp_browse_button.grid(row=1, column=2)

# Add Mouse Hover Effect
inp_browse_button.bind("<Enter>", inp_button_on_enter)
inp_browse_button.bind("<Leave>", inp_button_on_leave)

# Row = 2

# Output File Path -- Column = 0
# Define Label -- Column = 0
tk.Label(app, text="Output Path").grid(row=2, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e2 = tk.Entry(app, state='disabled')
e2.grid(row=2, column=1, padx=4, ipadx=111, ipady=3)


# Define a Browse Button To fetch Output Directory Location
# Function to Open a dialog for selecting output location
def out_dir_browse():
    out_dir = tkinter.filedialog.askdirectory(title="Select Output Directory")
    # print(out_dir)
    e2.config(state='normal')
    e2.insert(0, str(out_dir))
    e2.config(state='disabled')


# Function to Update color on hover for Output browse button
def out_button_on_enter(e):
    out_browse_button['background'] = 'Black'
    out_browse_button['foreground'] = 'white'


def out_button_on_leave(e):
    out_browse_button['background'] = 'SystemButtonFace'
    out_browse_button['foreground'] = 'black'


# Create Button for selecting output location
out_browse_button = tkinter.Button(app, text="Browse", command=out_dir_browse)
out_browse_button.grid(row=2, column=2)

# Add Mouse Hover Effect
out_browse_button.bind("<Enter>", out_button_on_enter)
out_browse_button.bind("<Leave>", out_button_on_leave)

# Row = 3

# Output File Name
# Define Label -- Column = 0
tk.Label(app, text="Output File Name").grid(row=3, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e3 = tk.Entry(app)
e3.grid(row=3, column=1, padx=4, ipadx=111, ipady=3)

# Row = 4

# Create a frame to include text -- Column = All
app_metadata_frame = tk.Frame(app, width=600, height=50)
app_metadata_frame.grid(row=4, column=0, columnspan=3)
app_metadata_frame.grid_propagate(False)
app_metadata_frame.update()
app_metadate_label = tk.Label(app_metadata_frame, text="Enter Metadata", wraplength=400,
                              font=("Arial Bold", 10, "italic", "underline"))
app_metadate_label.place(x=350, y=30, anchor="center")

# Row = 5

# Title
# Define Label -- Column = 0
tk.Label(app, text="Title").grid(row=5, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e5 = tk.Entry(app)
e5.grid(row=5, column=1, padx=4, ipadx=111, ipady=3)

# Row = 6

# Author
# Define Label -- Column = 0
tk.Label(app, text="Author").grid(row=6, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e6 = tk.Entry(app)
e6.grid(row=6, column=1, padx=4, ipadx=111, ipady=3)

# Row = 7

# Subject
# Define Label -- Column = 0
tk.Label(app, text="Subject").grid(row=7, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e7 = tk.Entry(app)
e7.grid(row=7, column=1, padx=4, ipadx=111, ipady=3)

# Row = 8

# Created Date
# Define Label -- Column = 0
tk.Label(app, text="Created Date").grid(row=8, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e8 = tk.Entry(app)
e8.grid(row=8, column=1, padx=4, ipadx=111, ipady=3)

# Row = 9

# Modified Date
# Define Label -- Column = 0
tk.Label(app, text="Modified Date").grid(row=9, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e9 = tk.Entry(app)
e9.grid(row=9, column=1, padx=4, ipadx=111, ipady=3)

# Row = 10

# Application
# Define Label -- Column = 0
tk.Label(app, text="Application").grid(row=10, column=0, ipadx=35, ipady=5, sticky=tk.W)
# Define TextBox -- Column = 1
e10 = tk.Entry(app)
e10.grid(row=10, column=1, padx=4, ipadx=111, ipady=3)

# Row = 11

# Keywords
# Define Label -- Column = 0
tk.Label(app, text="Keywords").grid(row=11, column=0, ipadx=35, ipady=50, sticky=tk.W)
# Define TextBox with a vertical scrollbar -- Column = 1
f = tk.Frame(app)
f.place(x=534, y=394)
scrollbar = tk.Scrollbar(f)
e11 = tk.Text(wrap="word", width=42, height=5, yscrollcommand=scrollbar.set)
scrollbar.config(command=e11.yview)
scrollbar.pack(side="left", fill="x", expand=True, ipady=20, padx=3)
e11.grid(row=11, column=1, padx=4, ipadx=5, ipady=3)

# Row = 12

# Close and Add Metadata Buttons in the same Grid Cell using Frame
# Define Frame
f1 = tk.Frame(app)
f1.grid(row=12, column=1, sticky="nsew")
# Add close button with image -- column = 1
close_button_img = tk.PhotoImage(file="img/close_button.png")
c_button = tk.Button(f1, text="Exit Application", image=close_button_img, borderwidth=0, cursor='hand2',
                     command=app.quit)

# Add "Add Metadata" Button with color change on hover
run_b2 = tk.Button(f1, text='Add Metadata', cursor='hand2', command=on_add_metadata_click)


# Color change on hover for Add Metadata Button
def metadate_button_on_enter(e):
    run_b2['background'] = 'Black'
    run_b2['foreground'] = 'white'


def metadata_button_on_leave(e):
    run_b2['background'] = 'SystemButtonFace'
    run_b2['foreground'] = 'black'


run_b2.bind("<Enter>", metadate_button_on_enter)
run_b2.bind("<Leave>", metadata_button_on_leave)

# Arrange Button in same grid cell
run_b2.pack(side="right")
c_button.pack(side="right", padx=10)

# Row = 13

# LinkedIn / Medium / GitHub Contacts in same grid cell using frame -- Column
# Define Frame
logo_frame = tk.Frame(app)
logo_frame.grid(row=13, column=0)


# Define a function to open webpage on click
def callback(url):
    webbrowser.open_new_tab(url)


# LinkedIn
linkedin_logo = tk.PhotoImage(file="img/linkedin_logo.png")
l_button = tkinter.Button(logo_frame, text="LinkedIn", image=linkedin_logo, borderwidth=0, cursor='hand2',
                          command=lambda: callback('https://www.linkedin.com/in/ashish568/'))

# GitHub
git_logo = tk.PhotoImage(file="img/github_logo.png")
g_button = tkinter.Button(logo_frame, text="GitHub", image=git_logo, borderwidth=0, cursor='hand2',
                          command=lambda: callback('https://github.com/ashishkr568'))

# Medium
medium_logo = tk.PhotoImage(file="img/medium_logo.png")
m_button = tkinter.Button(logo_frame, text="Medium", image=medium_logo, borderwidth=0, cursor='hand2',
                          command=lambda: callback('https://medium.com/@ashish.568'))

# Arrange Buttons in single grid
l_button.pack(side="left")
g_button.pack(side="left", ipadx=10)
m_button.pack(side="left")

# Close the App
app.mainloop()
