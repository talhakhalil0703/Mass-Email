import sys
import csv
import os
import smtplib
import tkinter as tk
from tkinter import *
from tkinter import scrolledtext
from tkinter import filedialog
from email.message import EmailMessage

USER = os.environ.get('GMAIL_USER')
PASS = os.environ.get('GMAIL_PASS')
RECIEVER = USER

readFile = None
fileName = None

first_names = []
last_names = []
emails = []

window = tk.Tk()
window.title("Mass Email")

data_path = Entry(window)
subject_entry = Entry(window)
data_path.grid(column = 0, columnspan = 2, row = 0)

subject_entry.grid(column = 0, columnspan = 2, row = 1)
subject_entry.insert(END, "Digitronics: ")

attachment_entry = Entry(window)
attachment_entry.grid(column = 0, columnspan = 2, row = 2)

body = scrolledtext.ScrolledText(window, height = 20)
body.grid(column = 0, columnspan = 4, row = 3)
body.insert(END, 'Thanks, \n\nDigitronics Team')

def find_information():
    if (data_path.get() == None):
        print("No CSV Loaded")
        return
    with open(data_path.get(), mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        line_count = 1
        for row in csv_reader:
            first_names.append(row["First Name"])
            last_names.append(row["Last Name"])
            emails.append(row["Email"])
            line_count += 1
        print(f'Processed {line_count} lines.')

def send_emails():
    if (first_names == None):
        print("No CSV Loaded")
        return
    else:
        for index, name in enumerate(first_names):
            RECIEVER = emails[index]
            msg = EmailMessage()
            msg['Subject'] = subject_entry.get()
            msg['From'] = USER
            msg['To'] = RECIEVER
            content = "Hi " + name.capitalize() + ", \n" + body.get(1.0, END)
            msg.set_content(content)
            global readFile, fileName
            if (fileName != None):
                msg.add_attachment(readFile, maintype = "application", subtype= "octet-stream", filename = fileName)

            with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()

                smtp.login(USER, PASS)
                smtp.send_message(msg)
                print("Email sent to: ", RECIEVER)

def browse_clicked():
    window.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("CSV files","*.csv"), ("all files","*.*")))
    data_path.delete(0, END)
    data_path.insert(END, window.filename)

def attachment_browse_clicked():
    window.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("PDF files","*.pdf"), ("all files","*.*")))
    attachment_entry.delete(0, END)
    attachment_entry.insert(END, window.filename)   
    global readFile, fileName
    with open (attachment_entry.get(), "rb") as f:
        readFile = f.read()
        fileName = f.name


def exit_clicked():
    sys.exit(0)

def send_email_clicked():
    find_information()
    send_emails()
    print("\n\nCopy of the email sent:\n\n")
    print(body.get(1.0, END))

load_csv_button = Button(window, text = 'LOAD EMAIL CSV', command = browse_clicked)
load_attachment_button = Button(window, text = 'LOAD ATTACHMENTS', command = attachment_browse_clicked)

send_email_button = Button(window, text = 'SEND EMAILS', command = send_email_clicked)
exit_button = Button(window, text = 'CLOSE', fg ='red', command = exit_clicked)


load_csv_button.grid(column = 3, row = 0)
load_attachment_button.grid(column = 3, row = 2)
send_email_button.grid(column = 0, row = 4)
exit_button.grid(column = 3, row = 4)

window.mainloop()
