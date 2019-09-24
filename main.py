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

first_names = []
last_names = []
emails = []

window = tk.Tk()
window.title("Mass Email")

data_path = Entry(window)
subject_entry = Entry(window)
data_path.grid(column = 0, columnspan = 2, row = 0)
subject_entry.grid(column = 0, columnspan = 2, row = 1)
subject_entry.insert(END, "Subject")

body = scrolledtext.ScrolledText(window, height = 20)
body.grid(column = 0, columnspan = 4, row = 2)
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

def exit_clicked():
    sys.exit(0)

def send_email_clicked():
    find_information()
    send_emails()
    print("\n\nCopy of the email sent:\n\n")
    print(body.get(1.0, END))

load_csv_button = Button(window, text = 'LOAD EMAIL CSV', command = browse_clicked)
send_email_buttonn = Button(window, text = 'SEND EMAILS', command = send_email_clicked)
exit_button = Button(window, text = 'CLOSE', fg ='red', command = exit_clicked)


load_csv_button.grid(column = 3, row = 0)
send_email_buttonn.grid(column = 0, row = 3)
exit_button.grid(column = 3, row = 3)

window.mainloop()
