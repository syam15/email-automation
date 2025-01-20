import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from tkinter import Tk, Label, Button, filedialog, Text, Scrollbar, END, Frame, Entry
from tkinter import ttk

# Email server configuration
SERVER = 'smtp.gmail.com'
PORT = 587
FROM = 'email smtp anda'
PASS = 'password smtp anda'

# Variables for file path, message content, and subject
file_path = None
custom_message = ""
custom_subject = ""

def send_emails():
    global file_path, custom_message, custom_subject
    try:
        if not file_path:
            status_text.insert(END, "Please select a file first!\n")
            return

        if not custom_subject.strip():
            status_text.insert(END, "Please enter a subject for the emails!\n")
            return

        if not custom_message.strip():
            status_text.insert(END, "Please enter a message to send!\n")
            return

        # Initialize SMTP server
        server = smtplib.SMTP(SERVER, PORT)
        server.set_debuglevel(0)
        server.ehlo()
        server.starttls()
        server.login(FROM, PASS)

        # Read the Excel file
        email_list = pd.read_excel(file_path)
        names = email_list['Full name']
        emails = email_list['Email Address']

        # Iterate and send emails
        for i in range(len(emails)):
            name = names[i]
            email = emails[i]

            # Email content
            msg = MIMEMultipart()
            msg['Subject'] = custom_subject
            msg['From'] = FROM
            msg['To'] = email

            personalized_message = custom_message.replace("{name}", name)
            msg.attach(MIMEText(personalized_message, 'plain'))
            server.sendmail(FROM, email, msg.as_string())

        server.quit()
        status_text.insert(END, "Emails sent successfully!\n")

    except Exception as e:
        status_text.insert(END, f"Error: {str(e)}\n")

def browse_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_label.config(text=f"Selected file: {file_path}")

def set_message():
    global custom_message
    custom_message = message_entry.get("1.0", END).strip()

def set_subject():
    global custom_subject
    custom_subject = subject_entry.get().strip()

def exit_app():
    root.destroy()

# GUI Setup
root = Tk()
root.title("Email Automation")
root.geometry("700x600")
root.configure(bg="#f0f0f0")

# Header Frame
header_frame = Frame(root, bg="#4CAF50", height=50)
header_frame.pack(fill="x")
header_label = Label(header_frame, text="Email Automation Tool", bg="#4CAF50", fg="white", font=("Arial", 18))
header_label.pack(pady=10)

# File Selection Frame
file_frame = Frame(root, bg="#f0f0f0")
file_frame.pack(pady=10)
file_label = Label(file_frame, text="Select an Excel file with email data:", bg="#f0f0f0", font=("Arial", 12))
file_label.pack(side="left", padx=5)
browse_button = Button(file_frame, text="Browse", command=browse_file, bg="#2196F3", fg="white", font=("Arial", 10), width=10)
browse_button.pack(side="left", padx=5)

# Subject Entry Frame
subject_frame = Frame(root, bg="#f0f0f0")
subject_frame.pack(pady=10)
subject_label = Label(subject_frame, text="Enter the email subject:", bg="#f0f0f0", font=("Arial", 12))
subject_label.pack(anchor="w", padx=5)
subject_entry = Entry(subject_frame, font=("Arial", 10), width=60)
subject_entry.pack(pady=5)
set_subject_button = Button(subject_frame, text="Set Subject", command=set_subject, bg="#2196F3", fg="white", font=("Arial", 10), width=15)
set_subject_button.pack(pady=5)

# Message Entry Frame
message_frame = Frame(root, bg="#f0f0f0")
message_frame.pack(pady=10)
message_label = Label(message_frame, text="Enter your message below (use {name} for recipient names):", bg="#f0f0f0", font=("Arial", 12))
message_label.pack(anchor="w", padx=5)
message_entry = Text(message_frame, wrap='word', height=8, width=70, font=("Arial", 10))
message_entry.pack(pady=5)
set_message_button = Button(message_frame, text="Set Message", command=set_message, bg="#2196F3", fg="white", font=("Arial", 10), width=15)
set_message_button.pack(pady=5)

# Action Buttons Frame
action_frame = Frame(root, bg="#f0f0f0")
action_frame.pack(pady=10)
send_button = Button(action_frame, text="Send Emails", command=send_emails, bg="#4CAF50", fg="white", font=("Arial", 12), width=15)
send_button.grid(row=0, column=0, padx=10)
exit_button = Button(action_frame, text="Exit", command=exit_app, bg="#F44336", fg="white", font=("Arial", 12), width=15)
exit_button.grid(row=0, column=1, padx=10)

# Status Frame
status_frame = Frame(root, bg="#f0f0f0")
status_frame.pack(pady=10)
status_label = Label(status_frame, text="Status:", bg="#f0f0f0", font=("Arial", 12))
status_label.pack(anchor="w", padx=5)
status_text = Text(status_frame, wrap='word', height=10, width=80, font=("Arial", 10), bg="#e8f5e9")
status_text.pack(pady=5)
scroll = Scrollbar(status_frame, command=status_text.yview)
status_text.configure(yscrollcommand=scroll.set)
scroll.pack(side="right", fill="y")

# Run the application
root.mainloop()
