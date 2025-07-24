from pathlib import Path
import os
import sys
import webbrowser
import tkinter as tk
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage, filedialog, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os.path import basename
import smtplib
import pandas as pd


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def open_file():
    attach_file_path = filedialog.askopenfilename()
    if attach_file_path:
        attach_entry.delete(0, tk.END)
        attach_entry.insert(0, attach_file_path)

def open_file_2():
    sender_file_path = filedialog.askopenfilename()
    if sender_file_path:
        sender_csv.delete(0, tk.END)
        sender_csv.insert(0, sender_file_path)

def helper():
    url = "https://docs.google.com/document/d/1v6Q8-xEABRbnzSPs7zMmxkq8E_9U8OWaM0ax10JI0HI/edit?tab=t.0#heading=h.aknepyjh3ttg"
    webbrowser.open(url)

def launch():
    email = emailid.get().strip()
    password = app_password.get().strip().replace(" ", "")
    subject = subject_entry.get().strip()
    body = body_entry.get("1.0", "end").strip()
    sender_csv_path = sender_csv.get().strip()
    resume = attach_entry.get().strip()

    # ✅ Input validation
    if not all([email, password, subject, body, sender_csv_path, resume]):
        messagebox.showerror("Input Error", "All fields are required.")
        return

    if not os.path.isfile(sender_csv_path):
        messagebox.showerror("File Error", f"CSV file not found:\n{sender_csv_path}")
        return

    if not os.path.isfile(resume):
        messagebox.showerror("File Error", f"Attachment file not found:\n{resume}")
        return

    try:
        df = pd.read_csv(sender_csv_path)
    except Exception as e:
        messagebox.showerror("CSV Error", f"Could not read CSV:\n{e}")
        return

    data_list = df.values.tolist()

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email, password)

        for data in data_list:
            to_email = data[0]
            to_name = data[1] if len(data) > 1 else ""
            to_company_name = data[2] if len(data) > 2 else ""

            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = to_email
            msg['Subject'] = subject

            # ✅ Format the HTML email body
            try:
                formatted_body = body.format(name=to_name, company_name=to_company_name)
            except Exception as e:
                messagebox.showerror("Template Error", f"Body formatting failed:\n{e}")
                print("Here is the error !")
                return

            msg.attach(MIMEText(formatted_body, 'html'))

            # ✅ Attachment handling
            try:
                with open(resume, 'rb') as file:
                    filename = basename(resume)
                    part = MIMEApplication(file.read(), Name=filename)
                    part['Content-Disposition'] = f'attachment; filename="{filename}"'
                    msg.attach(part)
            except Exception as e:
                messagebox.showerror("Attachment Error", f"Failed to attach file:\n{e}")
                return

            try:
                server.send_message(msg)
                print(f"Email sent to {to_email}")
            except Exception as e:
                print(f"Failed to send email to {to_email}: {e}")

        messagebox.showinfo("Success", "All HTML emails sent successfully.")

    except Exception as e:
        messagebox.showerror("SMTP Error", f"SMTP setup failed:\n{e}")
    finally:
        try:
            server.quit()
        except:
            pass


# GUI Setup
window = Tk()
window.geometry("1046x804")
window.configure(bg="#FFFFFF")

icon_path = resource_path("emailapp.ico")
if os.path.exists(icon_path):
    window.iconbitmap(icon_path)

canvas = Canvas(
    window, bg="#FFFFFF", height=804, width=1046,
    bd=0, highlightthickness=0, relief="ridge"
)

canvas.place(x=0, y=0)
canvas.create_rectangle(50.0, 72.0, 996.0, 691.0, fill="#072F5F", outline="")
canvas.create_text(454.0, 72.0, anchor="nw", text="AutoEmail", fill="#FFFFFF", font=("IstokWeb Regular", 30 * -1))

# Load assets with resource_path
emailid_image = PhotoImage(file=resource_path(os.path.join("assets", "entry_1.png")))
emailid_bg = canvas.create_image(177.5, 181.0, image=emailid_image)
emailid = Entry(bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0)
emailid.place(x=73.0, y=166.0, width=209.0, height=28.0)

app_password_image = PhotoImage(file=resource_path(os.path.join("assets", "entry_2.png")))
app_password_bg = canvas.create_image(465.5, 181.0, image=app_password_image)
app_password = Entry(bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0, show="*")
app_password.place(x=361.0, y=166.0, width=209.0, height=28.0)

sender_csv_image = PhotoImage(file=resource_path(os.path.join("assets", "entry_3.png")))
sender_csv_bg = canvas.create_image(773.5, 181.0, image=sender_csv_image)
sender_csv = Entry(bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0)
sender_csv.place(x=669.0, y=166.0, width=209.0, height=28.0)

subject_image = PhotoImage(file=resource_path(os.path.join("assets", "entry_4.png")))
subject_bg = canvas.create_image(178.5, 291.5, image=subject_image)
subject_entry = Entry(bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0)
subject_entry.place(x=74.0, y=277.0, width=209.0, height=27.0)

attach_entry_image = PhotoImage(file=resource_path(os.path.join("assets", "entry_5.png")))
attach_entry_bg = canvas.create_image(465.5, 291.5, image=attach_entry_image)
attach_entry = Entry(bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0)
attach_entry.place(x=361.0, y=277.0, width=209.0, height=27.0)

body_image = PhotoImage(file=resource_path(os.path.join("assets", "entry_6.png")))
body_bg = canvas.create_image(406.0, 511.5, image=body_image)
body_entry = Text(bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0)
body_entry.place(x=74.0, y=358.0, width=664.0, height=305.0)

canvas.create_text(362.0, 136.0, anchor="nw", text="App Password", fill="#FFFFFF", font=("IstokWeb Regular", 18 * -1))
canvas.create_text(74.0, 136.0, anchor="nw", text="Email Address", fill="#FFFFFF", font=("IstokWeb Regular", 18 * -1))
canvas.create_text(74.0, 244.0, anchor="nw", text="Email Subject", fill="#FFFFFF", font=("IstokWeb Regular", 18 * -1))
canvas.create_text(362.0, 244.0, anchor="nw", text="Attach Resume", fill="#FFFFFF", font=("IstokWeb Regular", 18 * -1))
canvas.create_text(73.0, 325.0, anchor="nw", text="Email Body", fill="#FFFFFF", font=("IstokWeb Regular", 18 * -1))
canvas.create_text(669.0, 136.0, anchor="nw", text="Sender Details CSV", fill="#FFFFFF", font=("IstokWeb Regular", 18 * -1))

button_image_1 = PhotoImage(file=resource_path(os.path.join("assets", "button_1.png")))
button_1 = Button(image=button_image_1, borderwidth=0, highlightthickness=0, command=launch, relief="flat")
button_1.place(x=794.0, y=499.0, width=156.0, height=82.0)

button_image_2 = PhotoImage(file=resource_path(os.path.join("assets", "button_2.png")))
button_2 = Button(image=button_image_2, borderwidth=0, highlightthickness=0, command=helper, relief="flat")
button_2.place(x=790.0, y=391.0, width=160.0, height=91.0)

button_image_3 = PhotoImage(file=resource_path(os.path.join("assets", "button_3.png")))
button_3 = Button(image=button_image_3, borderwidth=0, highlightthickness=0, command=open_file, relief="flat")
button_3.place(x=492.0, y=239.0, width=47.0, height=35.0)

button_image_4 = PhotoImage(file=resource_path(os.path.join("assets", "button_4.png")))
button_4 = Button(image=button_image_4, borderwidth=0, highlightthickness=0, command=open_file_2, relief="flat")
button_4.place(x=838.0, y=128.0, width=47.0, height=35.0)

window.resizable(False, False)
window.mainloop()
