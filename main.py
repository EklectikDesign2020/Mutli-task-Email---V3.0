import time
from tkinter import *
from tkinter import messagebox, filedialog
from PIL import ImageTk, Image
import smtplib
import csv
import threading
import os


# pc profile
username = os.getlogin()
# Avoid not assignment error
reader = '@'

# Window Configuration - Main window -
root = Tk()
root.title("Multi Task E-mailer - MTE ")
root.iconbitmap("MTE_icon.ico")
root.resizable(width=FALSE, height=FALSE)

# Value RadioButton
smtp_server = IntVar()
smtp_server.set("0")                                                                                                    # Set default value of radio Button


#################################################################
# Browse files .csv files.
#    if file is compatible will look for column Email and create a list
# Creates a variable count email equal to the number of email to send.
#    use for appear message box when all email is send
##################################################################
def csv_browser():
    global reader
    global count_email

    # Implement csv
    root.filename = filedialog.askopenfile(initialdir=f"C:/Users/{username}/Documents/csv/", title="select a file",
                                           filetypes=(("png files", "*.csv"), ("all file", "*.csv")))
    try:
        with open(root.filename.name, 'r') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=',')

            for email in reader:
                email_receiver_input.insert(0, email['Email'] + ', ')
                if '@' in email_receiver_input.get():                                                                   # Start to check for valid emails
                    while True:
                        count_email = 1
                        count_email += count_email
                        break
                else:
                    messagebox.showwarning('Email...', 'Non Email(s) find in the file .csv,'                            # Sow warning if no valid emails
                                           ' please check your file and try again.')
    except Exception as error_file:
        messagebox.showerror("Error .csv File...", f"{error_file}")


#################################################################
# Control if sender E-amil and receiver is correct format.
#    If incorrect warning message popup.
# Control by flag Bool.
#    If each true, moves to next function.
##################################################################
def check_email_format():
    flag1 = False
    flag2 = False

    if "@" not in email_input.get():
        messagebox.showwarning("E-mail Format...", "The E-mail(s) entered are incorrect format")
    else:
        flag1 = True
    if "@" not in email_receiver_input.get():
        messagebox.showwarning("E-mail Format...", "The E-mail(s) entered are incorrect format")
    else:
        flag2 = True

    if flag1 and flag2:
        thread()


#################################################################
# Main function is send the emails.
#    Sends once all conditions a met.
# Secondary functions.
#    Calculate speed of value progress bar.
#    Check if amount is digit or return a warning message to user.
#    Display up to date email(s) sent in live time.
#    Display message once all email(s) sent or error - email not sent with exception.
##################################################################
def sending_email():
    global reader
    global count_email

    # Second window
    top = Toplevel()
    top.overrideredirect(True)  # Disable close / and all window base function
    frame_title = LabelFrame(top, text='Email(s) sent', padx=5, pady=5)
    frame_title.pack()
    announce = Label(frame_title, text=" <Sending email>")
    announce.pack()

    with open(root.filename.name, 'r') as csvfile:
        reader = csv.DictReader(csvfile, delimiter=',')

    # Main function Sending email with all info needed (start).
        for e in reader:

            if smtp_server.get() == 1:
                try:
                    # Gmail
                    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
                    server.login(email_input.get(), password_input.get())
                    server.sendmail(email_input.get(), e['Email'] + ', ', message_box.get(1.0, END))
                    server.quit()
                    count_email = count_email - 1

                    email_sent_count = Label(top, text=f" {e['Email'] + ', '} - E-mail sent(s)")
                    email_sent_count.pack()
                    if count_email == 0:
                        time.sleep(10)                                                                                  # Help not to cut the task half way
                        messagebox.showinfo("Email sent", f"Email as been sent to {email_receiver_input.get()}")

                # Display error.
                except Exception as e:
                    messagebox.showerror("Error sending email...", f'{e} - E-mail can not be sent.')
                    break
            else:
                try:
                    # Live
                    server = smtplib.SMTP("smtp.live.com", 587)
                    server.starttls()
                    server.login(email_input.get(), password_input.get())

                    server.sendmail(email_input.get(), e, message_box.get(1.0, END))
                    server.quit()
                    count_email = count_email - 1

                    email_sent_count = Label(top, text=f" {e} - E-mail sent(s)")
                    email_sent_count.grid(row=8, column=0, sticky=W)
                    if count_email == 0:
                        messagebox.showinfo("Email sent", f"Email as been sent to {email_receiver_input.get()}")
                        top.destroy()

                    # Display error.
                except Exception as e:
                    messagebox.showerror("Error sending email...", f'{e} - E-mail can not be sent.')
                    break
        # Main function (end).


#################################################################
# Thread for email sent.
#    Sends the task of sending email to other thread.
#    Program do not freeze.
#
##################################################################
def thread():
    t = threading.Thread(target=sending_email)
    t.setDaemon(True)
    t.start()


# GUI interface (label, entry, image)
bottom_text = Label(root, font='None, 8',
                    text="Copyright - Design and coded by Eklectik Design - Eklectik.design@hotmail.com",
                    justify=CENTER)
bottom_text.grid(row=10, column=0, columnspan=2)
# Image
my_img = Image.open("image_MTE.png")
photo = ImageTk.PhotoImage(my_img)
label = Label(root, image=photo)
label.image = photo
label.grid(row=0, column=0, padx=10, pady=10, sticky=W)
bottom_text = Label(root, font='None, 10', text="<Company Name> MTE - V3.00", justify=CENTER)
bottom_text.grid(row=0, column=0, columnspan=2)

# Radio button
gmail = Radiobutton(root, text="Gmail", variable=smtp_server, value=1)
gmail.grid(row=3, column=1, sticky=W)
live = Radiobutton(root, text="Live", variable=smtp_server, value=2)
live.grid(row=3, column=1)

smtp_label = Label(root, text="                           Email provider :")
smtp_label.grid(row=3, column=0)

# Email text
email_info = Label(root, text="                                Sender E-mail :")
email_info.grid(row=1, column=0, sticky=W)
# Email text box
email_input = Entry(root, borderwidth=2, width=40, )
email_input.grid(row=1, column=1, pady=2, padx=15)

# Password text
password_info = Label(root, text="                           Sender Password :")
password_info.grid(row=2, column=0, sticky=W)
# Password text box
password_input = Entry(root, borderwidth=2, width=40, show='*')
password_input.grid(row=2, column=1, pady=2, padx=5)

# Email of receiver
email_receiver_info = Label(root, text="                                          Receiver :")
email_receiver_info.grid(row=5, column=0, sticky=W)
# Email of receiver text box
email_receiver_input = Entry(root, borderwidth=2, width=40)

email_receiver_input.grid(row=5, column=1, pady=2, padx=5)

# Instructions
instruction = Label(root, text="-- Please don't delete text <Subject:> --"
                               "\nType the subject of the Email the one line down the message.")
instruction.grid(row=7, column=0, columnspan=2)

# Message Box
frame_message = LabelFrame(root, text="Message :", padx=10, pady=7)
frame_message.grid(padx=5, pady=5, row=8, column=0, columnspan=2)
message_box = Text(frame_message, borderwidth=2, width=48, height=10)
message_box.insert(1.0, "Subject: ")
message_box.grid(row=8, column=0, columnspan=2, pady=2, padx=5)

# Buttons
send = Button(root, text="Send Email", command=check_email_format, width=10)
send.grid(row=9, column=0, sticky=W, pady=2, padx=4)
button_quit = Button(root, text="Quit Application", command=root.quit)
button_quit.grid(row=9, column=1, sticky=E, padx=5, pady=3)
browser_button = Button(root, text="Browser CSV", command=csv_browser)
browser_button.grid(row=6, column=1, sticky=W)

mainloop()
