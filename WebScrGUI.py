import json
import uuid
from tkinter import *
from tkinter import messagebox
from tkinter import StringVar
from tkinter import ttk
from SBIfunctions import *

root = Tk()


def login():
    login_sbi(userValue.get(), passVale.get(), browser.get())


def exit():
    root.destroy()
    try:
        exit_sbi()
    except IndexError:
        messagebox.showerror(title="Erorr", message="Website not open")


def prep_chq():
    prepare_chq(trPwd)


def savedata():
    try:
        with open("records.json", "r") as json_file:
            feeds = json.load(json_file)
        with open("records.json", "w") as f:
            entry = {
                "username": userValue.get(),
                "Password": passVale.get(),
                "trpwd": trPwd.get()
            }
            useq = usrTyp.get()
            feeds[useq] = entry
            json.dump(feeds, f)
    except:
        with open("records.json", "w") as f:
            json.dump({}, f)
        with open("records.json", "r") as jf:
            data_in = json.load(jf)
        with open("records.json", "w") as f:
            entry = {
                "username": userValue.get(),
                "Password": passVale.get(),
                "trpwd": trPwd.get()
            }
            data_in[usrTyp.get()] = entry
            json.dump(data_in, f)
    UserEntry.delete(0, END)
    passEntry.delete(0, END)
    tpassEntry.delete(0, END)
    usrtypEntry.delete(0, END)


def loaddata(*args):
    inVal= Combo1.get()
    try:
        with open("records.json", 'r') as f:
            mydict1 = json.load(f)
            for id, info in mydict1.items():
                for key in info:
                    if inVal in info[key]:
                        usrtyp = id
                        name = info["username"]
                        pwd = info["Password"]
                        trn_pwd = info["trpwd"]
            usrtypEntry.delete(0, END)
            usrtypEntry.insert(0, usrtyp)
            UserEntry.delete(0, END)
            UserEntry.insert(0, name)
            passEntry.delete(0, END)
            passEntry.insert(0, pwd)
            tpassEntry.delete(0, END)
            tpassEntry.insert(0, trn_pwd)
    except IndexError:
        messagebox.showinfo(title="Error", message="File has no Data")


def readdata():
    try:
        with open("records.json", 'r') as f:
            mydict1 = json.load(f)
            name_lst = []
            for id, info in mydict1.items():
                for key in info:
                    if key == "username":
                        name = info["username"]
                        name_lst.append(name)
        Combo1['values'] = name_lst
    except IndexError:
        print("Err")
        messagebox.showinfo(title="Error", message="File has no Data")


def delentry():
    try:
        with open("records.json", 'r') as f:
            mydict1 = json.load(f)
            x = mydict1.pop("user2")
        with open("records.json", "w") as jf:
            json.dump(mydict1, jf)
    except IndexError:
        print("Err")


def pendingCh():
    pending_ent()
    messagebox.showinfo(title="Information", message="Done!")


def stmt_SBI():
    statement_SBI()

root.geometry("600x400")

root.title("SBI Helper")
f1 = Frame(root, bg="white", borderwidth=1, relief=RAISED)
f1.pack(side=TOP, fill="x")
header = Label(f1, text="Welcome to SBI Helper tool", font="Helvetica 26 bold").pack()
Label(root, text="Enter Login User").place(x=10, y=95)
Label(root, text="Enter Login password").place(x=10, y=125)
Label(root, text="Enter Trns Password").place(x=10, y=155)
Label(root, text="Choose Browser").place(x=10, y=185)
Label(root, text="Choose User").place(x=10, y=215)
Label(root, text="Enter User Type").place(x=10, y=65)

userValue = StringVar()
passVale = StringVar()
trPwd = StringVar()
usrTyp = StringVar()

usrtypEntry = Entry(root, textvariable=usrTyp)
usrtypEntry.place(x=150, y=65)
UserEntry = Entry(root, textvariable=userValue)
UserEntry.place(x=150, y=95)
passEntry = Entry(root, textvariable=passVale, show="*")
passEntry.place(x=150, y=125)
tpassEntry = Entry(root, textvariable=trPwd, show="*")
tpassEntry.place(x=150, y=155)

browser = IntVar()
R1 = Radiobutton(root, text="MS Edge", variable=browser, value=1).place(x=150, y=185)
R2 = Radiobutton(root, text="Google Chrome", variable=browser, value=2).place(x=150, y=215)

Combo1 = ttk.Combobox(root)
Combo1.place(x=150, y=275)
Combo1.bind("<<ComboboxSelected>>", loaddata)

Btn1 = Button(root, text="Login", command=login).place(x=150, y=245)
Btn2 = Button(root, text="Save Credentials", command=savedata).place(x=300, y=65)
Btn3 = Button(root, text="Load Credentials", command=loaddata).place(x=300, y=95)
Btn4 = Button(root, text="Exit", command=exit).place(x=200, y=245)
Btn5 = Button(root, text="Load Pending Entries", command=pendingCh).place(x=450, y=65)
Btn6 = Button(root, text="Prepare Cheques", command=prep_chq).place(x=450, y=95)
Btn7 = Button(root, text="Load Statement", command=stmt_SBI).place(x=450, y=125)
Btn8 = Button(root, text="Load Users", command=readdata).place(x=150, y=305)
Btn9 = Button(root, text="Remove Users", command=delentry).place(x=250, y=305)

readdata()

mainloop()
