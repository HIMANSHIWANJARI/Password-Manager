from tkinter import *
from tkinter import messagebox
import random
import pyperclip
import json
import os
import openpyxl
from openpyxl import Workbook

# Excel setup
EXCEL_FILE = 'passwords.xlsx'

def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Website", "Email/Username", "Password"])
        wb.save(EXCEL_FILE)

def add_to_excel(data):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append(data)
    wb.save(EXCEL_FILE)

def generate_password():
    letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y',
               'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    symbols = ['!', '#', '$', '%', '&', '(', ')', '*', '+', '/', '@', '^', '}', '{', ':', '~']

    password_letters = [random.choice(letters) for _ in range(random.randint(8, 10))]
    password_numbers = [random.choice(numbers) for _ in range(random.randint(2, 4))]
    password_symbols = [random.choice(symbols) for _ in range(random.randint(2, 4))]

    password_list = password_letters + password_numbers + password_symbols
    random.shuffle(password_list)

    password = "".join(password_list)
    password_input.insert(0, password)
    pyperclip.copy(password)

def save():
    website = website_input.get()
    username = username_input.get()
    password = password_input.get()
    new_data = {
        website: {
            "username": username,
            "password": password
        }
    }

    if len(website) == 0 or len(password) == 0:
        messagebox.showinfo(
            title="Invalid Data", message="Please make sure you haven't left any fields empty.")
    else:
        try:
            with open("data.json", "r") as data_file:
                data = json.load(data_file)
        except FileNotFoundError:
            data = {}
        except json.JSONDecodeError:
            data = {}

        if website in data:
            is_ok = messagebox.askokcancel(title="Update", message=f"Details for {website} already exist. Do you want to update the details?")
            if is_ok:
                data.update(new_data)
        else:
            data.update(new_data)

        with open("data.json", "w") as data_file:
            json.dump(data, data_file, indent=4)

        # Add data to Excel
        add_to_excel([website, username, password])

        website_input.delete(0, END)
        password_input.delete(0, END)

def find_password():
    website = website_input.get()

    try:
        with open("data.json") as data_file:
            data = json.load(data_file)
    except FileNotFoundError:
        messagebox.showinfo(title="Error", message="No data file found.")
        return
    except json.JSONDecodeError:
        messagebox.showinfo(title="Error", message="Data file is empty or corrupted.")
        return

    if website in data:
        email = data[website]["username"]
        password = data[website]["password"]
        messagebox.showinfo(
            title=website, message=f"Email/Username: {email}\nPassword: {password}")
    else:
        messagebox.showinfo(
            title="Error", message=f"No details for {website} exists.")

# UI
window = Tk()
window.title("Password Manager")
window.config(padx=50, pady=50, bg="white")

canvas = Canvas(window, width=500, height=500, bg="white", highlightthickness=0)

# Adjust the path to logo.png if necessary
logo = PhotoImage(file="logo.png")

# Center the image in the canvas
canvas.create_image(250, 250, image=logo)
canvas.grid(column=1, row=0)

#website
website_label = Label(text="Website", bg="white")
website_label.grid(column=0, row=1)
website_input = Entry(width=70)
website_input.focus()
website_input.grid(column=1, row=1)

#Username
username_label = Label(text="Email/Username", bg="white")
username_label.grid(column=0, row=2)
username_input = Entry(width=70)
username_input.grid(column=1, row=2)

#Password
password_label = Label(text="Password", bg="white")
password_label.grid(column=0, row=3)
password_input = Entry(width=70)
password_input.grid(column=1, row=3)

generate_password_button = Button(text="Generate Password", command=generate_password)
generate_password_button.grid(column=2, row=3)

search_button = Button(text="Search", width=15, command=find_password)
search_button.grid(row=1, column=2)

add_button = Button(text="Add Password", width=40, command=save)
add_button.grid(row=4, column=1)

# Initialize the Excel file
initialize_excel()

window.mainloop()
