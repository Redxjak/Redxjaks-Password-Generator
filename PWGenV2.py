# Created by Redxjak

# Password Generator


import random
import string
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from openpyxl import Workbook, load_workbook


def generate_password(min_length, max_length, min_uppercase, min_lowercase, min_symbols):
    uppercase = string.ascii_uppercase
    lowercase = string.ascii_lowercase
    numbers = string.digits
    symbols = "!@#$%^&*()"

    characters = uppercase + lowercase + numbers + symbols
    password_characters = []

    length = random.randint(min_length, max_length)

    for i in range(min_uppercase):
        password_characters.append(random.choice(uppercase))

    for i in range(min_lowercase):
        password_characters.append(random.choice(lowercase))

    for i in range(min_symbols):
        password_characters.append(random.choice(symbols))

    minimum_required_length = len(password_characters)

    if minimum_required_length > length:
        length = minimum_required_length

    remaining_length = length - len(password_characters)

    for i in range(remaining_length):
        password_characters.append(random.choice(characters))

    random.shuffle(password_characters)

    return "".join(password_characters)


def choose_file_location():
    file_path = filedialog.asksaveasfilename(
        title="Choose Save Location",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if file_path:
        save_location_entry.delete(0, tk.END)
        save_location_entry.insert(0, file_path)


def create_password():
    site_name = site_entry.get()
    username = username_entry.get()
    save_location = save_location_entry.get()

    try:
        min_length = int(min_length_entry.get() or 10)
        max_length = int(max_length_entry.get() or 20)

        min_uppercase = int(min_uppercase_entry.get() or 0)
        min_lowercase = int(min_lowercase_entry.get() or 0)
        min_symbols = int(min_symbols_entry.get() or 0)
    except ValueError:
        messagebox.showerror(
            "Error",
            "Length and minimum requirements must be numbers."
        )
        return

    if min_length > max_length:
        messagebox.showerror(
            "Error",
            "Minimum length cannot be greater than maximum length."
        )
        return

    if site_name == "" or username == "" or save_location == "":
        messagebox.showerror("Error", "Please fill out all fields.")
        return

    password = generate_password(
        min_length,
        max_length,
        min_uppercase,
        min_lowercase,
        min_symbols
    )

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    password_output.delete(0, tk.END)
    password_output.insert(0, password)

    if os.path.exists(save_location):
        workbook = load_workbook(save_location)
        sheet = workbook.active
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Passwords"
        sheet.append(["Website/Application", "Username", "Password", "Timestamp"])

    sheet.append([site_name, username, password, timestamp])

    try:
        workbook.save(save_location)
    except PermissionError:
        messagebox.showerror(
            "File is open",
            "Please close the Excel file before generating a password."
        )
        return

    messagebox.showinfo("Password Generated", "Password Generated Successfully!")


# GUI Setup
window = tk.Tk()
window.title(" Redxjak's Password Generator")
window.geometry("400x550")

tk.Label(window, text="Redxjak's Password Generator", font=("Arial", 16)).pack(pady=10)

tk.Label(window, text="Website/Application Name:").pack()
site_entry = tk.Entry(window, width=45)
site_entry.pack()

tk.Label(window, text="Username:").pack()
username_entry = tk.Entry(window, width=45)
username_entry.pack()

tk.Label(window, text="Minimum Password Length: (Default 10)").pack()
min_length_entry = tk.Entry(window, width=45)
min_length_entry.pack()

tk.Label(window, text="Maximum Password Length: (Default 20)").pack()
max_length_entry = tk.Entry(window, width=45)
max_length_entry.pack()

tk.Label(window, text="Minimum Uppercase Letters:").pack()
min_uppercase_entry = tk.Entry(window, width=45)
min_uppercase_entry.pack()

tk.Label(window, text="Minimum Lowercase Letters:").pack()
min_lowercase_entry = tk.Entry(window, width=45)
min_lowercase_entry.pack()

tk.Label(window, text="Minimum Special Characters:").pack()
min_symbols_entry = tk.Entry(window, width=45)
min_symbols_entry.pack()

tk.Label(window, text="Save Location:").pack()

tk.Label(
    window,
    text="*Click 'Choose Save Location' and enter the file name, or select previous file to add to it.*",
    wraplength=350,
    justify="center",
    fg="gray",
    font=("Arial", 8)
).pack()

save_location_entry = tk.Entry(window, width=45)
save_location_entry.pack()

browse_button = tk.Button(
    window,
    text="Choose Save Location",
    command=choose_file_location
)
browse_button.pack(pady=5)

generate_button = tk.Button(
    window,
    text="Generate Password",
    command=create_password
)
generate_button.pack(pady=10)

tk.Label(window, text="Generated Password:").pack()
password_output = tk.Entry(window, width=45)
password_output.pack(pady=5)

window.mainloop()