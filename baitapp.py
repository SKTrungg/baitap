import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime
import os
import pandas as pd

DATA_FILE = "employee_data.csv"

def save_to_file(data):
    file_exists = os.path.isfile(DATA_FILE)
    with open(DATA_FILE, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(["Ma", "Ten", "Don vi", "Chuc danh", "Ngay sinh", "Gioi tinh", "So CMND", "Noi cap", "Ngay cap"])
        writer.writerow(data)

def load_file():
    if not os.path.isfile(DATA_FILE):
        return []
    with open(DATA_FILE, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        return list(reader)

def find_birthdays_today():
    employees = load_file()
    today = datetime.datetime.now().strftime('%d/%m/%Y')
    today_birthdays = [emp for emp in employees if emp['Ngay sinh'] == today]
    return today_birthdays

def export_data_to_excel():
    employees = load_file()
    if not employees:
        messagebox.showerror("Loi", "Khong co du lieu de xuat!")
        return
    df = pd.DataFrame(employees)
    df['Ngay sinh'] = pd.to_datetime(df['Ngay sinh'], format='%d/%m/%Y', errors='coerce')
    df = df.sort_values(by='Ngay sinh', ascending=True)
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if output_file:
        df.to_excel(output_file, index=False, encoding='utf-8')
        messagebox.showinfo("Thanh cong", f"Xuat du lieu thanh cong vao {output_file}")

def save_employee():
    data = [
        emp_id.get(),
        emp_name.get(),
        emp_unit.get(),
        emp_position.get(),
        emp_dob.get(),
        emp_gender.get(),
        emp_id_number.get(),
        emp_issue_place.get(),
        emp_issue_date.get()
    ]
    if "" in data:
        messagebox.showerror("Loi", "Vui long dien day du thong tin!")
        return
    save_to_file(data)
    messagebox.showinfo("Thanh cong", "Luu thong tin nhan vien thanh cong!")
    clear_inputs()

def show_today_birthdays():
    birthdays = find_birthdays_today()
    if not birthdays:
        messagebox.showinfo("Thong bao", "Khong co nhan vien nao sinh nhat hom nay.")
        return
    result = "\n".join([f"{emp['Ma']} - {emp['Ten']}" for emp in birthdays])
    messagebox.showinfo("Sinh nhat hom nay", result)

def clear_inputs():
    emp_id.set("")
    emp_name.set("")
    emp_unit.set("")
    emp_position.set("")
    emp_dob.set("")
    emp_gender.set("")
    emp_id_number.set("")
    emp_issue_place.set("")
    emp_issue_date.set("")

root = tk.Tk()
root.title("Quan ly thong tin nhan vien")
root.geometry("600x400")

emp_id = tk.StringVar()
emp_name = tk.StringVar()
emp_unit = tk.StringVar()
emp_position = tk.StringVar()
emp_dob = tk.StringVar()
emp_gender = tk.StringVar()
emp_id_number = tk.StringVar()
emp_issue_place = tk.StringVar()
emp_issue_date = tk.StringVar()

tk.Label(root, text="Ma:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_id).grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Ten:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_name).grid(row=0, column=3, padx=5, pady=5)

tk.Label(root, text="Don vi:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_unit).grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Chuc danh:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_position).grid(row=1, column=3, padx=5, pady=5)

tk.Label(root, text="Ngay sinh (dd/mm/yyyy):").grid(row=2, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_dob).grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Gioi tinh:").grid(row=2, column=2, padx=5, pady=5, sticky='w')
tk.OptionMenu(root, emp_gender, "Nam", "Nu").grid(row=2, column=3, padx=5, pady=5)

tk.Label(root, text="So CMND:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_id_number).grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Noi cap:").grid(row=3, column=2, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_issue_place).grid(row=3, column=3, padx=5, pady=5)

tk.Label(root, text="Ngay cap (dd/mm/yyyy):").grid(row=4, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_issue_date).grid(row=4, column=1, padx=5, pady=5)

tk.Button(root, text="Luu thong tin", command=save_employee).grid(row=5, column=0, columnspan=2, pady=10)
tk.Button(root, text="Sinh nhat hom nay", command=show_today_birthdays).grid(row=5, column=2, pady=10)
tk.Button(root, text="Xuat danh sach", command=export_data_to_excel).grid(row=5, column=3, pady=10)

root.mainloop()
