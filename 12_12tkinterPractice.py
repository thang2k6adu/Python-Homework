import tkinter as tk
from tkinter import messagebox
import pandas as pd
import csv
from datetime import datetime

# Khai báo danh sách nhân viên
employee_data = []

# Hàm lưu thông tin nhân viên vào file CSV
def save_to_csv():
    with open('employee_data.csv', mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([entry_id.get(), entry_name.get(), entry_department.get(), entry_position.get(),
                         entry_birthdate.get(), entry_gender.get(), entry_id_card.get(), entry_date_of_issue.get(), entry_issued_by.get()])
    messagebox.showinfo("Thông báo", "Thông tin nhân viên đã được lưu!")

# Hàm hiển thị nhân viên có sinh nhật hôm nay
def birthday_today():
    today = datetime.today().strftime('%d/%m')
    today_employees = []

    # Đọc từ file CSV và kiểm tra sinh nhật
    try:
        with open('employee_data.csv', mode='r') as file:
            reader = csv.reader(file)
            for row in reader:
                birthdate = row[4]  # Cột ngày sinh
                if birthdate[-5:] == today:  # Kiểm tra sinh nhật
                    today_employees.append(row)
        if today_employees:
            result = "\n".join([f"ID: {emp[0]}, Tên: {emp[1]}, Sinh nhật: {emp[4]}" for emp in today_employees])
            messagebox.showinfo("Sinh nhật hôm nay", result)
        else:
            messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên sinh nhật hôm nay.")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên.")

# Hàm xuất toàn bộ danh sách nhân viên ra file Excel, sắp xếp theo tuổi giảm dần
def export_to_excel():
    try:
        df = pd.read_csv('employee_data.csv', names=["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "CMND", "Ngày cấp", "Nơi cấp"])
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y')
        df['Tuổi'] = (datetime.now() - df['Ngày sinh']).dt.days // 365
        df_sorted = df.sort_values(by='Tuổi', ascending=False)  # Sắp xếp theo tuổi giảm dần
        df_sorted.to_excel('employee_data_sorted.xlsx', index=False, engine='openpyxl')
        messagebox.showinfo("Thông báo", "Đã xuất toàn bộ danh sách nhân viên ra file Excel.")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên.")

# Tạo giao diện Tkinter
root = tk.Tk()
root.title("Quản lý Nhân viên")

# Các nhãn và trường nhập liệu
tk.Label(root, text="Mã nhân viên").grid(row=0, column=0)
entry_id = tk.Entry(root)
entry_id.grid(row=0, column=1)

tk.Label(root, text="Tên nhân viên").grid(row=1, column=0)
entry_name = tk.Entry(root)
entry_name.grid(row=1, column=1)

tk.Label(root, text="Đơn vị").grid(row=2, column=0)
entry_department = tk.Entry(root)
entry_department.grid(row=2, column=1)

tk.Label(root, text="Chức danh").grid(row=3, column=0)
entry_position = tk.Entry(root)
entry_position.grid(row=3, column=1)

tk.Label(root, text="Ngày sinh (dd/mm/yyyy)").grid(row=4, column=0)
entry_birthdate = tk.Entry(root)
entry_birthdate.grid(row=4, column=1)

tk.Label(root, text="Giới tính").grid(row=5, column=0)
entry_gender = tk.Entry(root)
entry_gender.grid(row=5, column=1)

tk.Label(root, text="Số CMND").grid(row=6, column=0)
entry_id_card = tk.Entry(root)
entry_id_card.grid(row=6, column=1)

tk.Label(root, text="Ngày cấp CMND (dd/mm/yyyy)").grid(row=7, column=0)
entry_date_of_issue = tk.Entry(root)
entry_date_of_issue.grid(row=7, column=1)

tk.Label(root, text="Nơi cấp CMND").grid(row=8, column=0)
entry_issued_by = tk.Entry(root)
entry_issued_by.grid(row=8, column=1)

# Nút lưu thông tin nhân viên
btn_save = tk.Button(root, text="Lưu thông tin nhân viên", command=save_to_csv)
btn_save.grid(row=9, column=0, columnspan=2)

# Nút sinh nhật hôm nay
btn_birthday_today = tk.Button(root, text="Sinh nhật ngày hôm nay", command=birthday_today)
btn_birthday_today.grid(row=10, column=0, columnspan=2)

# Nút xuất toàn bộ danh sách
btn_export_excel = tk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel)
btn_export_excel.grid(row=11, column=0, columnspan=2)

# Chạy ứng dụng Tkinter
root.mainloop()
