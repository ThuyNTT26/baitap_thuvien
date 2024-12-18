import tkinter as tk
from tkinter import messagebox, filedialog
import csv
from datetime import datetime
import pandas as pd
import os

#Hàm lưu thông tin vào file csv
def luu_thong_tin():
    # Lấy giá trị từ các trường
    id_value = id_entry.get().strip()
    name_value = name_entry.get().strip()
    unit_value = unit_entry.get().strip()

    #Kiểm tra các trường bắt buộc
    if not id_value or not name_value or not unit_value:
        messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ các mục: Mã, Tên và Đơn vị!")
    # #Kiểm tra nếu không tích chọn checkbox thì thông báo lỗi
    if not khachhang.get() and not nhacungcap.get():
        messagebox.showerror("Lỗi", "Vui lòng chọn 'Là khách hàng' hoặc 'Là nhà cung cấp'!")

    if not os.path.exists("employees.csv"):
        messagebox.showerror("Lỗi", "Không tìm thấy tệp dữ liệu!")
        return
    data={
        "Mã": id_entry.get(),
        "Tên": name_entry.get(),
        "Đơn vị": unit_entry.get(),
        "Chức danh": title_entry.get(),
        "Giới tính": gender_var.get(),
        "Ngày sinh": dob_entry.get().strip(),
        "Số CMND": id_card_entry.get().strip(),
        "Nơi cấp": issue_place_entry.get().strip(),
        "Ngày cấp": issue_date_entry.get().strip(),
        "Là khách hàng": "Có" if khachhang.get() else "Không",
        "Là nhà cung cấp": "Có" if nhacungcap.get() else "Không",
    }

    f= open("employees.csv", "a", newline="", encoding="utf-8")
    writer = csv.DictWriter(f, fieldnames=data.keys())
    if f.tell() == 0:
        writer.writeheader()
    writer.writerow(data)
    messagebox.showinfo("Thành công", "Lưu thông tin thành công!")
    xoa_thong_tin()

#Hàm xóa thông tin đã nhập
def xoa_thong_tin():
    id_entry.delete(0, tk.END)
    name_entry.delete(0, tk.END)
    unit_entry.delete(0, tk.END)
    title_entry.delete(0, tk.END)
    gender_var.set("")
    dob_entry.delete(0, tk.END)
    id_card_entry.delete(0, tk.END)
    issue_place_entry.delete(0, tk.END)
    issue_date_entry.delete(0, tk.END)


#Hàm hiển thị nhân viên sinh hôm nay
def sinh_nhat():
    today = datetime.now().strftime('%d.%m.%Y')
    try:
        f = open("employees.csv", "r", newline="",encoding="utf-8")
        reader = csv.DictReader(f)
        results = [row for row in reader if row.get("Ngày sinh")==today]
    except FileNotFoundError:
        results = []
    if results:
        result_text = "\n".join([f"{r['Mã']} - {r['Tên']}" for r in results])
        messagebox.showinfo("Kết quả", f"Nhân viên sinh nhật hôm nay:\n{result_text}")


#Hàm xuất danh sách ra file Excel
def xuat_ds():
    try:
        data = pd.read_csv("employees.csv", encoding="utf-8")
        data["Ngày sinh"] = pd.to_datetime(data["Ngày sinh"], format="%d.%m.%Y", errors="coerce")
        sorted_data = data.sort_values(by="Ngày sinh", ascending=False)
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            sorted_data.to_excel(file_path, index=False)
            messagebox.showinfo("Thành công", "Xuất danh sách thành công")

    except Exception as e:
        messagebox.showerror("Lỗi", f"Định dạng ngày tháng không hợp lệ: {e}")

#Tạo giao diện
root = tk.Tk()
root.title("Thông tin nhân viên")

tk.Label(root, text="Mã:").grid(row=0, column=0, padx=5, pady=5)
id_entry = tk.Entry(root)
id_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Tên:").grid(row=1, column=0, padx=5, pady=5)
name_entry = tk.Entry(root)
name_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Đơn vị:").grid(row=2, column=0, padx=5, pady=5)
unit_entry = tk.Entry(root)
unit_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Chức danh").grid(row=3, column=0, padx=5, pady=5)
title_entry = tk.Entry(root)
title_entry.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Giới tính").grid(row=4, column=0, padx=5, pady=5)
gender_var = tk.StringVar()
tk.Radiobutton(root,text="Nam", variable=gender_var, value="Nam").grid(row=4, column=1)
tk.Radiobutton(root,text="Nữ", variable=gender_var, value="Nữ").grid(row=4, column=2)

tk.Label(root, text="Ngày sinh (DD/MM/YYYY):").grid(row=5, column=0, padx=5, pady=5)
dob_entry = tk.Entry(root)
dob_entry.grid(row=5, column=1, padx=5, pady=5)

tk.Label(root, text="Số CMND").grid(row=6, column=0, padx=5, pady=5)
id_card_entry = tk.Entry(root)
id_card_entry.grid(row=6, column=1, padx=5, pady=5)

tk.Label(root,text="Nơi cấp").grid(row=7, column=0, padx=5, pady=5)
issue_place_entry = tk.Entry(root)
issue_place_entry.grid(row=7, column=1, padx=5, pady=5)

tk.Label(root, text="Ngày cấp (DD/MM/YYYY):").grid(row=8, column=0, padx=5, pady=5)
issue_date_entry = tk.Entry(root)
issue_date_entry.grid(row=8, column=1, padx=5, pady=5)

khachhang = tk.IntVar()
nhacungcap = tk.IntVar()
tk.Checkbutton(root,text="Là khách hàng", variable=khachhang).grid(row=9, column=0, padx=5, pady=5)
tk.Checkbutton(root, text="Là nhà cung cấp", variable=nhacungcap).grid(row=9, column=1, padx=5, pady=5)

tk.Button(root, text="Lưu thông tin", command=luu_thong_tin).grid(row=10, column=0, padx=5, pady=5)
tk.Button(root,text="Sinh nhật hôm nay", command=sinh_nhat).grid(row=10, column=1, padx=5, pady=5)
tk.Button(root, text="Xuất danh sách", command=xuat_ds).grid(row=10, column=2, padx=5, pady=5)

root.mainloop()






