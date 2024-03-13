import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import pandas as pd
import threading

def get_sheet_names(workbook_path, loading_window):
    try:
        loading_window.deiconify()  # Hiển thị cửa sổ loading
        loading_label.config(text="Loading...")
        loading_bar.start()  # Khởi động thanh tiến trình
        loading_label.update()
        workbook = openpyxl.load_workbook(workbook_path.replace('/', '\\'))
        sheet_names = workbook.sheetnames
        loading_window.withdraw()  # Ẩn cửa sổ loading khi đã xong
        root.deiconify()  # Hiển thị lại cửa sổ giao diện sau khi loading xong
        return sheet_names
    except Exception as e:
        messagebox.showerror("Error", str(e))
        loading_window.withdraw()  # Ẩn cửa sổ loading khi gặp lỗi
        root.deiconify()  # Hiển thị lại cửa sổ giao diện khi gặp lỗi
        return []

def browse_file():
    def browse_file_thread():
        filename = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")))
        if filename:
            entry_workbook_path.delete(0, tk.END)
            entry_workbook_path.insert(tk.END, filename)
            sheet_names = get_sheet_names(filename, loading_window)
            combobox_sheet_name.delete(0, tk.END)
            for sheet_name in sheet_names:
                combobox_sheet_name.insert(tk.END, sheet_name)
            check_run_button_state()  # Kiểm tra lại trạng thái của nút "Run Script" sau khi có dữ liệu

    threading.Thread(target=browse_file_thread).start()

def select_all():
    combobox_sheet_name.select_set(0, tk.END)
    check_run_button_state()

def deselect_all():
    combobox_sheet_name.selection_clear(0, tk.END)
    check_run_button_state()

def check_run_button_state():
    if combobox_sheet_name.curselection():
        run_button.config(state="normal")  # Bật nút nếu có sheet được chọn
    else:
        run_button.config(state="disabled")  # Ẩn nút nếu không có sheet được chọn

def run_script():
    def process_data():
        nonlocal workbook_path, sheet_names_selected
        if workbook_path and sheet_names_selected:
            dfs = get_data(workbook_path, sheet_names_selected, loading_window)
            if dfs:
                for sheet_name, df in dfs.items():
                    save_path = f"{workbook_path.split('.')[0]}_{sheet_name}.csv"
                    df.to_csv(save_path, index=False)
                print("Done!")
                messagebox.showinfo("Success", "Data extracted and saved successfully!")
        else:
            messagebox.showwarning("Warning", "Please select workbook and sheet.")

    root.withdraw()  # Ẩn cửa sổ giao diện khi bắt đầu loading
    workbook_path = entry_workbook_path.get()
    sheet_names = combobox_sheet_name.curselection()  # Lấy chỉ số của các sheet được chọn
    if workbook_path and sheet_names:
        sheet_names_selected = [combobox_sheet_name.get(idx) for idx in sheet_names]
        loading_thread = threading.Thread(target=process_data)
        loading_thread.start()
    else:
        messagebox.showwarning("Warning", "Please select workbook and sheet.")

def get_data(workbook_path, sheet_names, loading_window):
    try:
        loading_window.deiconify()  # Hiển thị cửa sổ loading
        loading_label.config(text="Loading...")
        loading_bar.start()  # Khởi động thanh tiến trình
        loading_label.update()
        dfs = {}
        for sheet_name in sheet_names:
            dfs[sheet_name] = pd.read_excel(workbook_path, sheet_name=sheet_name, engine='openpyxl')
            print(f"Data retrieved from {workbook_path} (Sheet: {sheet_name})")
        loading_window.withdraw()  # Ẩn cửa sổ loading khi đã xong
        root.deiconify()  # Hiển thị lại cửa sổ giao diện sau khi loading xong
        return dfs
    except Exception as e:
        messagebox.showerror("Error", str(e))
        loading_window.withdraw()  # Ẩn cửa sổ loading khi gặp lỗi
        root.deiconify()  # Hiển thị lại cửa sổ giao diện khi gặp lỗi
        return {}

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

# Tạo cửa sổ giao diện
root = tk.Tk()
root.title("Excel Data Extractor")

# Label và Entry cho đường dẫn workbook
label_workbook_path = tk.Label(root, text="Workbook Path:")
label_workbook_path.grid(row=0, column=0, padx=5, pady=5)
entry_workbook_path = tk.Entry(root, width=50)
entry_workbook_path.grid(row=0, column=1, padx=5, pady=5)
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

# Label và Listbox cho tên sheet
label_sheet_name = tk.Label(root, text="Sheet Name:")
label_sheet_name.grid(row=1, column=0, padx=5, pady=5)
combobox_sheet_name = tk.Listbox(root, selectmode="multiple", width=48, height=5)
combobox_sheet_name.grid(row=1, column=1, padx=5, pady=5)
combobox_sheet_name.bind("<<ListboxSelect>>", lambda event: check_run_button_state())  # Gắn sự kiện khi chọn sheet

# Nút chọn tất cả
select_all_button = tk.Button(root, text="Select All", command=select_all)
select_all_button.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

# Nút bỏ chọn tất cả
deselect_all_button = tk.Button(root, text="Deselect All", command=deselect_all)
deselect_all_button.grid(row=1, column=3, padx=5, pady=5, sticky="ew")

# Cửa sổ loading
loading_window = tk.Toplevel(root)
loading_window.withdraw()  # Ẩn cửa sổ loading khi khởi tạo
loading_label = tk.Label(loading_window, text="Loading...", font=('Helvetica', 12))
loading_label.pack(pady=10)
loading_bar = ttk.Progressbar(loading_window, mode='indeterminate')
loading_bar.pack(pady=5)

# Nút chạy mã
run_button = tk.Button(root, text="Run Script", command=run_script, state="disabled")
run_button.grid(row=2, column=1, padx=5, pady=5)

# Đặt vị trí căn giữa cho cửa sổ giao diện
center_window(root)
# Đặt vị trí căn giữa cho cửa sổ loading
center_window(loading_window)

# Chạy vòng lặp chính của ứng dụng
root.mainloop()