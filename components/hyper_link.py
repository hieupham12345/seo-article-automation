import tkinter as tk
from tkinter import filedialog
import threading
import sys
import os

# Lấy thư mục gốc của project
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Thêm thư mục src vào sys.path
sys.path.append(os.path.join(BASE_DIR, "../"))

from  src.add_hyperlink import add_hyper_link_folder


class InsertHyperlinkComponent(tk.Frame):
    def __init__(self, master=None, default_folder=r"C:\Users\Hieu Pham\Downloads\SEO\full_seos\normalize_seos"):
        super().__init__(master)
        self.master = master
        self.default_folder = default_folder
        self.create_widgets()

    def create_widgets(self):
        # Frame cho chọn folder
        folder_frame = tk.Frame(self)
        folder_frame.pack(padx=10, pady=10, fill='x')

        folder_label = tk.Label(folder_frame, text="Folder:")
        folder_label.pack(side="left")

        # Giới hạn chiều ngang khoảng 100px (ước tính ~12 ký tự)
        self.folder_entry = tk.Entry(folder_frame, width=12)
        self.folder_entry.pack(side="left", fill='x', expand=True)
        self.folder_entry.insert(0, self.default_folder)  # Gán giá trị mặc định

        browse_button = tk.Button(folder_frame, text="Browse", command=self.browse_folder)
        browse_button.pack(side="left", padx=(5, 0))

        # Frame cho ô nhập text (nhiều dòng)
        text_frame = tk.Frame(self)
        text_frame.pack(padx=10, pady=10, fill='both', expand=True)

        text_label = tk.Label(text_frame, text="Text (File name, Anchor, Link, Paragrahp position):")
        text_label.pack(anchor="nw")

        # Giới hạn chiều ngang khoảng 100px (ước tính ~12 ký tự)
        self.text_box = tk.Text(text_frame, height=10, width=12)
        self.text_box.pack(fill='both', expand=True)

        # Nút 'Run'
        run_button = tk.Button(self, text="Run", command=self.run_function)
        run_button.pack(pady=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)
            

    def run_function(self):
        thread = threading.Thread(target=self.run_add_hyperlink)
        thread.start()

    def run_add_hyperlink(self):
        folder_path = self.folder_entry.get()
        text_data = self.text_box.get("1.0", tk.END).strip()

        print("Bắt đầu chạy add_hyper_link_folder...")
        add_hyper_link_folder(folder_path, text_data)
        print("Chạy xong!")

# Test nếu chạy trực tiếp file này
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Insert hyperlink (anchor text)")

    app = InsertHyperlinkComponent(root)
    app.pack(fill='both', expand=True)

    root.mainloop()
