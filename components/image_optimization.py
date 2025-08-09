import os
import json
import tkinter as tk
from tkinter import filedialog
import sys
from src.image_process.image_process import (
    convert_images_to_format,
    process_text_in_images,
    rename_images,
    add_logo_to_images,
    add_frame_to_images
)

# ---------------------------
# Cấu hình mặc định và lưu vào config.json
# ---------------------------
ROOT_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
CONFIG_PATH = os.path.join(ROOT_PATH, "config.json")
CONFIG_SECTION = "image_optimization"

DEFAULT_CONFIG = {
    "folder_path": r"C:\Users\Hieu Pham\Downloads\anh seo",
    "logo_path": r"C:\Users\Hieu Pham\Downloads\anh seo\logo\logo-j88.png",
    "box_style": "linear_gradient",
    "format_logo": "jpg",
    "format_no_logo": "jpg",
    "length": 800,
    "width": 400,
    "font_path": r"font\Bungee-Regular.ttf"
}

def load_config():
    """
    Load cấu hình từ file config.json. Nếu chưa có phần CONFIG_SECTION
    thì khởi tạo với DEFAULT_CONFIG và lưu vào file.
    """
    config_data = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                config_data = json.load(f)
        except Exception as e:
            print(f"[ERROR] Failed to load config: {e}")
    if CONFIG_SECTION not in config_data:
        config_data[CONFIG_SECTION] = DEFAULT_CONFIG.copy()
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            print(f"[ERROR] Failed to save default config: {e}")
    else:
        # Cập nhật các key còn thiếu nếu có
        for key, value in DEFAULT_CONFIG.items():
            config_data[CONFIG_SECTION].setdefault(key, value)
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            print(f"[ERROR] Failed to update config: {e}")
    return config_data.get(CONFIG_SECTION, {})

def save_config(new_config):
    """
    Lưu cấu hình mới vào file config.json (chỉ cập nhật phần CONFIG_SECTION).
    """
    config_data = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                config_data = json.load(f)
        except Exception as e:
            print(f"[ERROR] Failed to load config for saving: {e}")
    config_data[CONFIG_SECTION] = new_config
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)
    except Exception as e:
        print(f"[ERROR] Failed to save config: {e}")

# Lấy cấu hình mặc định từ config.json (hoặc DEFAULT_CONFIG nếu chưa có)
config = load_config()

# ---------------------------
# Phần giao diện Tkinter và xử lý ảnh
# ---------------------------
class RedirectText(object):
    def __init__(self, widget):
        self.widget = widget
    def write(self, string):
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)
    def flush(self):
        pass

class ImageOptimizationFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Frame chứa các input (căn trái)
        inputs_frame = tk.Frame(self)
        inputs_frame.pack(anchor="w", padx=10, pady=10)
        
        # Folder path với Browse
        tk.Label(inputs_frame, text="Folder Path:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.folder_path_entry = tk.Entry(inputs_frame, width=50)
        self.folder_path_entry.insert(0, config.get("folder_path", ""))
        self.folder_path_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        browse_folder_btn = tk.Button(inputs_frame, text="Browse", command=self.browse_folder)
        browse_folder_btn.grid(row=0, column=2, sticky="w", padx=5, pady=5)
        
        # Logo path (hoặc Frame path) với Browse
        tk.Label(inputs_frame, text="Logo/Frame Path:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.logo_path_entry = tk.Entry(inputs_frame, width=50)
        self.logo_path_entry.insert(0, config.get("logo_path", ""))
        self.logo_path_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        browse_logo_btn = tk.Button(inputs_frame, text="Browse", command=self.browse_logo)
        browse_logo_btn.grid(row=1, column=2, sticky="w", padx=5, pady=5)
        
        # Box style (OptionMenu)
        tk.Label(inputs_frame, text="Box Style:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.box_style_var = tk.StringVar(value=config.get("box_style", "linear_gradient"))
        box_style_options = ["linear_gradient", "gradient", "None"]
        self.box_style_menu = tk.OptionMenu(inputs_frame, self.box_style_var, *box_style_options)
        self.box_style_menu.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # Format logo/frame (OptionMenu) - dùng chung
        tk.Label(inputs_frame, text="Format Logo/Frame:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.format_logo_var = tk.StringVar(value=config.get("format_logo", "jpg"))
        format_options = ["jpg", "jpeg", "webp", "png"]
        self.format_logo_menu = tk.OptionMenu(inputs_frame, self.format_logo_var, *format_options)
        self.format_logo_menu.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        # Format no logo (OptionMenu)
        tk.Label(inputs_frame, text="Format No Logo:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.format_no_logo_var = tk.StringVar(value=config.get("format_no_logo", "jpg"))
        self.format_no_logo_menu = tk.OptionMenu(inputs_frame, self.format_no_logo_var, *format_options)
        self.format_no_logo_menu.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        
        # Length
        tk.Label(inputs_frame, text="Length:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.length_entry = tk.Entry(inputs_frame, width=10)
        self.length_entry.insert(0, str(config.get("length", 800)))
        self.length_entry.grid(row=5, column=1, sticky="w", padx=5, pady=5)
        
        # Width
        tk.Label(inputs_frame, text="Width:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        self.width_entry = tk.Entry(inputs_frame, width=10)
        self.width_entry.insert(0, str(config.get("width", 400)))
        self.width_entry.grid(row=6, column=1, sticky="w", padx=5, pady=5)
        
        # Font path với Browse
        tk.Label(inputs_frame, text="Font Path:").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        self.font_path_entry = tk.Entry(inputs_frame, width=50)
        self.font_path_entry.insert(0, config.get("font_path", ""))
        self.font_path_entry.grid(row=7, column=1, sticky="w", padx=5, pady=5)
        browse_font_btn = tk.Button(inputs_frame, text="Browse", command=self.browse_font)
        browse_font_btn.grid(row=7, column=2, sticky="w", padx=5, pady=5)
        
        # Frame chứa checkbutton (căn trái)
        checks_frame = tk.Frame(self)
        checks_frame.pack(anchor="w", padx=10, pady=10)
        
        self.do_convert_var = tk.BooleanVar(value=True)
        self.do_process_text_var = tk.BooleanVar(value=True)
        self.do_rename_var = tk.BooleanVar(value=True)
        
        # Thêm 2 checkbutton "Add Logo" và "Add Frame" chỉ được chọn 1
        self.do_add_logo_var = tk.BooleanVar(value=True)
        self.do_add_frame_var = tk.BooleanVar(value=False)
        
        self.check_convert = tk.Checkbutton(checks_frame, text="Convert Images Format", variable=self.do_convert_var)
        self.check_convert.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        self.check_process_text = tk.Checkbutton(checks_frame, text="Process Text in Images", variable=self.do_process_text_var)
        self.check_process_text.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        self.check_rename = tk.Checkbutton(checks_frame, text="Rename Images", variable=self.do_rename_var)
        self.check_rename.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        # Checkbutton Add Logo
        self.check_add_logo = tk.Checkbutton(
            checks_frame,
            text="Add Logo to Images",
            variable=self.do_add_logo_var,
            command=self.on_add_logo_check  # Hàm callback
        )
        self.check_add_logo.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        # Checkbutton Add Frame
        self.check_add_frame = tk.Checkbutton(
            checks_frame,
            text="Add Frame to Images",
            variable=self.do_add_frame_var,
            command=self.on_add_frame_check  # Hàm callback
        )
        self.check_add_frame.grid(row=1, column=2, sticky="w", padx=5, pady=5)
        
        # Nút Run Pipeline, căn trái
        run_button = tk.Button(self, text="Run", command=self.run_pipeline)
        run_button.pack(anchor="w", padx=10, pady=10)
        
        # Nếu muốn sử dụng ô Logs:
        # log_frame = tk.Frame(self)
        # log_frame.pack(fill="both", padx=10, pady=10, expand=True)
        # self.log_text = tk.Text(log_frame, height=10, wrap="word")
        # self.log_text.pack(side="left", fill="both", expand=True)
        # scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        # scrollbar.pack(side="right", fill="y")
        # self.log_text.config(yscrollcommand=scrollbar.set)
        # sys.stdout = RedirectText(self.log_text)
    
    def on_add_logo_check(self):
        """
        Nếu chọn Add Logo, bỏ chọn Add Frame.
        """
        if self.do_add_logo_var.get():
            self.do_add_frame_var.set(False)
    
    def on_add_frame_check(self):
        """
        Nếu chọn Add Frame, bỏ chọn Add Logo.
        """
        if self.do_add_frame_var.get():
            self.do_add_logo_var.set(False)
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path_entry.delete(0, tk.END)
            self.folder_path_entry.insert(0, folder_selected)
    
    def browse_logo(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.webp")])
        if file_selected:
            self.logo_path_entry.delete(0, tk.END)
            self.logo_path_entry.insert(0, file_selected)
    
    def browse_font(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Font Files", "*.ttf;*.otf")])
        if file_selected:
            self.font_path_entry.delete(0, tk.END)
            self.font_path_entry.insert(0, file_selected)
    
    def run_pipeline(self):
        # Lấy các giá trị từ input
        folder_path = self.folder_path_entry.get()
        logo_path = self.logo_path_entry.get()
        box_style = self.box_style_var.get()
        format_logo = self.format_logo_var.get()
        format_no_logo = self.format_no_logo_var.get()
        try:
            length = int(self.length_entry.get())
            width = int(self.width_entry.get())
        except ValueError:
            print("Length và Width phải là số nguyên.\n")
            return
        font_path = self.font_path_entry.get()
        
        # Cập nhật lại config với các giá trị hiện tại
        new_config = {
            "folder_path": folder_path,
            "logo_path": logo_path,
            "box_style": box_style,
            "format_logo": format_logo,
            "format_no_logo": format_no_logo,
            "length": length,
            "width": width,
            "font_path": font_path
        }
        save_config(new_config)
        
        # Lấy trạng thái các checkbutton
        do_convert = self.do_convert_var.get()
        do_process_text = self.do_process_text_var.get()
        do_rename = self.do_rename_var.get()
        do_add_logo = self.do_add_logo_var.get()
        do_add_frame = self.do_add_frame_var.get()
        
        process_images_pipeline(
            folder_path, format_no_logo, font_path, box_style, logo_path,
            format_logo, length, width,
            do_convert, do_process_text, do_rename, do_add_logo, do_add_frame
        )
        print("Pipeline executed.\n")

def process_images_pipeline(folder_path, format_no_logo, font_path, box_style, logo_path, format_logo, length, width,
                            do_convert, do_process_text, do_rename, do_add_logo, do_add_frame):
    """
    Thực hiện xử lý ảnh theo các bước:
    1. Chuyển đổi định dạng và nén ảnh về kích thước (length x width)
    2. Thêm text lên ảnh nếu tên file có chứa khoảng trắng (với tùy chọn box_style)
    3. Đổi tên file theo chuẩn
    4. Thêm logo hoặc frame vào ảnh (chỉ 1 trong 2, tùy người dùng chọn)
    """
    if do_convert:
        convert_images_to_format(folder_path, format_no_logo, length, width)
    if do_add_frame:
        add_frame_to_images(folder_path, logo_path, format_logo, length, width)
    if do_process_text:
        process_text_in_images(folder_path, font_path, box_style=box_style)
    if do_rename:
        rename_images(folder_path)
    
    # Chỉ thêm Logo hoặc Frame (ưu tiên Logo nếu do_add_logo=True)
    if do_add_logo:
        add_logo_to_images(folder_path, logo_path, format_logo, length, width)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Image Optimization Tool")
    frame = ImageOptimizationFrame(root)
    frame.pack(fill="both", expand=True)
    root.mainloop()