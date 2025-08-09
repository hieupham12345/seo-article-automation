# components/seo_write.py
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import os
import sys
import json
from tkinter import ttk

from src.format.format_timeNewRoman import process_folder_format_time
from src.format.format_arial import process_folder_format_arial

from src.controller import (
    generate_outline_func,
    write_full_func,
    normalize_seo,
    fix_spin_folder,
)

# ---------------------------
# Cấu hình file config
# ---------------------------
# Xác định ROOT_PATH (thư mục gốc của dự án), giả sử file này nằm trong folder con của root folder.
ROOT_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
CONFIG_PATH = os.path.join(ROOT_PATH, "config.json")
CONFIG_SECTION = "seo_write"

# Các giá trị mặc định dùng để khởi tạo phần cấu hình SEO Write trong config.json
DEFAULT_CONFIG = {
    "number_of_h2": 3,
    "number_of_h3": 3,
    "number_of_image": 3,
    "max_key": 17,
    "language": "Vietnamese",
    "seo_category": "Formal",
    "input_folder": r"C:\Users\Hieu Pham\Downloads\SEO",
    "model_write": "Gemini 1.5 pro [3000]",
    "model_outline": "Gemini 1.5 pro [1000]",
    "model_normalize": "Chatgpt 4o mini [2000]",
    "learn_data": "Không dùng từ như anh em, cược thủ, bet thủ ...",
    "outline_data": "",
    "brand_keyword": "",
    "keywords": "",
    "word_density": 2.8,
}

def load_console_config():
    """
    Load phần cấu hình SEO Write từ file config.json nằm ở ROOT_PATH.
    Nếu chưa tồn tại phần 'seo_write' thì khởi tạo với DEFAULT_CONFIG và lưu vào file.
    Nếu tồn tại, cập nhật các key bị thiếu với giá trị mặc định.
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

def save_console_config(new_section_data):
    """
    Cập nhật phần cấu hình SEO Write (key 'seo_write') vào config.json.
    Chỉ cập nhật phần này mà không ghi đè các phần cấu hình khác.
    """
    config_data = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                config_data = json.load(f)
        except Exception as e:
            print(f"[ERROR] Failed to load config for saving: {e}")
    config_data[CONFIG_SECTION] = new_section_data
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)
    except Exception as e:
        print(f"[ERROR] Failed to save config: {e}")

# Lấy cấu hình dành cho SEO Write từ file config.json (hoặc từ DEFAULT_CONFIG nếu chưa có)
console_config = load_console_config()

# ---------------------------
# Các giá trị mặc định và mapping model
# ---------------------------
DEFAULT_MAX_WORD = 3000
DEFAULT_NORMALIZE_MODEL_TYPE = "chatgpt"
DEFAULT_NORMALIZE_MODEL_NAME = "gpt-4o-mini"

DEFAULT_TEMPERATURE = 1.1
DEFAULT_SEARCH_TYPE = "er"
DEFAULT_IS_SPIN_EDITOR = True

MODEL_OPTIONS_OUTLINE = {

    "Gemini 1.5 pro [1000]": ("gemini", "gemini-1.5-pro"),
    "Gemini 2.0 flash [1000]": ("gemini", "gemini-2.0-flash"),
    "Chatgpt 4.1 mini [1000]": ("chatgpt", "gpt-4.1-mini")
}


MODEL_OPTIONS_WRITE = {
    "Gemini 2.5 pro [3000]": ("gemini", "gemini-2.5-pro-preview-03-25"),

    "Gemini 1.5 pro [3000]": ("gemini", "gemini-1.5-pro"),
    "Chatgpt 4.5 [3000]": ("chatgpt", "gpt-4.5-preview"),
    "Chatgpt 4.1 [3000]": ("chatgpt", "gpt-4.1"),
    "Gemini 2.5 flash [1000]": ("gemini", "gemini-2.5-flash-preview-04-17"),

    "Gemini 2.0 flash [1000]": ("gemini", "gemini-2.0-flash"),
    "Chatgpt 4.1 mini [1000]": ("chatgpt", "gpt-4.1-mini")
}
# claude-3-7-sonnet-20250219


MODEL_OPTIONS_NORMALIZE = {
    "Chatgpt 4o mini [2000]": ("chatgpt", "gpt-4o-mini"),
    "Gemini 2.0 flash [2000]": ("gemini", "gemini-2.0-flash"),
}
#
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
            

# ---------------------------
# Hỗ trợ redirect print -> Text widget
# ---------------------------
class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, msg):
        self.text_widget.after(0, self._append_text, msg)

    def flush(self):
        pass

    def _append_text(self, msg):
        self.text_widget.insert(tk.END, msg)
        self.text_widget.see(tk.END)

# ---------------------------
# Lớp SEO Write Interface
# ---------------------------
class SEOWriteFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()
        self.load_settings()  # Load các thiết lập từ console_config

    def create_widgets(self):
        # Tạo bố cục chia làm 2 phần: trái và phải
        left_frame = tk.Frame(self)
        left_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        right_frame = tk.Frame(self)
        right_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        # --- General Settings ---
        general_frame = tk.LabelFrame(left_frame, text="General Settings", padx=5, pady=5)
        general_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        tk.Label(general_frame, text="Number of H2:", anchor="w").grid(row=0, column=0, sticky="w", padx=0, pady=5)
        self.entry_number_of_h2 = ttk.Combobox(general_frame, values=[2, 3, 4, 5], state="readonly", width=5)
        default_h2 = console_config.get("number_of_h2", 3)  # Nếu không có key, dùng 3 làm mặc định
        self.entry_number_of_h2.set(default_h2)
        self.entry_number_of_h2.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        tk.Label(general_frame, text="Number of H3:", anchor="w").grid(row=1, column=0, sticky="w", padx=0, pady=5)
        self.entry_number_of_h3 = ttk.Combobox(general_frame, values=[2, 3, 4, 5], state="readonly", width=5)
        default_h3 = console_config.get("number_of_h3", 3)  # Nếu không có key, dùng 3 làm mặc định
        self.entry_number_of_h3.set(default_h3)
        self.entry_number_of_h3.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Label(general_frame, text="Number of Images:").grid(row=2, column=0, sticky="w")
        self.entry_number_of_image = ttk.Combobox(general_frame, values=[3, 4, 5, 6, 7, 8, 9], state="readonly", width=5)
        default_h3 = console_config.get("number_of_image", 3)  # Nếu không có key, dùng 3 làm mặc định
        self.entry_number_of_image.set(default_h3)  # Giá trị mặc định
        self.entry_number_of_image.grid(row=2, column=1, padx=5, pady=5, sticky="w")


        tk.Label(general_frame, text="Max Key:").grid(row=3, column=0, sticky="w")
        self.entry_max_key = ttk.Combobox(general_frame, values=[13, 17, 22, 30, 40, 50, 60], state="readonly", width=5)
        self.entry_max_key.set(17)  # Giá trị mặc định
        self.entry_max_key.grid(row=3, column=1, padx=5, pady=5, sticky="w")


        tk.Label(general_frame, text="Language:").grid(row=4, column=0, sticky="w")
        languages = [
                        "English", "Vietnamese", "Japanese", "Chinese", "Korean", "French",
                        "German", "Spanish", "Portuguese", "Russian", "Italian", "Dutch",
                        "Arabic", "Turkish", "Thai", "Hindi", "Bengali", "Indonesian",
                        "Malay", "Polish"
                    ]

        default_language = console_config.get("language", languages[0])  # Nếu không có key, mặc định là languages[0]
        self.language_var = tk.StringVar(value=default_language)
        language_menu = tk.OptionMenu(general_frame, self.language_var, *languages)
        language_menu.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        # Thay đổi SEO Category thành OptionMenu với 2 lựa chọn
        tk.Label(general_frame, text="SEO style:").grid(row=5, column=0, sticky="w")
        seo_categories = ["Formal", "Formal-Creative","Storytelling"]
        self.seo_category_var = tk.StringVar(value=seo_categories[0])
        seo_category_menu = tk.OptionMenu(general_frame, self.seo_category_var, *seo_categories)
        seo_category_menu.grid(row=5, column=1, padx=5, pady=5, sticky="w")


 

        

        # --- Model Settings ---
        model_frame = tk.LabelFrame(left_frame, text="Model Settings", padx=5, pady=5)
        model_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        tk.Label(model_frame, text="Outline Model:").grid(row=0, column=0, sticky="w")
        self.outline_model_var = tk.StringVar(value=console_config.get("model_outline"))
        outline_model_menu = tk.OptionMenu(model_frame, self.outline_model_var, *MODEL_OPTIONS_OUTLINE.keys())
        outline_model_menu.grid(row=0, column=1, sticky="ew")

        tk.Label(model_frame, text="Write Model:").grid(row=1, column=0, sticky="w")
        self.write_model_var = tk.StringVar(value=console_config.get("model_write"))
        write_model_menu = tk.OptionMenu(model_frame, self.write_model_var, *MODEL_OPTIONS_WRITE.keys())
        write_model_menu.grid(row=1, column=1, sticky="ew")

        tk.Label(model_frame, text="Normalize Model:").grid(row=2, column=0, sticky="w")
        self.normalize_model_var = tk.StringVar(value=console_config.get("model_normalize"))
        normalize_model_menu = tk.OptionMenu(model_frame, self.normalize_model_var, *MODEL_OPTIONS_NORMALIZE.keys())
        normalize_model_menu.grid(row=2, column=1, sticky="ew")

        # --- Folder Settings ---
        folder_frame = tk.LabelFrame(left_frame, text="Folder Settings", padx=5, pady=5)
        folder_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

        tk.Label(folder_frame, text="Input Folder:").grid(row=0, column=0, sticky="w")
        self.input_folder_entry = tk.Entry(folder_frame, width=50)
        self.input_folder_entry.grid(row=0, column=1)
        tk.Button(folder_frame, text="Browse", command=lambda: self.browse_folder(self.input_folder_entry)).grid(row=0, column=2)

        # --- Learn Data Input ---
        learn_frame = tk.LabelFrame(left_frame, text="Learn Data (<= 1000 Words)", padx=5, pady=5)
        learn_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        self.learn_data_text = scrolledtext.ScrolledText(learn_frame, width=40, height=5)
        self.learn_data_text.pack(fill="both", expand=True)

        # --- Outline Data Input ---
        outline_data_frame = tk.LabelFrame(right_frame, text="Outline Data", padx=5, pady=5)
        outline_data_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        self.outline_data_text = scrolledtext.ScrolledText(outline_data_frame, width=40, height=10)
        self.outline_data_text.pack(fill="both", expand=True)

        # --- Other Settings và Buttons ---
        other_frame = tk.LabelFrame(right_frame, text="Other Settings", padx=5, pady=5)
        other_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

   

        tk.Label(other_frame, text="Keywords:").grid(row=1, column=0, sticky="nw")
        self.keywords_text = scrolledtext.ScrolledText(other_frame, width=40, height=10)
        self.keywords_text.grid(row=1, column=1, columnspan=2)
        

        tk.Label(other_frame, text="Brand keyword:").grid(row=0, column=0, sticky="w")

        self.brand_keyword_text = tk.Entry(other_frame, width=20)
        self.brand_keyword_text.grid(row=0, column=1, sticky="ew")


        tk.Label(other_frame, text="Word Density:").grid(row=2, column=0, sticky="w")
        self.word_density_var = tk.StringVar(value="2.8")
        word_density_options = ["1.5","2.0", "2.5", "2.8", "3.5"]
        word_density_menu = tk.OptionMenu(other_frame, self.word_density_var, *word_density_options)
        word_density_menu.grid(row=2, column=1, sticky="w")
                


        action_frame = tk.Frame(right_frame, padx=5, pady=5)
        action_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        self.generate_outline_btn = tk.Button(action_frame, text="Generate Outline", command=lambda: self.run_in_thread(self.task_generate_outline, self.generate_outline_btn))
        self.generate_outline_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.write_full_btn = tk.Button(action_frame, text="Write", command=lambda: self.run_in_thread(self.task_write_full, self.write_full_btn))
        self.write_full_btn.grid(row=1, column=0, padx=5, pady=5, sticky="ew")



        # Tạo Frame chứa các nút liên quan đến Vietnamese & English
        self.lang_restricted_frame = tk.Frame(action_frame, padx=5, pady=5, relief="groove", bd=2)
        self.lang_restricted_frame.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

        # Thêm nhãn để thông báo ngôn ngữ giới hạn
        self.lang_label = tk.Label(self.lang_restricted_frame, text="Vietnamese and English only", fg="red", font=("Arial", 10, "bold"))
        self.lang_label.pack(pady=(0, 5))

        # Các nút chức năng
        self.normalize_seo_btn = tk.Button(self.lang_restricted_frame, text="Normalize SEO [2000]", 
                                        command=lambda: self.confirm_action("Normalize SEO", self.task_normalize_seo, self.normalize_seo_btn))
        self.normalize_seo_btn.pack(fill="x", padx=5, pady=2)

        self.fix_spin_folder_btn = tk.Button(self.lang_restricted_frame, text="Fix Spineditor [10,000]", 
                                            command=lambda: self.confirm_action("Fix Spineditor", self.task_fix_spin_folder, self.fix_spin_folder_btn))
        self.fix_spin_folder_btn.pack(fill="x", padx=5, pady=2)

        self.task_format_time_btn = tk.Button(self.lang_restricted_frame, text="Format Time new roman", 
                                            command=lambda: self.run_in_thread(self.task_format_time, self.task_format_time_btn))
        self.task_format_time_btn.pack(fill="x", padx=5, pady=2)

        self.task_format_arial_btn = tk.Button(self.lang_restricted_frame, text="Format Arial", 
                                            command=lambda: self.run_in_thread(self.task_format_arial, self.task_format_arial_btn))
        self.task_format_arial_btn.pack(fill="x", padx=5, pady=2)

        
        self.status_label = tk.Label(self, text="Idle")
        self.status_label.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

    def load_settings(self):
        # Đẩy các giá trị từ console_config vào các widget

        # Số lượng H2
        self.entry_number_of_h2.delete(0, tk.END)
        self.entry_number_of_h2.insert(0, console_config.get("number_of_h2", 3))

        # Số lượng H3
        self.entry_number_of_h3.delete(0, tk.END)
        self.entry_number_of_h3.insert(0, console_config.get("number_of_h3", 3))

        # Số lượng Images
        self.entry_number_of_image.delete(0, tk.END)
        self.entry_number_of_image.insert(0, console_config.get("number_of_image", 3))

        # Max Key
        self.entry_max_key.delete(0, tk.END)
        self.entry_max_key.insert(0, console_config.get("max_key", 17))

        # Language: sử dụng OptionMenu với StringVar
        self.language_var.set(console_config.get("language", "Vietnamese"))

        # SEO Category: sử dụng OptionMenu với StringVar
        self.seo_category_var.set(console_config.get("seo_category", "Formal"))

        # Input Folder
        self.input_folder_entry.delete(0, tk.END)
        self.input_folder_entry.insert(0, console_config.get("input_folder", r"C:\Users\Hieu Pham\Downloads\SEO"))

        # Các model settings
        self.write_model_var.set(console_config.get("model_write", "Gemini 1.5 pro [3000]"))
        self.outline_model_var.set(console_config.get("model_outline", "Gemini 1.5 pro [1000]"))
        self.normalize_model_var.set(console_config.get("model_normalize", "Chatgpt 4o mini [2000]"))

        # Learn Data
        self.learn_data_text.delete("1.0", tk.END)
        self.learn_data_text.insert(tk.END, console_config.get("learn_data", ""))

        # Outline Data
        self.outline_data_text.delete("1.0", tk.END)
        self.outline_data_text.insert(tk.END, console_config.get("outline_data", ""))

        # Keywords
        self.keywords_text.delete("1.0", tk.END)
        self.keywords_text.insert(tk.END, console_config.get("keywords", ""))

        # Brand keyword
        self.brand_keyword_text.delete(0, tk.END)  # Xóa nội dung cũ
        self.brand_keyword_text.insert(0, console_config.get("brand_keyword", ""))  # Chèn nội dung mới

        self.word_density_var.set(str(console_config.get("word_density", 2.8)))


    def save_settings(self):
        # Thu thập giá trị hiện tại từ giao diện và cập nhật vào phần cấu hình SEO Write
        new_config = {
            "number_of_h2": int(self.entry_number_of_h2.get()),
            "number_of_h3": int(self.entry_number_of_h3.get()),
            "number_of_image": int(self.entry_number_of_image.get()),
            "max_key": int(self.entry_max_key.get()),
            "language": self.language_var.get(),
            "seo_category": self.seo_category_var.get(),
            "input_folder": self.input_folder_entry.get(),
            "model_write": self.write_model_var.get(),
            "model_outline": self.outline_model_var.get(),
            "model_normalize": self.normalize_model_var.get(),
            "learn_data": self.learn_data_text.get("1.0", tk.END).strip(),
            "outline_data": self.outline_data_text.get("1.0", tk.END).strip(),
            "brand_keyword": self.brand_keyword_text.get().strip(),
            "keywords": self.keywords_text.get("1.0", tk.END).strip(),
            "word_density": float(self.word_density_var.get()),
        }
        
        # Cập nhật console_config toàn cục (nếu cần dùng lại)
        global console_config
        console_config = new_config
        save_console_config(new_config)

    def confirm_action(self, action_name, task_func, button):
        confirm = messagebox.askyesno("Confirmation", f"Are you sure you want to perform '{action_name}'?")
        if confirm:
            self.run_in_thread(task_func, button)
            
    def browse_folder(self, entry):
        folder = filedialog.askdirectory()
        if folder:
            entry.delete(0, tk.END)
            entry.insert(0, folder)

    def run_in_thread(self, task, button):
        def thread_target():
            error_message = None  # Khai báo trước để tránh NameError
            try:
                task()
            except Exception as e:
                error_message = str(e)  # Lưu lỗi vào biến
            finally:
                self.after(0, lambda: button.config(state=tk.NORMAL))
                self.after(0, lambda: self.status_label.config(text="Idle"))
                self.save_settings()

                # Nếu có lỗi, hiển thị messagebox
                if error_message:
                    self.after(0, lambda: messagebox.showerror("Error", error_message))
                else:
                    self.after(0, lambda: messagebox.showinfo("Success", "Task finished successfully."))


        button.config(state=tk.DISABLED)
        self.status_label.config(text="Running...")
        threading.Thread(target=thread_target, daemon=True).start()


    def task_generate_outline(self):
        base_folder_val = self.input_folder_entry.get()
        keywords = self.keywords_text.get("1.0", tk.END).strip().splitlines()

        # if len(keywords) > 5:
        #     messagebox.showwarning("Keyword Limit", "Only a maximum of 5 keywords are allowed at a time.")
        #     return

        learn_data = self.learn_data_text.get("1.0", tk.END).strip() or ""
        outline_data = self.outline_data_text.get("1.0", tk.END).strip() or ""
        h2 = int(self.entry_number_of_h2.get())
        h3 = int(self.entry_number_of_h3.get())
        language_val = self.language_var.get()
        seo_category_val = self.seo_category_var.get()
        selected = self.outline_model_var.get()
        outline_model_type_val, outline_model_name_val = MODEL_OPTIONS_OUTLINE[selected]
        
        print("[INFO] Generating outline...")
        generate_outline_func(
            keywords,
            base_folder_val,
            h2, h3,
            outline_model_name_val,
            outline_model_type_val,
            learn_data,
            language_val,
            seo_category_val,
            outline_data
        )
        print("[INFO] Generate outline - DONE.")

    def task_write_full(self):
        base_folder_val = self.input_folder_entry.get()
        keywords = self.keywords_text.get("1.0", tk.END).strip().splitlines()

        learn_data = self.learn_data_text.get("1.0", tk.END).strip() or ""
        selected = self.write_model_var.get()
        write_model_type_val, write_model_name_val = MODEL_OPTIONS_WRITE[selected]
        language_val = self.language_var.get()
        seo_category_val = self.seo_category_var.get()
        

        print("[INFO] Writing full content...")
        write_full_func(
            keywords,
            base_folder_val,
            write_model_name_val,
            write_model_type_val,
            language_val,
            seo_category_val,
            learn_data
        )
        print("[INFO] Write full - DONE.")


    #Vietnamese and English only
    from tkinter import messagebox

    def task_normalize_seo(self):
        base_folder_val = self.input_folder_entry.get()
        max_key_val = int(self.entry_max_key.get())
        number_of_image_val = int(self.entry_number_of_image.get())
        
        selected = self.normalize_model_var.get()
        normalize_model_type_val, normalize_model_name_val = MODEL_OPTIONS_NORMALIZE[selected]   
        language_val = self.language_var.get()
        brand_key = self.brand_keyword_text.get()

        if language_val not in ["Vietnamese", "English"]:
            messagebox.showerror("Error", "Only Vietnamese and English are supported!")
            return

        print("[INFO] Normalize SEO...")
        word_density = float(self.word_density_var.get())

        normalize_seo(
            base_folder_val,
            max_key_val,
            number_of_image_val,
            normalize_model_name_val,
            normalize_model_type_val,
            language_val,
            word_density,
            brand_key  
        )
        print("[INFO] Normalize SEO - DONE.")

    def task_fix_spin_folder(self):
        base_folder_val = self.input_folder_entry.get()
        normalize_seos_folder = os.path.join(base_folder_val, "full_seos", "normalize_seos")
        selected = self.normalize_model_var.get()
        normalize_model_type_val, normalize_model_name_val = MODEL_OPTIONS_NORMALIZE[selected]   
        temperature_val = DEFAULT_TEMPERATURE
        is_spineditor_val = DEFAULT_IS_SPIN_EDITOR
        search_type_val = DEFAULT_SEARCH_TYPE

        language_val = self.language_var.get()
        if language_val not in ["Vietnamese", "English"]:
            messagebox.showerror("Error", "Only Vietnamese and English are supported!")
            return

        print("[INFO] Fix spin folder...")
        fix_spin_folder(
            normalize_seos_folder,
            normalize_model_name_val,
            normalize_model_type_val,
            temperature_val,
            is_spineditor_val,
            search_type_val
        )
        print("[INFO] Fix spin folder - DONE.")

    def task_format_time(self):
        base_folder_val = self.input_folder_entry.get()
        normalize_seos_folder = os.path.join(base_folder_val, "full_seos", "normalize_seos")

        language_val = self.language_var.get()
        if language_val not in ["Vietnamese", "English"]:
            messagebox.showerror("Error", "Only Vietnamese and English are supported!")
            return

        print("[INFO] Format time new roman...")
        process_folder_format_time(normalize_seos_folder)
        print("[INFO] Format time new roman - DONE.")

    def task_format_arial(self):
        base_folder_val = self.input_folder_entry.get()
        normalize_seos_folder = os.path.join(base_folder_val, "full_seos", "normalize_seos")

        language_val = self.language_var.get()
        if language_val not in ["Vietnamese", "English"]:
            messagebox.showerror("Error", "Only Vietnamese and English are supported!")
            return

        print("[INFO] Format Arial...")
        process_folder_format_arial(normalize_seos_folder)
        print("[INFO] Format Arial - DONE.")
