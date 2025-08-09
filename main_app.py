import tkinter as tk
from tkinter import messagebox
import sys
import os

from header import Header
from components.seo_write import SEOWriteFrame
from components.image_optimization import ImageOptimizationFrame
from components.hyper_link import InsertHyperlinkComponent

# Thêm các thư viện cần thiết cho việc chống debugger
import threading
import time

import os
import sys
import atexit
import tkinter as tk
from tkinter import messagebox

sys.path.append(os.path.abspath(os.path.dirname(__file__)))  # Ensure root path
sys.path.append(os.path.join(os.path.dirname(__file__), "src"))  # Add src directory

def check_debugger():
    """
    Kiểm tra xem có debugger nào được gắn vào tiến trình hay không.
    Nếu có, in ra thông báo và thoát chương trình.
    """
    if sys.gettrace() is not None:
        print("🚫 Debugger detected! Exiting...")
        sys.exit(1)

def anti_debugger_thread():
    """
    Một thread chạy ngầm để liên tục kiểm tra debugger.
    Phương pháp này nhằm phát hiện trường hợp debugger được gắn sau khi ứng dụng khởi động.
    """
    while True:
        check_debugger()
        time.sleep(1)  # Kiểm tra mỗi 1 giây

# Bắt đầu thread chống debugger
threading.Thread(target=anti_debugger_thread, daemon=True).start()

class RedirectText:
    """Redirect stdout to a Text widget for logging."""
    def __init__(self, widget):
        self.widget = widget
        self.stdout_backup = sys.stdout  # Save original stdout

    def write(self, string):
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)

    def flush(self):
        pass

    def restore(self):
        sys.stdout = self.stdout_backup  # Restore original stdout

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SEO Tools Interface")

        # Set window size
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.9)
        self.geometry(f"{window_width}x{window_height}+{(screen_width - window_width) // 2}+0")

        # Header
        self.header = Header(self, self.switch_component)
        self.header.pack(side=tk.TOP, fill=tk.X)

        # Main function container
        self.container = tk.Frame(self)
        self.container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Log frame with scrollbar
        log_frame = tk.Frame(self)
        log_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=(0, 5))

        self.log_text = tk.Text(log_frame, height=17, wrap="word", bg="#f5f5f5")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # Redirect stdout to logs
        self.logger = RedirectText(self.log_text)
        sys.stdout = self.logger

        # Initialize main components
        self.frames = {
            "seo_write": SEOWriteFrame(self.container),
            "image_optimization": ImageOptimizationFrame(self.container),
            "hyper_link": InsertHyperlinkComponent(self.container),
        }
        self.current_frame = None

        # Display the first component
        self.switch_component("seo_write")

        # Handle window close event
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def switch_component(self, component_name):
        if self.current_frame:
            self.current_frame.pack_forget()
        frame = self.frames.get(component_name)
        if frame:
            frame.pack(fill=tk.BOTH, expand=True)
            self.current_frame = frame

    def on_close(self):
        """Ask for confirmation before closing the app."""
        if messagebox.askokcancel("Exit Confirmation", "Are you sure you want to exit?"):
            # Lưu cấu hình nếu đang ở component SEO Write
            if isinstance(self.current_frame, SEOWriteFrame):
                self.current_frame.save_settings()
                print("⚙️ Đã lưu cấu hình SEO Write")
                
            self.logger.restore()  # Restore stdout before closing
            self.destroy()

PIDFILE = 'app.pid'

def check_single_instance():
    if os.path.exists(PIDFILE):
        print("Ứng dụng đã được mở sẵn!")
        sys.exit(1)

    with open(PIDFILE, 'w') as f:
        f.write(str(os.getpid()))

    def remove_pidfile():
        if os.path.exists(PIDFILE):
            os.remove(PIDFILE)
    atexit.register(remove_pidfile)

def main():
    # Kiểm tra single instance trước
    check_single_instance()

    # Tiếp tục các logic khởi tạo app
    print("✅ Launching MainApplication directly.")
    app = MainApplication()
    app.mainloop()

if __name__ == "__main__":
    main()