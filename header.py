import tkinter as tk

class Header(tk.Frame):
    def __init__(self, parent, switch_callback, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.switch_callback = switch_callback

        options = [
            ("SEO Write", "seo_write"),
            ("Image Optimization", "image_optimization"),
            ("Insert hyperlink (anchor text)", "hyper_link"),
        ]

        for text, name in options:
            btn = tk.Button(self, text=text, command=lambda n=name: self.switch_callback(n))
            btn.pack(side=tk.LEFT, padx=5, pady=5)