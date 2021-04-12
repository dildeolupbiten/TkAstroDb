# -*- coding: utf-8 -*-

from .modules import tk, ttk
from .utilities import delete_nonnumeric_chars


class EntryFrame(tk.Frame):
    def __init__(self, texts, title, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.label = tk.Label(master=self, text=title, fg="red")
        self.label.pack()
        self.frame = tk.Frame(master=self)
        self.frame.pack()
        self.widgets = self.create_widgets(texts=texts)
        
    def create_widgets(self, texts):
        widgets = {}
        for index, text in enumerate(texts, 1):
            label = tk.Label(master=self.frame, text=text, fg="blue")
            label.grid(row=index, column=0, sticky="w")
            entry = ttk.Entry(master=self.frame, width=5)
            entry.grid(row=index, column=1, sticky="w")
            entry.bind(
                sequence="<KeyRelease>",
                func=delete_nonnumeric_chars
            )
            widgets[text] = entry
        return widgets
