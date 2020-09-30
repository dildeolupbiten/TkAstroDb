# -*- coding: utf-8 -*-

from .modules import tk


class MsgBox(tk.Toplevel):
    msgbox = []

    def __init__(self, title, message, level, icons):
        super().__init__()
        for i in self.msgbox:
            i.destroy()
        self.geometry("350x100")
        self.title(title)
        self.resizable(width=False, height=False)
        self.icons = icons
        self.level_icon = self.icons[level]["img"]
        self.frame = tk.Frame(master=self)
        self.frame.pack(expand=True, fill="both")
        self.icon_label = tk.Label(
            master=self.frame,
            image=self.level_icon,
        )
        self.icon_label.pack(side="left", expand=True, fill="both")
        self.message_label = tk.Label(
            master=self.frame,
            text=message,
            font="Arial 14 bold",
            anchor="w"
        )
        self.message_label.pack(side="left", expand=True, fill="both")
        self.button = tk.Button(
            master=self,
            text="OK",
            command=self.destroy,
        )
        self.button.pack(side="bottom")
        self.msgbox.append(self)
        self.wait_window()
        
