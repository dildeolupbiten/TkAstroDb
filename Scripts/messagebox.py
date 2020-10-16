# -*- coding: utf-8 -*-

from .modules import tk


class MsgBox(tk.Toplevel):
    msgbox = []

    def __init__(
            self, 
            title, 
            message, 
            level, 
            icons, 
            width=350, 
            height=100,
            wait=True,
    ):
        super().__init__()
        for i in self.msgbox:
            i.destroy()
        self.geometry(f"{width}x{height}")
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
        self.button_frame = tk.Frame(master=self)
        self.button_frame.pack(side="bottom")
        self.button = tk.Button(
            master=self.button_frame,
            text="OK",
            command=self.destroy,
        )
        self.button.pack(side="bottom")
        self.msgbox.append(self)
        if wait:
            self.wait_window()
        
        
class ChoiceBox(MsgBox):
    def __init__(
            self, 
            wait=False, 
            width=350, 
            height=100,
            *args,
            **kwargs
    ):
        super().__init__(
            wait=wait, 
            width=width, 
            height=height, 
            *args, 
            **kwargs
        )
        self.choice = False
        self.button["command"] = self.ok
        self.button["width"] = 7
        self.button.pack_forget()
        self.button.pack(side="left", padx=10)
        self.cancel_button = tk.Button(
            master=self.button_frame,
            text="Cancel",
            command=self.cancel,
            width=7
        )
        self.cancel_button.pack(side="right", padx=10)
        self.wait_window()

    def ok(self):
        self.choice = True
        self.destroy()

    def cancel(self):
        self.choice = False
        self.destroy()
