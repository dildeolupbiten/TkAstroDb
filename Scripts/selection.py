# -*- coding: utf-8 -*-

from .constants import AYANAMSHA
from .modules import tk, ttk, ConfigParser


class Selection(tk.Toplevel):
    def __init__(self, title, catalogue, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title(title)
        self.config = ConfigParser()
        self.config.read("defaults.ini")
        self.catalogue = catalogue
        self.resizable(width=False, height=False)
        self.topframe = tk.Frame(master=self)
        self.topframe.pack(side="top")
        self.bottomframe = tk.Frame(master=self)
        self.bottomframe.pack(side="bottom")
        self.checkbuttons = {}
        self.button = tk.Button(
            master=self.bottomframe,
            text="Apply",
            command=lambda: self.apply(title=title.upper())
        )
        self.button.pack()

    def apply(self, title):
        pass


class SingleSelection(Selection):
    def __init__(self, title, catalogue, *args, **kwargs):
        self.done = False
        super().__init__(
            title=title,
            catalogue=catalogue,
            *args,
            **kwargs
        )
        for i, j in enumerate(self.catalogue):
            var = tk.BooleanVar()
            if self.config[title.upper()]["selected"] == j:
                var.set(True)
            checkbutton = tk.Checkbutton(
                master=self.topframe,
                text=j,
                variable=var
            )
            checkbutton.grid(row=i, column=0, sticky="w")
            self.checkbuttons[j] = [checkbutton, var]
            self.configure_checkbuttons(option=j)
        self.ayanamsha = self.config["AYANAMSHA"]["selected"]
        if (
            "Sidereal" in self.checkbuttons 
            and 
            self.checkbuttons["Sidereal"][1].get()
        ):         
            self.create_frame(catalogue=AYANAMSHA)
        else:
            self.frame = None
        self.wait_window()

    def apply(self, title):
        config = ConfigParser()
        config.read("defaults.ini")
        for i in self.catalogue:
            if self.checkbuttons[i][1].get():
                config[title] = {"selected": i}
                if i == "Sidereal":
                    config["AYANAMSHA"] = {"selected": self.ayanamsha}
                with open("defaults.ini", "w") as f:
                    config.write(f)
        self.done = True
        self.destroy()

    def configure_checkbuttons(self, option):
        return self.checkbuttons[option][0].configure(
            command=lambda: self.check_uncheck(option=option)
        )

    def check_uncheck(self, option):
        for i in self.catalogue:
            if i != option:
                self.checkbuttons[i][1].set("0")
                self.checkbuttons[i][0].configure(
                    variable=self.checkbuttons[i][1]
                )
        if "Sidereal" in self.checkbuttons:
            if self.checkbuttons["Sidereal"][1].get():
                self.create_frame(catalogue=AYANAMSHA)
            else:
                self.frame.destroy()
                self.frame = None

    def create_frame(self, catalogue):
        self.frame = tk.Frame(master=self.topframe)
        self.frame.grid(row=len(self.catalogue), column=0)
        label = tk.Label(
            master=self.frame, 
            text="Ayanamsha", 
            font="Default 10 bold"
        )
        label.pack()
        style = ttk.Style()
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "white")]
        )
        combobox = ttk.Combobox(
            master=self.frame, 
            values=list(catalogue),
            style="TCombobox"
        )
        combobox.pack()
        combobox.insert("0", self.ayanamsha)
        combobox["state"] = "readonly"
        combobox.bind(
            sequence="<<ComboboxSelected>>", 
            func=self.change_ayanamsha
        )
    
    def change_ayanamsha(self, event):
        self.ayanamsha = event.widget.get()
        event.widget.selection_clear()


class MultipleSelection(Selection):
    def __init__(
            self, 
            title, 
            catalogue, 
            get=False, 
            *args, 
            **kwargs
    ):
        super().__init__(
            title=title,
            catalogue=catalogue,
            *args,
            **kwargs
        )
        self.result = []
        self.get = get
        self.check_all = tk.BooleanVar()
        self.check_all.set(False)
        self.select_all = tk.Checkbutton(
            master=self.topframe,
            text="Check/Uncheck All",
            variable=self.check_all
        )
        self.select_all.grid(row=0, column=0, sticky="w")
        self.checkbuttons["Check/Uncheck All"] = [
            self.select_all, self.check_all
        ]
        for i, j in enumerate(catalogue):
            var = tk.BooleanVar()
            if not self.get:
                if self.config[title.upper()][j] == "true":
                    var.set(True)
                else:
                    var.set(False)
                text = j.title().replace("_", " ")
            else:
                text = j
            checkbutton = tk.Checkbutton(
                master=self.topframe,
                text=text,
                variable=var
            )
            checkbutton.grid(row=i + 1, column=0, sticky="w")
            self.checkbuttons[text] = [checkbutton, var]
        for k, v in self.checkbuttons.items():
            v[0]["command"] = self.check_command
        self.select_all["command"] = self.check_all_command
        if self.get:
            self.wait_window()

    def check_all_command(self):
        if self.check_all.get():
            for key, value in self.checkbuttons.items():
                value[-1].set(True)                  
        else:
            for key, value in self.checkbuttons.items():
                value[-1].set(False)

    def check_command(self):
        for key, value in self.checkbuttons.items():
            if key in [
                "Sum Of Aspects",
                "Yod",
                "T-Square",
                "Grand Trine",
                "Mystic Rectangle",
                "Grand Cross",
                "Kite"
            ]:
                if value[-1].get():
                    self.checkbuttons["Aspects"][-1].set(True)
            if key == "Sum Of Midpoints":
                if value[-1].get():
                    self.checkbuttons["Midpoints"][-1].set(True)

    def apply(self, title):
        if not self.get:
            selected = {}
            config = ConfigParser()
            config.read("defaults.ini")
            for k, v in self.checkbuttons.items():
                if k == "Check/Uncheck All":
                    continue
                if v[1].get():
                    selected[k.replace(" ", "_")] = "true"
                else:
                    selected[k.replace(" ", "_")] = "false"
            config[title] = selected
            with open("defaults.ini", "w") as f:
                config.write(f)
        else:
            for k, v in self.checkbuttons.items():
                if k == "Check/Uncheck All":
                    continue
                if v[1].get():
                    self.result.append(k)
        self.destroy()
