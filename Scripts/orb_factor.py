# -*- coding: utf -*-

from .modules import tk, ttk, ConfigParser


class OrbFactor(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Orb Factor")
        self.resizable(width=False, height=False)
        self.aspects = [
            "Conjunction",
            "Semi-Sextile",
            "Semi-Square",
            "Sextile",
            "Quintile",
            "Square",
            "Trine",
            "Sesquiquadrate",
            "BiQuintile",
            "Quincunx",
            "Opposite"
        ]
        self.frame = tk.Frame(master=self)
        self.frame.pack()
        self.widgets = {}
        self.create_widgets()
        self.button = tk.Button(master=self, text="Apply", command=self.apply)
        self.button.pack()

    def create_widgets(self):
        config = ConfigParser()
        config.read("defaults.ini")
        for i, j in enumerate(self.aspects):
            label = tk.Label(master=self.frame, text=f"{j}")
            label.grid(row=i, column=0, sticky="w")
            entry = ttk.Entry(master=self.frame, width=4)
            entry.grid(row=i, column=1)
            entry.insert(0, config["ORB FACTORS"][j])
            self.widgets[j] = entry

    def apply(self):
        config = ConfigParser()
        config.read("defaults.ini")
        for k, v in self.widgets.items():
            config["ORB FACTORS"][k] = v.get()
        with open("defaults.ini", "w", encoding="utf-8") as f:
            config.write(f)
        self.destroy()
