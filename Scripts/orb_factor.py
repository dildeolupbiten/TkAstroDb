# -*- coding: utf -*-

from .modules import tk, ttk, ConfigParser


class OrbFactor(tk.Toplevel):
    def __init__(
        self,
        title,
        catalogue,
        config_key,
        *args,
        **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.title(title)
        self.resizable(width=False, height=False)
        self.catalogue = catalogue
        self.frame = tk.Frame(master=self)
        self.frame.pack()
        self.widgets = {}
        self.create_widgets(config_key)
        self.button = tk.Button(
            master=self,
            text="Apply",
            command=lambda: self.apply(config_key)
        )
        self.button.pack()

    def create_widgets(self, config_key):
        config = ConfigParser()
        config.read("defaults.ini")
        for i, j in enumerate(self.catalogue):
            label = tk.Label(master=self.frame, text=f"{j}")
            label.grid(row=i, column=0, sticky="w")
            entry = ttk.Entry(master=self.frame, width=4)
            entry.grid(row=i, column=1)
            entry.insert(0, config[config_key][j])
            self.widgets[j] = entry

    def apply(self, config_key):
        config = ConfigParser()
        config.read("defaults.ini")
        for k, v in self.widgets.items():
            config[config_key][k] = v.get()
        with open("defaults.ini", "w", encoding="utf-8") as f:
            config.write(f)
        self.destroy()
