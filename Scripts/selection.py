# -*- coding: utf-8 -*-

from .modules import tk, ConfigParser


class Selection(tk.Toplevel):
    def __init__(self, title, catalogue, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title(title)
        self.config = ConfigParser()
        self.config.read("defaults.ini")
        self.selected = self.config[title.upper()]["selected"]
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
        self.wm_protocol("WM_DELETE_WINDOW", lambda: None)

    def apply(self, title):
        pass


class SingleSelection(Selection):
    def __init__(self, title, catalogue, *args, **kwargs):
        super().__init__(
            title=title,
            catalogue=catalogue,
            *args,
            **kwargs
        )
        for i, j in enumerate(self.catalogue):
            var = tk.StringVar()
            if j == self.selected:
                var.set(value="1")
            else:
                var.set(value="0")
            checkbutton = tk.Checkbutton(
                master=self.topframe,
                text=j,
                variable=var
            )
            checkbutton.grid(row=i, column=0, sticky="w")
            self.checkbuttons[j] = [checkbutton, var]
            self.configure_checkbuttons(option=j)
        self.wait_window()

    def apply(self, title):
        config = ConfigParser()
        config.read("defaults.ini")
        for i in self.catalogue:
            if self.checkbuttons[i][1].get() == "1":
                config[title] = {"selected": i}
                with open("defaults.ini", "w") as f:
                    config.write(f)
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


class MultipleSelection(Selection):
    def __init__(self, title, catalogue, *args, **kwargs):
        super().__init__(
            title=title,
            catalogue=catalogue,
            *args,
            **kwargs
        )
        self.geometry("200x400")
        self.check_all = tk.StringVar()
        self.check_all.set("0")
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
            var = tk.StringVar()
            if j in self.selected:
                var.set("1")
            else:
                var.set("0")
            checkbutton = tk.Checkbutton(
                master=self.topframe,
                text=j,
                variable=var
            )
            checkbutton.grid(row=i + 1, column=0, sticky="w")
            self.checkbuttons[j] = [checkbutton, var]
        self.check_all.set("0")
        self.select_all["command"] = self.check_all_command

    def check_all_command(self):
        if self.check_all.get() == "1":
            for values in self.checkbuttons.values():
                values[-1].set("1")
                values[0].configure(variable=values[-1])
        else:
            for values in self.checkbuttons.values():
                values[-1].set(",")
                values[0].configure(variable=values[-1])

    def apply(self, title):
        selected = []
        config = ConfigParser()
        config.read("defaults.ini")
        for k, v in self.checkbuttons.items():
            if v[1].get() == "1" and k != "Check/Uncheck All":
                selected.append(k)
        config[title] = {"selected": ", ".join(selected)}
        with open("defaults.ini", "w") as f:
            config.write(f)
        self.destroy()
