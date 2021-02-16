# -*- coding: utf-8 -*-

from .modules import tk, ttk


class SearchFrame(tk.Frame):
    def __init__(
            self,
            database,
            treeview,
            info_var,
            *args,
            **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.menu = None
        self.id = None
        self.found = None
        self.treeview = treeview
        self.label = tk.Label(
            master=self,
            text="Search a name",
            fg="red"
        )
        self.label.pack()
        self.frame = tk.Frame(master=self)
        self.frame.pack()
        self.button = tk.Button(
            master=self.frame,
            text="Add",
            command=lambda: self.command(
                info_var=info_var
            )
        )
        self.var = tk.StringVar()
        self.combobox = ttk.Combobox(
            master=self.frame,
            values=[],
            textvariable=self.var
        )
        self.combobox.bind(
            sequence="<Return>",
            func=lambda event: self.search(database=database)
        )
        self.combobox.bind(
            sequence="<<ComboboxSelected>>",
            func=lambda event: self.change_item()
        )
        self.combobox.bind(
            sequence="<KeyRelease>",
            func=lambda event: self.button.pack_forget()
        )
        self.combobox.pack(side="left")
        self.grid(row=0, column=0, columnspan=4)

    def is_added(self):
        added = False
        for child in self.treeview.get_children():
            values = self.treeview.item(child)["values"]
            if int(self.id) == int(values[0]):
                added = True
        return added

    def command(self, info_var):
        if not self.is_added():
            values = [self.id] + self.found[self.id]
            self.treeview.insert(
                index=len(self.treeview.get_children()),
                parent="",
                values=values
            )
            info_var.set(len(self.treeview.get_children()))
            self.button.pack_forget()

    def change_item(self):
        self.id = self.var.get().split(" ")[0]
        if not self.is_added():
            self.button.pack(side="right")

    def search(self, database):
        if self.combobox.get():
            value = self.combobox.get().lower()
            self.id = None
            self.found = {
                str(i[0]): i[1:]
                for i in database
                if value.lower() in i[1].lower()
            }
            if self.found:
                self.combobox["values"] = [
                    f"{k} {v[0]}" for k, v in self.found.items()
                ]
                self.combobox.event_generate("<Down>")
        else:
            self.combobox["values"] = []
