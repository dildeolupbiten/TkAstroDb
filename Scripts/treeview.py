# -*- coding: utf-8 -*-

from .modules import tk, ttk


class Treeview(ttk.Treeview):
    def __init__(
            self,
            columns,
            width=None,
            anchor=None,
            x_scrollbar=None,
            *args,
            **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.columns = columns
        if x_scrollbar:
            self.x_scrollbar = tk.Scrollbar(
                master=self.master,
                orient="horizontal"
            )
            self.configure(
                xscrollcommand=self.x_scrollbar.set,
                height=5
            )
            self.x_scrollbar.configure(command=self.xview)
            self.x_scrollbar.pack(side="bottom", fill="x")
        self.y_scrollbar = tk.Scrollbar(
            master=self.master,
            orient="vertical"
        )
        self.configure(
            yscrollcommand=self.y_scrollbar.set,
            style="Treeview"
        )
        self.y_scrollbar.configure(command=self.yview)
        self.y_scrollbar.pack(side="right", fill="y")
        self.configure(
            show="headings",
            columns=[f"#{i + 1}" for i in range(len(self.columns))],
            height=10,
            selectmode="extended"
        )
        self.pack(side="left", expand=True, fill="both")
        for index, column in enumerate(self.columns):
            if width:
                width = width
            else:
                width = 125
            if anchor:
                anchor = anchor
            else:
                anchor = "center"
            self.column(
                column=f"#{index + 1}",
                minwidth=75,
                width=width,
                anchor=anchor,
            )
            self._heading(col=index, text=column)
        self.bind(
            sequence="<Control-a>",
            func=lambda event: self.select_all()
        )
        self.bind(
            sequence="<Control-A>",
            func=lambda event: self.select_all()
        )

    def _heading(self, col, text):
        self.heading(
            column=f"#{col + 1}",
            text=text,
            command=lambda: self.sort(col=col, reverse=False)
        )

    def sort(self, col: int, reverse: bool):
        column = [
            (self.set(k, col), k)
            for k in self.get_children("")
        ]
        try:
            column.sort(key=lambda t: int(t[0]), reverse=reverse)
        except ValueError:
            column.sort(reverse=reverse)
        for index, (val, k) in enumerate(column):
            self.move(k, "", index)
        self.heading(
            column=col,
            command=lambda: self.sort(col=col, reverse=not reverse)
        )

    def select_all(self):
        for child in self.get_children():
            self.selection_add(child)
