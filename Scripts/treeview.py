# -*- coding: utf-8 -*-

from .modules import np, tk, ttk


class Treeview(ttk.Treeview):
    def __init__(
            self,
            columns,
            values=None,
            *args,
            **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.columns = columns
        if values:
            self.style = ttk.Style()
            self.style.map(
                "Treeview",
                background=[("selected", "#00b4a7")],
                foreground=[("selected", "#000000")]
            )
            self.style.configure("Treeview.Heading", background="#ffffff")
            self.style.configure("Treeeview.Cell", fieldbackground="red")
            self.style.map(
                "Treeview.Heading",
                background=[("active", "#00b4a7")],
                foreground=[("active", "#000000")]
            )
        else:
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
            if values:
                if index == 0:
                    width = 200
                    anchor = "w"
                else:
                    width = 75
                    anchor = "center"
            else:
                width = 125
                anchor = "center"
            self.column(
                column=f"#{index + 1}",
                minwidth=75,
                width=width,
                anchor=anchor,
            )
            self._heading(col=index, text=column, values=values)
        self.bind(
            sequence="<Control-a>",
            func=lambda event: self.select_all()
        )
        self.bind(
            sequence="<Control-A>",
            func=lambda event: self.select_all()
        )
        if values:
            self.insert_values(values)

    def _heading(self, col, text, values):
        if values:
            self.heading(column=f"#{col + 1}", text=text)
        else:
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

    def insert_values(self, values):
        count = 0
        total = []
        for i in ["sign", "house"]:
            for k, v in values[i].items():
                if k in ["Dayscores", "Effect of Houses"]:
                    total.append(np.array([*v.values()][:-1]))
                    self.insert(
                        parent="",
                        index=count,
                        value=[k, *v.values()],
                        tags=("total",)
                    )
                    self.tag_configure('total', foreground="red")
                else:
                    self.insert(
                        parent="",
                        index=count,
                        value=[k, *v.values()]
                    )
                count += 1
        total = [round(float(i), 2) for i in total[0] * total[1]]
        total += [round(sum(total), 2)]
        arr = ["Total Scores"]
        arr += total
        self.insert(
            parent="",
            index=count,
            value=arr,
            tags=("last", )
        )
        self.tag_configure('last', foreground="blue")
