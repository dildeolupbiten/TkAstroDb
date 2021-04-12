# -*- coding: utf-8 -*-

from .messagebox import MsgBox
from .utilities import delete_nonnumeric_chars
from .modules import (
    np, tk, ttk, plt, binom, FigureCanvasTkAgg, NavigationToolbar2Tk
)


class Plot(tk.Toplevel):
    def __init__(self, icons, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Plot")
        self.icons = icons
        self.resizable(width=False, height=False)
        self.left_frame = tk.Frame(master=self, width=200)
        self.left_frame.pack(side="left")
        self.right_frame = tk.Frame(master=self, width=600)
        self.right_frame.pack(side="right")
        self.figure = plt.Figure()
        self.canvas = FigureCanvasTkAgg(
            figure=self.figure,
            master=self.right_frame
        )
        self.canvas.get_tk_widget().pack()
        self.navbar = NavigationToolbar2Tk(
            canvas=self.canvas,
            window=self.right_frame
        )
        self.create_entries()

    def create_entries(self):
        title = tk.Label(
            master=self.left_frame,
            text="Title",
            font="Default 11 bold"
        )
        title.pack()
        entry = ttk.Entry(master=self.left_frame, width=30)
        entry.pack()
        frame = tk.Frame(master=self.left_frame)
        frame.pack()
        entries = {"title": entry}
        for i, (j, k) in enumerate(zip(["Case", "Control"], ["n", "X"])):
            tk.Label(
                master=frame,
                text=j,
                font="Default 11 bold"
            ).grid(row=0, column=i + 1)
            tk.Label(
                master=frame,
                text=k,
                font="Default 11 bold"
            ).grid(row=1 + i, column=0)
            entries[j] = {}
            for m, n in enumerate(["n", "X"]):
                entry = ttk.Entry(master=frame, width=8)
                entry.grid(row=1 + m, column=i + 1)
                entry.bind(
                    sequence="<KeyRelease>",
                    func=lambda event: delete_nonnumeric_chars(
                        event=event,
                        _type=int
                    )
                )
                entries[j][n] = entry
        button = tk.Button(
            master=frame,
            text="Plot",
            command=lambda: self.plot(entries)
        )
        button.grid(row=3, column=0, columnspan=3)

    @staticmethod
    def draw_binomial_distribution(
        ax,
        n,
        k,
        title,
        color,
        confidence_interval=False
    ):
        values = {m: binom.pmf(n=n, p=k / n, k=m) for m in range(n + 1)}
        x = np.array(list(values.keys()))
        y = np.array(list(values.values()))
        p = k / n
        ax.plot(x, y, label=title, color=color)
        sd = (k * (1 - p)) ** 0.5
        mean = [(k, k), (0, values[int(k)])]
        if int(k) == k:
            k = int(k)
        else:
            k = round(k, 2)
        ax.plot(
            mean[0],
            mean[1],
            color=color,
            label=f"n = {n}\nX = {k}",
            linestyle="--"
        )
        if confidence_interval:
            for key, value in {
                "#0000ff" if color == "blue" else "#ff0000"
                if color == "red" else "#00ff00": (0.98, "85 %"),
                "#4444ff" if color == "blue" else "#ff4444"
                if color == "red" else "#44ff44": (1.96, "97.5 %"),
                "#8888ff" if color == "blue" else "#ff8888"
                if color == "red" else "#88ff88": (2.58, "99.5 %")
            }.items():
                ax.fill_between(
                    x,
                    y,
                    where=np.logical_and(
                        x < k + value[0] * sd,
                        x > k - value[0] * sd
                    ),
                    color=key,
                    alpha=0.7,
                    label=f"{value[0]} SD ({value[1]})"
                )

    def plot(self, entries):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        title = entries["title"].get()
        try:
            n_case = int(entries["Case"]["n"].get())
        except ValueError:
            MsgBox(
                title="Warning",
                level="warning",
                message="Invalid value for n case.",
                icons=self.icons
            )
            return
        try:
            x_case = int(entries["Case"]["X"].get())
        except ValueError:
            MsgBox(
                title="Warning",
                level="warning",
                message="Invalid value for X case.",
                icons=self.icons
            )
            return
        try:
            n_control = int(entries["Control"]["n"].get())
        except ValueError:
            MsgBox(
                title="Warning",
                level="warning",
                message="Invalid value for n control.",
                icons=self.icons
            )
            return
        try:
            x_control = int(entries["Control"]["X"].get())
        except ValueError:
            MsgBox(
                title="Warning",
                level="warning",
                message="Invalid value for X control.",
                icons=self.icons
            )
            return
        self.draw_binomial_distribution(
            ax=ax,
            n=n_case,
            k=x_case,
            title="Observed",
            color="red",
            confidence_interval=True
        )
        self.draw_binomial_distribution(
            ax=ax,
            n=n_case,
            k=n_case/n_control*x_control,
            title="Expected",
            color="blue",
            confidence_interval=True
        )
        ax.set_xlabel("Number Of People")
        ax.set_ylabel("Probability Mass Function")
        ax.set_title(title)
        self.figure.legend(*ax.get_legend_handles_labels())
        self.canvas.draw()
