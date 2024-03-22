# -*- coding: utf-8 -*-

from .frame import Frame
from .libs import (
    np, Qt, norm, binom, QFrame, QHBoxLayout, QVBoxLayout,
    QLabel, Figure, FigureCanvas, QSpinBox, NavigationToolbar
)


class Navbar(NavigationToolbar):
    toolitems = [
        item
        for item in NavigationToolbar.toolitems
        if item[0] in ("Home", "Back", "Forward", "Pan", "Zoom", "Save")
    ]


class Canvas(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.figure = Figure(facecolor="#212529")
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = Navbar(self.canvas, self)
        layout = QVBoxLayout()
        layout.addWidget(self.canvas)
        layout.addWidget(self.toolbar)
        self.setStyleSheet("border: 1 solid #343a40; border-radius: 5px;")
        self.setLayout(layout)

    def draw_binomial_distribution(
        self,
        ns,
        ks,
        colors
    ):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_axis_off()
        for n, k, color, n_case in zip(ns, ks, colors, [0, ns[0]]):
            if not n:
                return
            if k > n:
                return
            p = k / n
            if n_case and k >= 20:
                sd = ((k * (1 - p)) ** 0.5) * n_case / n
                k = k * n_case / n
                values = np.linspace(k - 3 * sd, k + 3 * sd, 100)
                x = values
                y = norm.pdf(values, k, sd)
                values = {i: j for i, j in zip(x, y)}
                diffs = [abs(k - i) for i in x]
                min_diff = diffs.index(min(diffs))
                mean = [(k, k), (0, values[x[min_diff]])]
            else:
                n_case = n
                sd = (k * (1 - p)) ** 0.5
                values = {m: binom.pmf(n=n, p=k/n, k=m) for m in range(n + 1)}
                x = np.array(list(values.keys()))
                y = np.array(list(values.values()))
                mean = [(k, k), (0, values[int(k)])]
            k = int(k) if int(k) == k else round(k, 2)
            label = f"n = {n_case}\nX = {k}"
            ax.plot(x, y, color=color)
            ax.plot(
                mean[0],
                mean[1],
                color=color,
                label=label,
                linestyle="--"
            )
            for key, value in {
                "#0000ff" if color == "blue" else "#ff0000"
                if color == "red" else "#00ff00": (0.98, "85 %"),
                "#5555ff" if color == "blue" else "#ff5555"
                if color == "red" else "#77ff77": (1.96, "97.5 %"),
                "#aaaaff" if color == "blue" else "#ffaaaa"
                if color == "red" else "#aaffaa": (2.58, "99.5 %")
            }.items():
                ax.fill_between(
                    x,
                    y,
                    where=np.logical_and(x < k + value[0] * sd, x > k - value[0] * sd),
                    color=key,
                    alpha=0.7,
                    label=f"{value[0]} SD ({value[1]})"
                )
        self.figure.legend(
            *ax.get_legend_handles_labels(),
            facecolor="#212529",
            edgecolor="#212529",
            labelcolor="white"
        )
        self.canvas.draw()


class Comparison(Frame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.frame = QFrame(self)
        self.frame.box = QVBoxLayout(self.frame)
        self.frame.box.setAlignment(Qt.AlignVCenter)
        self.frame.setLayout(self.frame.box)
        self.widgets = {}
        for text, value in zip(["x (Case)", "n (Case)", "x (Control)", "n (Control)"], [50, 100, 500, 1000]):
            frame = QFrame(self.frame)
            layout = QHBoxLayout(frame)
            frame.setLayout(layout)
            label = QLabel(frame)
            label.setText(text)
            label.setFixedWidth(100)
            spinbox = QSpinBox(frame)
            spinbox.setFixedWidth(200)
            spinbox.setSingleStep(1)
            spinbox.setMinimum(1)
            spinbox.setMaximum(100000)
            spinbox.setValue(value)
            layout.addWidget(label)
            layout.addWidget(spinbox)
            spinbox.valueChanged.connect(self.change_results)
            self.frame.box.addWidget(frame)
            self.widgets[text] = spinbox
        self.canvas = Canvas(self)
        self.box.addWidget(self.frame)
        self.box.addWidget(self.canvas)
        self.canvas.draw_binomial_distribution(
            ns=[100, 1000],
            ks=[50, 500],
            colors=["red", "green"]
        )

    def change_results(self):
        x_case = self.widgets["x (Case)"].value()
        n_case = self.widgets["n (Case)"].value()
        x_control = self.widgets["x (Control)"].value()
        n_control = self.widgets["n (Control)"].value()
        self.canvas.draw_binomial_distribution(
            ns=[n_case, n_control],
            ks=[x_case, x_control],
            colors=["red", "green"]
        )
