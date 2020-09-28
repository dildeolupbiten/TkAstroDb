# -*- coding: utf-8 -*-

from .about import About
from .database import Database
from .modules import tk, Thread
from .orb_factor import OrbFactor
from .utilities import check_update
from .constants import HOUSE_SYSTEMS
from .selection import SingleSelection
from .calculations import find_observed_values, select_calculation
from .export import export_link, export_lat_frequency, export_year_frequency


class Menu(tk.Menu):
    def __init__(self, icons, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.master["menu"] = self
        self.add_command(
            label="Database",
            command=lambda: Database(
                root=self.master,
                icons=icons
            )
        )
        self.calculations_menu = tk.Menu(master=self, tearoff=False)
        self.export_menu = tk.Menu(master=self, tearoff=False)
        self.options_menu = tk.Menu(master=self, tearoff=False)
        self.help_menu = tk.Menu(master=self, tearoff=False)
        self.add_cascade(label="Calculations", menu=self.calculations_menu)
        self.add_cascade(label="Export", menu=self.export_menu)
        self.add_cascade(label="Options", menu=self.options_menu)
        self.add_cascade(label="Help", menu=self.help_menu)
        self.calculations_menu.add_command(
            label="Find Observed Values",
            command=lambda: Thread(
                target=lambda: find_observed_values(
                    widget=self.master,
                    icons=icons
                )
            ).start()
        )
        self.calculations_menu.add_command(
            label="Find Expected Values",
            command=lambda: Thread(
                target=lambda: select_calculation(
                    icons=icons,
                    calculation_type="expected",
                    input1="observed_values.xlsx",
                    input2="control_group.xlsx",
                    output="expected_values.xlsx"
                )
            ).start()
        )
        self.calculations_menu.add_command(
            label="Find Chi-Square Values",
            command=lambda: Thread(
                target=lambda: select_calculation(
                    icons=icons,
                    calculation_type="chi-square",
                    input1="observed_values.xlsx",
                    input2="expected_values.xlsx",
                    output="chi-square.xlsx"
                )
            ).start()
        )
        self.calculations_menu.add_command(
            label="Find Effect Size Values",
            command=lambda: Thread(
                target=lambda: select_calculation(
                    icons=icons,
                    calculation_type="effect-size",
                    input1="observed_values.xlsx",
                    input2="expected_values.xlsx",
                    output="effect-size.xlsx"
                )
            ).start()
        )
        self.calculations_menu.add_command(
            label="Find Cohen's D Effect Size Values",
            command=lambda: Thread(
                target=lambda: select_calculation(
                    icons=icons,
                    calculation_type="cohen's d",
                    input1="observed_values.xlsx",
                    input2="expected_values.xlsx",
                    output="cohens_d_effect.xlsx"
                )
            ).start()
        )
        self.calculations_menu.add_command(
            label="Find Binomial Limit Values",
            command=lambda: Thread(
                target=lambda: select_calculation(
                    icons=icons,
                    calculation_type="binomial limit",
                    input1="observed_values.xlsx",
                    input2="control_group.xlsx",
                    output="binomial_limit.xlsx"
                )
            ).start()
        )
        self.export_menu.add_command(
            label="Links",
            command=lambda: export_link(
                widget=self.master,
                icons=icons
            )
        )
        self.export_menu.add_command(
            label="Latitude Frequency",
            command=lambda: export_lat_frequency(
                widget=self.master,
                icons=icons
            )
        )
        self.export_menu.add_command(
            label="Year Frequency",
            command=lambda: export_year_frequency(
                widget=self.master,
                icons=icons
            )
        )
        self.options_menu.add_command(
            label="House System",
            command=lambda: SingleSelection(
                title="House System",
                catalogue=HOUSE_SYSTEMS
            )
        )
        self.options_menu.add_command(
            label="Method",
            command=lambda: SingleSelection(
                title="Method",
                catalogue=["Subcategory", "Independent"]
            )
        )
        self.options_menu.add_command(
            label="Orb Factor",
            command=OrbFactor
        )
        self.help_menu.add_command(
            label="About",
            command=About
        )
        self.help_menu.add_command(
            label="Check for Updates",
            command=lambda: check_update(icons=icons)
        )
