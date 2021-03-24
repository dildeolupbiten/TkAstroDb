# -*- coding: utf-8 -*-

from .about import About
from .database import Database
from .modules import tk, Thread
from .orb_factor import OrbFactor
from .constants import HOUSE_SYSTEMS, ASPECTS
from .selection import SingleSelection, MultipleSelection
from .calculations import find_observed_values, select_calculation
from .utilities import check_update, table_selection, merge_databases
from .export import export_link, export_lat_frequency, export_year_frequency


class Menu(tk.Menu):
    def __init__(self, icons, version, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.master["menu"] = self
        self.database_menu = tk.Menu(master=self, tearoff=False)
        self.spreadsheet_menu = tk.Menu(master=self, tearoff=False)
        self.calculations_menu = tk.Menu(master=self, tearoff=False)
        self.export_menu = tk.Menu(master=self, tearoff=False)
        self.options_menu = tk.Menu(master=self, tearoff=False)
        self.help_menu = tk.Menu(master=self, tearoff=False)
        self.add_cascade(label="Database", menu=self.database_menu)
        self.add_cascade(label="Spreadsheet", menu=self.spreadsheet_menu)
        self.add_cascade(label="Calculations", menu=self.calculations_menu)
        self.add_cascade(label="Export", menu=self.export_menu)
        self.add_cascade(label="Options", menu=self.options_menu)
        self.add_cascade(label="Help", menu=self.help_menu)
        self.database_menu.add_command(
            label="Open",
            command=lambda: Database(
                root=self.master,
                icons=icons
            )
        )
        self.database_menu.add_command(
            label="Merge And Convert",
            command=lambda: merge_databases(
                multiple_selection=MultipleSelection,
                icons=icons,
                widget=self.master
            )
        )
        self.spreadsheet_menu.add_command(
            label="Table Selection",
            command=lambda: table_selection(
                multiple_selection=MultipleSelection
            )
        )
        self.calculations_menu.add_command(
            label="Find Observed Values",
            command=lambda: find_observed_values(
                widget=self.master,
                icons=icons,
                menu=self.calculations_menu,
                version=version
            )
        )
        self.calculations_menu.add_command(
            label="Find Expected Values",
            command=lambda: Thread(
                target=lambda: select_calculation(
                    icons=icons,
                    calculation_type="expected",
                    input1="observed_values.xlsx",
                    input2="control_group.xlsx",
                    output="expected_values.xlsx",
                    widget=self.master
                ),
                daemon=True
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
                    output="chi-square.xlsx",
                    widget=self.master
                ),
                daemon=True
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
                    output="effect-size.xlsx",
                    widget=self.master
                ),
                daemon=True
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
                    output="cohens_d_effect.xlsx",
                    widget=self.master
                ),
                daemon=True
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
                    output="binomial_limit.xlsx",
                    widget=self.master
                ),
                daemon=True
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
            label="Category Selection",
            command=lambda: SingleSelection(
                title="Category Selection",
                catalogue=["Basic", "Advanced"]
            )
        )
        self.options_menu.add_command(
            label="Zodiac",
            command=lambda: SingleSelection(
                title="Zodiac",
                catalogue=["Tropical", "Sidereal"]
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
            command=lambda: OrbFactor(
                title="Orb Factor",
                catalogue=ASPECTS,
                config_key="ORB FACTORS"
            )
        )
        self.options_menu.add_command(
            label="Midpoint Orb Factor",
            command=lambda: OrbFactor(
                title="Midpoint Orb Factor",
                catalogue=["orb-factor"],
                config_key="MIDPOINT ORB FACTOR"
            )
        )
        self.help_menu.add_command(
            label="About",
            command=lambda: About(version=version)
        )
        self.help_menu.add_command(
            label="Check for updates",
            command=lambda: check_update(icons=icons)
        )
