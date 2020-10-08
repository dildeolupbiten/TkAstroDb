# -*- coding: utf-8 -*-

from .messagebox import MsgBox
from .treeview import Treeview
from .selection import SingleSelection
from .utilities import tbutton_command, check_all_command
from .modules import (
    os, tk, ET, json, open_new, logging, ConfigParser, Thread
)


class Database:
    def __init__(self, root, icons):
        self.mode = None
        self.icons = icons
        self.database = None
        self.category_dict = {}
        self.all_categories = {}
        self.category_names = []
        self.choose_operation(root=root)
        
    def load_database(self, root, filename):
        if filename.endswith(".xml"):
            self.load_adb(filename=filename)
        else:
            self.load_json(filename=filename)
        DatabaseFrame(
            master=root,
            database=self.database,
            all_categories=self.all_categories,
            category_names=self.category_names,
            icons=self.icons,
            mode=self.mode
        )

    def choose_operation(self, root):
        if not os.path.exists("Database"):
            os.makedirs("Database")
        if not os.listdir("Database"):
            return
        else:
            SingleSelection(
                title="Database",
                catalogue=[i for i in os.listdir("Database")]
            )
            config = ConfigParser()
            config.read("defaults.ini")
            filename = config["DATABASE"]["selected"]
            filename = os.path.join(".", "Database", filename)
            Thread(target=lambda: self.load_database(root, filename)).start()

    def load_adb(self, filename):
        self.mode = "adb"
        self.database = []
        self.category_dict = {}
        logging.info(f"Parsing {filename} file...")
        tree = ET.parse(filename)
        root = tree.getroot()
        for i in range(1000000):
            try:
                user_data = []
                for gender, roddenrating, bdata, adb_link, categories in \
                        zip(
                            root[i + 2][1].findall("gender"),
                            root[i + 2][1].findall("roddenrating"),
                            root[i + 2][1].findall("bdata"),
                            root[i + 2][2].findall("adb_link"),
                            root[i + 2][3].findall("categories")
                        ):
                    name = root[i + 2][1][0].text
                    sbdate_dmy = bdata[1].text
                    sbtime = bdata[2].text
                    jd_ut = bdata[2].get("jd_ut")
                    lat = bdata[3].get("slati")
                    lon = bdata[3].get("slong")
                    place = bdata[3].text
                    country = bdata[4].text
                    category = [
                        (
                            categories[j].get("cat_id"),
                            categories[j].text
                        )
                        for j in range(len(categories))
                    ]
                    for cate in category:
                        if cate[0] not in self.category_dict.keys():
                            self.category_dict[cate[0]] = cate[1]
                    user_data.append(int(root[i + 2].get("adb_id")))
                    user_data.append(name)
                    user_data.append(gender.text)
                    user_data.append(roddenrating.text)
                    user_data.append(sbdate_dmy)
                    user_data.append(sbtime)
                    user_data.append(jd_ut)
                    user_data.append(lat)
                    user_data.append(lon)
                    user_data.append(place)
                    user_data.append(country)
                    user_data.append(adb_link.text)
                    user_data.append(category)
                    if len(user_data) != 0:
                        self.database.append(user_data)
            except IndexError:
                break
        try:
            logging.info("Completed parsing.")
            logging.info(f"{len(self.database)} records are available.")
        except tk.TclError:
            return
        self.group_categories()

    def group_categories(self):
        logging.info(f"Started grouping categories.")
        self.all_categories = {}
        if self.mode == "adb":
            if self.database[0][-1].startswith("Type"):
                index = -3
            else:
                index = -1
        else:
            index = -3
        for record in self.database:
            for category in record[index]:
                if (category[0], category[1]) not in self.all_categories:
                    if category[1] is None:
                        pass
                    self.all_categories[(category[0], category[1])] = []
                self.all_categories[
                    (category[0], category[1])
                ].append(record)
        self.category_names = sorted(
            [i for i in self.category_dict.values() if i is not None]
        )
        logging.info(f"Completed grouping categories.")

    def load_json(self, filename):
        if filename == "./Database/None":
            return
        self.mode = "normal"
        logging.info(f"Parsing {filename} file...")
        with open(filename, encoding="utf-8") as file:
            self.database = json.load(file)
        self.category_dict = {}
        if not isinstance(self.database, dict):
            self.mode = "adb"
            for record in self.database:
                for cate in record[-3]:
                    if cate[0] not in self.category_dict:
                        self.category_dict[cate[0]] = cate[1]
        else:
            self.mode = "normal"
            count = 1
            for key, value in self.database.items():
                if not value["Categories"]:
                    cat = "None"
                    if cat not in self.category_dict.values():
                        self.category_dict[count] = cat
                        count += 1
                else:
                    for cat in value["Categories"]:
                        if cat not in self.category_dict.values():
                            self.category_dict[count] = cat
                            count += 1
            r = {v: k for k, v in self.category_dict.items()}
            db = []
            for key, value in self.database.items():
                if value["Categories"]:
                    cate = [[r[cat], cat] for cat in value["Categories"]]
                else:
                    cate = [[r["None"], "None"]]
                info = [
                    v for k, v in value.items()
                    if k not in [
                        "Access Time", "Update Time", "Categories"
                    ]
                ]
                record = [key] + \
                    info + \
                    [cate] + \
                    [value["Access Time"], value["Update Time"]]
                db.append(record)
            self.database = db
        logging.info("Completed parsing.")
        logging.info(f"{len(self.database)} records are available.")
        self.group_categories()


class DatabaseFrame(tk.Frame):
    widgets = []

    def __init__(
            self,
            database,
            all_categories,
            category_names,
            icons,
            mode,
            *args,
            **kwargs
    ):
        super().__init__(*args, **kwargs)
        for i in self.widgets:
            i.destroy()
        self.pack()
        self.database = database
        self.all_categories = all_categories
        self.category_names = category_names
        self.icons = icons
        self.mode = mode
        self.displayed_results = []
        self.selected_categories = []
        self.selected_ratings = []
        self.checkbuttons = {}
        self.treeview_menu = None
        self.entry_menu = None
        self.info_var = tk.StringVar()
        self.info_var.set("0")
        self.topframe = tk.Frame(master=self)
        self.topframe.pack()
        self.midframe = tk.Frame(master=self)
        self.midframe.pack()
        self.bottomframe = tk.Frame(master=self)
        self.bottomframe.pack()
        if self.mode == "adb":
            self.columns = [
                "Adb ID", "Name", "Gender", "Rodden Rating", "Date",
                "Hour", "Julian Date", "Latitude", "Longitude", "Place",
                "Country", "Adb Link", "Category"
            ]
        else:
            self.columns = [
                "No", "Name", "Gender", "Rodden Rating", "Date",
                "Hour", "Julian Date", "Latitude", "Longitude", "Type",
                "Wing", "URL", "Category", "Access Time", "Update Time"
            ]
        self.treeview = Treeview(
            master=self.midframe,
            columns=self.columns,
            height=5
        )
        self.treeview.bind(
            sequence="<Button-1>",
            func=lambda event: self.destroy_menu(self.treeview_menu)
        )
        self.treeview.bind(
            sequence="<Button-3>",
            func=lambda event: self.button_3_on_treeview(event=event)
        )
        self.treeview.bind(
            sequence="<Control-a>",
            func=lambda event: self.treeview.selection_set(
                self.treeview.get_children()
            )
        )
        self.entry_button_frame = tk.Frame(master=self.topframe)
        self.entry_button_frame.grid(row=0, column=0)
        self.search_label = tk.Label(
            master=self.entry_button_frame,
            text="Search A Record By Name: ",
            fg="red"
        )
        self.search_label.grid(row=0, column=0, padx=5, sticky="w", pady=5)
        self.search_entry = tk.Entry(master=self.entry_button_frame)
        self.search_entry.grid(row=0, column=1, padx=5, pady=5)
        self.search_entry.bind(
            sequence="<KeyRelease>",
            func=lambda event: self.search_func()
        )
        self.search_entry.bind(
            sequence="<Button-1>",
            func=lambda event: self.destroy_menu(self.entry_menu))
        self.search_entry.bind(
            sequence="<Button-3>",
            func=self.button_3_on_entry
        )
        self.search_entry.bind(
            sequence="<Control-KeyRelease-a>",
            func=lambda event: self.search_entry.select_range("0", "end")
        )
        self.found_record = tk.Label(master=self.entry_button_frame, text="")
        self.found_record.grid(row=1, column=0, padx=5, pady=5)
        self.add_button = tk.Button(master=self.entry_button_frame, text="Add")
        self.category_label = tk.Label(
            master=self.entry_button_frame,
            text="Categories:",
            fg="red"
        )
        self.category_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.rrating_label = tk.Label(
            master=self.entry_button_frame,
            text="Rodden Rating:",
            fg="red"
        )
        self.rrating_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.category_button = tk.Button(
            master=self.entry_button_frame,
            text="Select",
            command=self.select_categories
        )
        self.category_button.grid(row=2, column=1, padx=5, pady=5)
        self.rating_button = tk.Button(
            master=self.entry_button_frame,
            text="Select",
            command=self.select_ratings
        )
        self.rating_button.grid(row=3, column=1, padx=5, pady=5)
        self.create_checkbutton()
        self.display_button = tk.Button(
            master=self.topframe,
            text="Display Records",
            command=self.display_results
        )
        self.display_button.grid(row=10, column=0, columnspan=4, pady=10)
        self.total_msgbox_info = tk.Label(
            master=self.bottomframe,
            text="Total = "
        )
        self.total_msgbox_info.grid(row=0, column=0)
        self.msgbox_info = tk.Label(
            master=self.bottomframe,
            textvariable=self.info_var
        )
        self.msgbox_info.grid(row=0, column=1)
        self.widgets.append(self)

    def add_command(self, record):
        if record in self.displayed_results:
            pass
        else:
            num = len(self.treeview.get_children()) + 1
            self.treeview.insert("", num, values=[col for col in record])
            self.displayed_results.append(record)
            self.info_var.set(len(self.displayed_results))
        self.add_button.grid_forget()
        self.found_record.configure(text="")
        self.search_entry.delete("0", "end")

    def search_func(self):
        self.update()
        save_record = ""
        count = 0
        for record in self.database:
            if self.search_entry.get() == record[1]:
                index = self.database.index(record)
                count += 1
                self.found_record.configure(text=f"Record Found = {count}")
                self.add_button.grid(row=1, column=1, padx=5, pady=5)
                self.add_button.configure(
                    command=lambda: self.add_command(
                        record=self.database[index]
                    )
                )
                save_record += self.database[index][1]
        if save_record != self.search_entry.get() or \
                save_record == self.search_entry.get() == "":
            self.found_record.configure(text="")
            self.add_button.grid_forget()

    def create_checkbutton(self):
        check_frame = tk.Frame(master=self.topframe)
        check_frame.grid(row=0, column=2)
        if self.mode == "adb":
            names = (
                "event",
                "human",
                "male",
                "female",
                "North Hemisphere",
                "South Hemisphere"
             )
        else:
            names = (
                "male",
                "female",
                "North Hemisphere",
                "South Hemisphere"
            )
        for i, j in enumerate(names):
            var = tk.StringVar()
            var.set(value="0")
            checkbutton = tk.Checkbutton(
                master=check_frame,
                text=f"Do not display {j} charts.",
                variable=var)
            checkbutton.grid(row=i, column=2, columnspan=2, sticky="w")
            self.checkbuttons[j] = [var, checkbutton]

    def south_north_check(self, item):
        north = self.checkbuttons["North Hemisphere"][0]
        south = self.checkbuttons["South Hemisphere"][0]
        if north.get() == "1" and south.get() == "0":
            if isinstance(item[7], str):
                if "n" in item[7]:
                    pass
                else:
                    self.insert_to_treeview(item)
            else:
                if item[7] > 0:
                    pass
                else:
                    self.insert_to_treeview(item)
        elif north.get() == "0" and south.get() == "1":
            if isinstance(item[7], str):
                if "s" in item[7]:
                    pass
                else:
                    self.insert_to_treeview(item)
            else:
                if item[7] < 0:
                    pass
                else:
                    self.insert_to_treeview(item)
        elif north.get() == "0" and south.get() == "0":
            self.insert_to_treeview(item)
        elif north.get() == "1" and south.get() == "1":
            pass

    def male_female_check(self, item):
        male = self.checkbuttons["male"][0]
        female = self.checkbuttons["female"][0]
        if male.get() == "1" and female.get() == "0":
            if item[2] == "M":
                pass
            else:
                self.south_north_check(item)
        elif male.get() == "0" and female.get() == "1":
            if item[2] == "F":
                pass
            else:
                self.south_north_check(item)
        elif male.get() == "0" and female.get() == "0":
            self.south_north_check(item)
        elif male.get() == "1" and female.get() == "1":
            if item[2] == "F" or item[2] == "M":
                pass
            else:
                self.south_north_check(item)

    def insert_to_treeview(self, item):
        num = len(self.treeview.get_children()) + 1
        self.treeview.insert("", num, values=[col for col in item])
        self.info_var.set(len(self.displayed_results))
        self.displayed_results.append(item)
        self.update()

    def display_results(self):
        self.treeview.delete(*self.treeview.get_children())
        self.displayed_results = []
        if self.mode == "adb":
            event = self.checkbuttons["event"][0]
            human = self.checkbuttons["human"][0]
        else:
            event = tk.StringVar()
            event.set("0")
            human = tk.StringVar()
            human.set("0")
        try:
            for key, value in self.all_categories.items():
                if key[1] in self.selected_categories:
                    for item in value:
                        if item[3] in self.selected_ratings:
                            if item in self.displayed_results:
                                pass
                            else:
                                if (
                                        event.get() == "0"
                                        and
                                        human.get() == "0"
                                ):
                                    if item[0] == 3546 or item[0] == 68092:
                                        pass
                                    else:
                                        self.male_female_check(item)
                                elif (
                                        event.get() == "1"
                                        and
                                        human.get() == "0"
                                ):
                                    if item[2] == "N/A":
                                        pass
                                    elif item[0] == 3546:
                                        pass
                                    else:
                                        self.male_female_check(item)
                                elif (
                                        event.get() == "0"
                                        and
                                        human.get() == "1"
                                ):
                                    if item[2] != "N/A" or item[0] == 68092:
                                        pass
                                    else:
                                        self.south_north_check(item)
                                elif (
                                        event.get() == "1"
                                        and
                                        human.get() == "1"
                                ):
                                    pass
        except tk.TclError:
            return
        self.info_var.set(len(self.displayed_results))
        self.update()
        if len(self.displayed_results) == 0:
            MsgBox(
                title="Info",
                message="No record is inserted.",
                icons=self.icons,
                level="info"
            )
        elif len(self.displayed_results) == 1:
            MsgBox(
                title="Info",
                message="1 record is inserted.",
                icons=self.icons,
                level="info"
            )
        else:
            MsgBox(
                title="Display Records",
                message=f"{len(self.displayed_results)} "
                        f"records are inserted.",
                icons=self.icons,
                level="info"
            )
        self.update()

    def select_ratings(self):
        self.selected_ratings = []
        toplevel = tk.Toplevel()
        toplevel.geometry("300x250")
        toplevel.resizable(width=False, height=False)
        toplevel.title("Select Rodden Ratings")
        toplevel.update()
        rating_frame = tk.Frame(master=toplevel)
        rating_frame.pack(side="top")
        button_frame = tk.Frame(master=toplevel)
        button_frame.pack(side="bottom")
        tbutton = tk.Button(master=button_frame, text="Apply")
        tbutton.pack()
        cvar_list = []
        checkbutton_list = []
        check_all = tk.BooleanVar()
        check_uncheck = tk.Checkbutton(
            master=rating_frame,
            text="Check/Uncheck All",
            variable=check_all
        )
        check_all.set(False)
        check_uncheck.grid(row=0, column=0, sticky="nw")
        for num, c in enumerate(
                ["AA", "A", "B", "C", "DD", "X", "XX", "AX"],
                1
        ):
            self.update()
            rating = c
            cvar = tk.BooleanVar()
            cvar_list.append([cvar, rating])
            checkbutton = tk.Checkbutton(
                master=rating_frame,
                text=rating,
                variable=cvar
            )
            checkbutton_list.append(checkbutton)
            cvar.set(False)
            checkbutton.grid(row=num, column=0, sticky="nw")
        tbutton.configure(
            command=lambda: tbutton_command(
                cvar_list,
                toplevel,
                self.selected_ratings
            )
        )
        check_uncheck.configure(
            command=lambda: check_all_command(
                check_all,
                cvar_list,
                checkbutton_list
            )
        )

    def select_categories(self):
        self.selected_categories = []
        toplevel = tk.Toplevel()
        toplevel.title("Select Categories")
        toplevel.resizable(width=False, height=False)
        toplevel.update()
        canvas_frame = tk.Frame(master=toplevel)
        canvas_frame.pack(side="top")
        button_frame = tk.Frame(master=toplevel)
        button_frame.pack(side="bottom")
        tcanvas = tk.Canvas(master=canvas_frame)
        tframe = tk.Frame(master=tcanvas)
        tscrollbar = tk.Scrollbar(
            master=canvas_frame,
            orient="vertical",
            command=tcanvas.yview
        )
        tcanvas.configure(yscrollcommand=tscrollbar.set)
        tscrollbar.pack(side="right", fill="y")
        tcanvas.pack()
        tcanvas.create_window((4, 4), window=tframe, anchor="nw")
        tframe.bind(
            "<Configure>",
            lambda event: tcanvas.configure(
                scrollregion=tcanvas.bbox("all")
            )
        )
        tbutton = tk.Button(master=button_frame, text="Apply")
        tbutton.pack()
        cvar_list = []
        checkbutton_list = []
        check_all = tk.BooleanVar()
        check_uncheck_ = tk.Checkbutton(
            master=tframe,
            text="Check/Uncheck All",
            variable=check_all
        )
        check_all.set(False)
        check_uncheck_.grid(row=0, column=0, sticky="nw")
        for num, category in enumerate(self.category_names, 1):
            try:
                self.update()
                cvar = tk.BooleanVar()
                cvar_list.append([cvar, category])
                checkbutton = tk.Checkbutton(
                    master=tframe,
                    text=category,
                    variable=cvar
                )
                checkbutton_list.append(checkbutton)
                cvar.set(False)
                checkbutton.grid(row=num, column=0, sticky="nw")
            except tk.TclError:
                return
        tbutton.configure(
            command=lambda: tbutton_command(
                cvar_list,
                toplevel,
                self.selected_categories
            )
        )
        check_uncheck_.configure(
            command=lambda: check_all_command(
                check_all,
                cvar_list,
                checkbutton_list
            )
        )
        self.update()

    def button_3_remove(self):
        selected = self.treeview.selection()
        if not selected:
            pass
        else:
            for i in selected:
                for j in self.displayed_results:
                    if j[0] == self.treeview.item(i)["values"][0]:
                        self.displayed_results.remove(j)
                self.treeview.delete(i)
            self.info_var.set(len(self.displayed_results))

    def button_3_open_url(self):
        selected = self.treeview.selection()
        if not selected:
            pass
        else:
            try:
                values = self.treeview.item(selected)["values"]
            except tk.TclError:
                return
            if values[11] != "None":
                open_new(values[11])

    def button_3_on_treeview(self, event):
        self.destroy_menu(self.treeview_menu)
        self.treeview_menu = tk.Menu(master=None, tearoff=False)
        if self.mode == "adb":
            label = "Open ADB Page"
        else:
            label = "Open Wikipedia Page"
        self.treeview_menu.add_command(
            label=label,
            command=self.button_3_open_url
        )
        self.treeview_menu.add_command(
            label="Remove",
            command=self.button_3_remove
        )
        self.treeview_menu.post(event.x_root, event.y_root)

    def button_3_on_entry(self, event):
        self.destroy_menu(self.entry_menu)
        self.entry_menu = tk.Menu(master=None, tearoff=False)
        self.entry_menu.add_command(
            label="Copy",
            command=lambda: self.focus_get().event_generate('<<Copy>>'))
        self.entry_menu.add_command(
            label="Cut",
            command=lambda: self.focus_get().event_generate('<<Cut>>'))
        self.entry_menu.add_command(
            label="Paste",
            command=lambda: self.focus_get().event_generate('<<Paste>>'))
        self.entry_menu.add_command(
            label="Remove",
            command=lambda: self.focus_get().event_generate('<<Clear>>'))
        self.entry_menu.add_command(
            label="Select All",
            command=lambda: self.focus_get().event_generate('<<SelectAll>>'))
        self.entry_menu.post(event.x_root, event.y_root)

    @staticmethod
    def destroy_menu(menu):
        if menu:
            menu.destroy()

    def select_all_items(self):
        self.treeview.selection_set(self.treeview.get_children())
