# -*- coding: utf-8 -*-

from .entry import EntryFrame
from .treeview import Treeview
from .search import SearchFrame
from .selection import SingleSelection
from .messagebox import MsgBox, ChoiceBox
from .utilities import (
    tbutton_command, check_all_command, load_database, from_xml
)
from .modules import (
    os, tk, ttk, open_new, logging, ConfigParser, Thread
)


class Database:
    def __init__(self, root, icons):
        self.icons = icons
        self.database = None
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
            category_names=self.category_names,
            icons=self.icons,
        )

    def choose_operation(self, root):
        if not os.path.exists("Database"):
            os.makedirs("Database")
        if not os.listdir("Database"):
            return
        else:
            selection = SingleSelection(
                title="Database",
                catalogue=[
                    i for i in os.listdir("Database")
                    if i.endswith("xml") or i.endswith("json")
                ]
            )
            if selection.done:
                config = ConfigParser()
                config.read("defaults.ini")
                filename = config["DATABASE"]["selected"]
                filename = os.path.join(".", "Database", filename)
                Thread(
                    target=lambda: self.load_database(root, filename),
                    daemon=True
                ).start()
            else:
                return

    def load_adb(self, filename):
        logging.info(f"Parsing {filename} file...")
        self.database, self.category_names = from_xml(filename)
        try:
            logging.info("Completed parsing.")
            logging.info(f"{len(self.database)} records are available.")
        except tk.TclError:
            return

    def load_json(self, filename):
        if filename == "./Database/None":
            return
        logging.info(f"Parsing {filename} file...")
        self.database = load_database(filename=filename)
        self.category_names = []
        if len(self.database[0]) == 13:
            index = -1
        else:
            index = -3
        for record in self.database:
            for cate in record[index]:
                if cate[1] and cate[1] not in self.category_names:
                    self.category_names.append(cate[1])
        self.category_names = sorted(self.category_names)
        try:
            logging.info("Completed parsing.")
            logging.info(f"{len(self.database)} records are available.")
        except tk.TclError:
            return


class DatabaseFrame(tk.Frame):
    widgets = []

    def __init__(
            self,
            database,
            category_names,
            icons,
            *args,
            **kwargs
    ):
        super().__init__(*args, **kwargs)
        for i in self.widgets:
            i.destroy()
        self.pack()
        self.database = database
        self.category_names = category_names
        self.icons = icons
        self.displayed_results = []
        self.included = []
        self.ignored = []
        self.selected_ratings = []
        self.found_categories = []
        self.checkbuttons = {}
        self.category_menu = None
        self.treeview_menu = None
        self.entry_menu = None
        self.pressed_return = 0
        self.start = ""
        self.end = ""
        self.info_var = tk.StringVar()
        self.info_var.set("0")
        self.topframe = tk.Frame(master=self)
        self.topframe.pack()
        self.midframe = tk.Frame(master=self)
        self.midframe.pack()
        self.bottomframe = tk.Frame(master=self)
        self.bottomframe.pack()
        self.columns = [
            "Adb ID", "Name", "Gender", "Rodden Rating", "Date",
            "Hour", "Julian Date", "Latitude", "Longitude", "Place",
            "Country", "Adb Link", "Category"
        ]
        if len(database[0]) == 15:
            self.columns.extend(["Type", "Wing"])
        self.treeview = Treeview(
            master=self.midframe,
            columns=self.columns,
            height=5,
            x_scrollbar=True
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
        self.search = SearchFrame(
            master=self.entry_button_frame,
            treeview=self.treeview,
            database=self.database,
            info_var=self.info_var
        )
        self.category_label = tk.Label(
            master=self.entry_button_frame,
            text="Categories:",
            fg="red"
        )
        self.category_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.rrating_label = tk.Label(
            master=self.entry_button_frame,
            text="Rodden Rating:",
            fg="red"
        )
        self.rrating_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.category_button = tk.Button(
            master=self.entry_button_frame,
            text="Select",
            command=self.select_categories
        )
        self.category_button.grid(row=1, column=1, padx=5, pady=5)
        self.rating_button = tk.Button(
            master=self.entry_button_frame,
            text="Select",
            command=self.select_ratings
        )
        self.rating_button.grid(row=2, column=1, padx=5, pady=5)
        self.create_checkbutton()
        self.entry_frame = EntryFrame(
            master=self.topframe,
            texts=["From", "To"],
            title="Select Year Range"
        )
        self.entry_frame.grid(row=10, column=0, columnspan=4, pady=10)
        self.button_frame = tk.Frame(master=self.topframe)
        self.button_frame.grid(row=11, column=0, columnspan=4, pady=10)
        self.get_button = tk.Button(
            master=self.button_frame,
            text="Get Records",
            command=self.get_records,
            width=12
        )
        self.get_button.pack(side="left")
        self.display_button = tk.Button(
            master=self.button_frame,
            text="Display Records",
            command=self.display_results,
            width=12
        )
        self.display_button.pack(side="left")
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

    def create_checkbutton(self):
        check_frame = tk.Frame(master=self.entry_button_frame)
        check_frame.grid(row=1, column=2, pady=10, rowspan=2)
        names = (
            "event",
            "human",
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
                text=f"Do not select {j}.",
                variable=var)
            checkbutton.grid(row=i, column=2, columnspan=2, sticky="w")
            self.checkbuttons[j] = [var, checkbutton]
            
    def get_records(self, display=False):
        self.displayed_results = []
        event = self.checkbuttons["event"][0]
        human = self.checkbuttons["human"][0]
        male = self.checkbuttons["male"][0]
        female = self.checkbuttons["female"][0]
        north = self.checkbuttons["North Hemisphere"][0]
        south = self.checkbuttons["South Hemisphere"][0]
        if len(self.database[0]) == 13:
            index = -1
        else:
            index = -3
        self.start = self.entry_frame.widgets["From"].get()
        self.end = self.entry_frame.widgets["To"].get()
        for record in self.database:
            if record[3] not in self.selected_ratings:
                continue
            if event.get() == "1" and record[2] == "N/A":
                continue
            if human.get() == "1" and record[2] in ["F", "M"]:
                continue
            if male.get() == "1" and record[2] == "M":
                continue
            if female.get() == "1" and record[2] == "F":
                continue
            if (
                isinstance(record[7], str)
                and
                north.get() == "1"
                and
                "n" in record[7]
            ):
                continue
            if (
                isinstance(record[7], float)
                and
                north.get() == "1"
                and
                record[7] > 0
            ):
                continue
            if (
                isinstance(record[7], str)
                and
                south.get() == "1"
                and
                "s" in record[7]
            ):
                continue
            if (
                isinstance(record[7], float)
                and
                south.get() == "1"
                and
                record[7] < 0
            ):
                continue
            if record[0] in [3546, 68092]:
                continue
            if not any(
                category[1] in self.included
                for category in record[index]
            ):
                continue
            if any(
                category[1] in self.ignored
                for category in record[index]
            ):
                continue
            year = int(record[4].split()[2])
            if (
                self.start 
                and 
                self.end 
                and not int(self.start) <= year <= int(self.end)
            ):
                continue
            self.displayed_results += [record]
        if not display:
            self.inform_user(message="gotten")

    def display_results(self):
        if self.treeview.get_children():
            msg = "There are records in treeview.\n" \
                  "Are you sure you want to insert \nthe records?\n" \
                  "If you press 'OK', \nthe records in treeview would be\n" \
                  "deleted."
            choicebox = ChoiceBox(
                title="Warning",
                level="warning",
                message=msg,
                icons=self.icons,
                width=400,
                height=200,
            )
            if choicebox.choice:
                pass
            else:
                return
        self.get_records(display=True)
        self.treeview.delete(*self.treeview.get_children())
        for index, i in enumerate(self.displayed_results):
            try:
                self.treeview.insert(
                    parent="",
                    index=index,
                    values=i
                )                
                self.info_var.set(index + 1)
                self.update()
            except tk.TclError:
                return
        self.inform_user()
                
    def inform_user(self, message="inserted"):
        self.update()
        if len(self.displayed_results) == 0:
            MsgBox(
                title="Info",
                message=f"No record is {message}.",
                icons=self.icons,
                level="info"
            )
        elif len(self.displayed_results) == 1:
            MsgBox(
                title="Info",
                message=f"1 record is {message}.",
                icons=self.icons,
                level="info"
            )
        else:
            MsgBox(
                title="Display Records",
                message=f"{len(self.displayed_results)} "
                        f"records are {message}.",
                icons=self.icons,
                level="info"
            )

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

    def category_widgets(self, master, container):
        search_label = tk.Label(
            master=master,
            text="Search a category",
            font="Default 9 bold"
        )
        search_label.pack()
        entry = ttk.Entry(master=master)
        entry.pack()
        entry.bind(
            sequence="<KeyRelease>",
            func=lambda event: self.search_category(
                event=event,
                treeview=treeview
            )
        )
        entry.bind(
            sequence="<Return>",
            func=lambda event: self.goto_next_category(
                treeview=treeview,
                event=event
            )
        )
        frame = tk.Frame(master=master)
        frame.pack()
        treeview = Treeview(
            master=frame,
            columns=["Categories"],
            width=400,
            anchor="w"
        )
        treeview.pack()
        for index, i in enumerate(self.category_names):
            treeview.insert(
                index=index,
                parent="",
                values=f"\"{i}\"",
                tag=index
            )
        var = tk.StringVar()
        var.set("Selected = 0")
        info_label = tk.Label(
            master=master,
            textvariable=var
        )
        info_label.pack()
        treeview.bind(
            sequence="<Button-3>",
            func=lambda event: self.button_3_on_cat_treeview(
                event=event,
                var=var,
                container=container
            )
        )
        treeview.bind(
            sequence="<Button-1>",
            func=lambda event: self.destroy_menu(menu=self.category_menu)
        )

    def select_categories(self):
        config = ConfigParser()
        config.read("defaults.ini")
        selection = config["CATEGORY SELECTION"]["selected"]
        if selection == "Basic":
            self.select_basic_categories()
        elif selection == "Advanced":
            self.select_advanced_categories()

    def select_basic_categories(self):
        self.included = []
        self.ignored = []
        included = []
        toplevel = tk.Toplevel()
        toplevel.title("Select Categories")
        toplevel.resizable(width=False, height=False)
        toplevel.update()
        self.category_widgets(
            master=toplevel,
            container=included
        )
        button = tk.Button(
            master=toplevel,
            text="Apply",
            command=lambda: self.apply_selection(
                included=included,
                ignored=[],
                master=toplevel
            )
        )
        button.pack(side="bottom")

    def select_advanced_categories(self):
        self.included = []
        self.ignored = []
        included = []
        ignored = []
        toplevel = tk.Toplevel()
        toplevel.title("Select Categories")
        toplevel.resizable(width=False, height=False)
        toplevel.update()
        main_frame = tk.Frame(master=toplevel)
        main_frame.pack()
        left_frame = tk.Frame(master=main_frame)
        left_frame.pack(side="left")
        left_label = tk.Label(
            master=left_frame,
            text="Include",
            font="Default 12 bold"
        )
        left_label.pack()
        right_frame = tk.Frame(master=main_frame)
        right_frame.pack(side="right")
        right_label = tk.Label(
            master=right_frame,
            text="Ignore",
            font="Default 12 bold"
        )
        right_label.pack()
        self.category_widgets(
            master=left_frame,
            container=included,
        )
        self.category_widgets(
            master=right_frame,
            container=ignored,
        )
        button = tk.Button(
            master=toplevel,
            text="Apply",
            command=lambda: self.apply_selection(
                included=included,
                ignored=ignored,
                master=toplevel
            )
        )
        button.pack(side="bottom")

    @staticmethod
    def change_status_of_selected(var, container, widget, mode):
        selection = widget.selection()
        for i in selection:
            item = widget.item(i)["values"][0]
            tag = widget.item(i)["tags"][0]
            if mode == "add" and item not in container:
                widget.tag_configure(tag, foreground="red")
                container.append(item)
            elif mode == "remove" and item in container:
                widget.tag_configure(tag, foreground="black")
                container.remove(item)
        var.set(f"Selected = {len(container)}")

    def button_3_on_cat_treeview(self, event, var, container):
        self.destroy_menu(self.category_menu)
        self.category_menu = tk.Menu(master=None, tearoff=False)
        self.category_menu.add_command(
            label="Add",
            command=lambda: self.change_status_of_selected(
                var=var,
                container=container,
                widget=event.widget,
                mode="add"
            )
        )
        self.category_menu.add_command(
            label="Remove",
            command=lambda: self.change_status_of_selected(
                var=var,
                container=container,
                widget=event.widget,
                mode="remove"
            )
        )
        self.category_menu.post(event.x_root, event.y_root)

    def goto_next_category(self, event, treeview):
        if event.widget.get() and self.found_categories:
            if self.pressed_return + 1 == len(self.found_categories):
                self.pressed_return = 0
            else:
                self.pressed_return += 1
            key = list(self.found_categories)[self.pressed_return]
            value = self.found_categories[key]
            treeview.yview_moveto(
                key / len(treeview.get_children())
            )
            treeview.selection_set(value)

    def search_category(self, event, treeview):
        if event.widget.get().lower() and event.keysym != "Return":
            self.pressed_return = 0
            self.found_categories = {
                i: j for i, j in enumerate(treeview.get_children())
                if (
                    event.widget.get().lower()
                    in
                    treeview.item(j)["values"][0].lower()
                )
            }
            for i, j in self.found_categories.items():
                treeview.yview_moveto(
                    i / len(treeview.get_children())
                )
                treeview.selection_set(j)
                break

    def apply_selection(self, included, ignored, master):
        self.included = included
        self.ignored = ignored
        master.destroy()

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
        self.treeview_menu.add_command(
            label="Open ADB Page",
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
