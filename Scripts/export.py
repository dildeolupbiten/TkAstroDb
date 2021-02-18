# -*- coding: utf-8 -*-

from .modules import tk
from .messagebox import MsgBox
from .utilities import dd_to_dms, convert_coordinates


def export_link(widget, icons):
    displayed_results = []
    for i in widget.winfo_children():
        if hasattr(i, "included"):
            displayed_results += i.displayed_results
    if not displayed_results:
        MsgBox(
            title="Warning",
            level="warning",
            icons=icons,
            message="Please select and display records."
        )
        return
    with open(file="links.txt", mode="w", encoding="utf-8") as f:
        for i, j in enumerate(displayed_results):
            if len(j) == 13:
                url = j[-2]
            else:
                url = j[-4]
            f.write(f"{i + 1}. {url}\n")
    MsgBox(
        title="Info",
        level="info",
        message=f"{len(displayed_results)} links were exported.",
        icons=icons
    )


def export_lat_frequency(widget, icons):
    displayed_results = []
    for i in widget.winfo_children():
        if hasattr(i, "included"):
            displayed_results += i.displayed_results
    if not displayed_results:
        MsgBox(
            title="Warning",
            level="warning",
            icons=icons,
            message="Please select and display records."
        )
        return
    latitude_freq_north = {
        f"{i}\u00b0 - {i + 1}\u00b0": [] for i in range(90)
    }
    latitude_freq_south = {
        f"{-i}\u00b0 - {-i - 1}\u00b0": [] for i in range(90)
    }
    latitudes = []
    for item in displayed_results:
        latitudes.append(convert_coordinates(item[7]))
    for i in latitudes:
        for j in range(90):
            if j <= i < j + 1:
                latitude_freq_north[
                    f"{j}\u00b0 - {j + 1}\u00b0"].append(i)
            elif -j - 1 <= i < -j:
                latitude_freq_south[
                    f"{-j}\u00b0 - {-j - 1}\u00b0"].append(i)
    edit_latitude_freq_north = {
        keys: len(values)
        for keys, values in latitude_freq_north.items()
        if len(values) != 0
    }
    edit_latitude_freq_south = {
        keys: len(values)
        for keys, values in latitude_freq_south.items()
        if len(values) != 0
    }
    with open(
        file="latitude-frequency.txt",
        mode="w",
        encoding="utf-8"
    ) as f:
        f.write("Latitude Intervals\n\n")
        for i, j in edit_latitude_freq_south.items():
            f.write(f"{i} = {j}\n")
        for i, j in edit_latitude_freq_north.items():
            f.write(f"{i} = {j}\n")
        f.write(f"\nMean Latitude = "
                f"{dd_to_dms(sum(latitudes) / len(latitudes))}\n")
        f.write(f"\nTotal = {len(displayed_results)}")
        MsgBox(
            title="Info",
            message=f"{len(displayed_results)} records were exported.",
            icons=icons,
            level="info"
        )


def export_year_frequency(widget, icons):
    displayed_results = []
    for i in widget.winfo_children():
        if hasattr(i, "included"):
            displayed_results += i.displayed_results
    if not displayed_results:
        MsgBox(
            title="Warning",
            level="warning",
            icons=icons,
            message="Please select and display records."
        )
        return
    toplevel = tk.Toplevel()
    toplevel.title("Year Frequency")
    toplevel.geometry("200x100")
    toplevel.resizable(width=False, height=False)
    frame = tk.Frame(master=toplevel)
    frame.pack()
    date_entries = []
    freq_frmt = [0, 2000, 100]
    years = [int(i[4].split(" ")[2]) for i in displayed_results]
    if len(years) != 0:
        freq_frmt[0], freq_frmt[1], freq_frmt[2] = \
            min(years), max(years), 100
    for i, j in enumerate(("Minimum", "Maximum", "Step")):
        date_label = tk.Label(master=frame, text=j)
        date_label.grid(row=i, column=0, sticky="w")
        date_entry = tk.Entry(master=frame, width=5)
        date_entry.grid(row=i, column=1, sticky="w")
        date_entry.insert("1", f"{freq_frmt[i]}")
        date_entries.append(date_entry)
    apply_button = tk.Button(
        master=frame,
        text="Apply",
        command=lambda: year_frequency_command(
            toplevel=toplevel,
            date_entries=date_entries,
            years=years,
            freq_frmt=freq_frmt,
            displayed_results=displayed_results,
            icons=icons
        )
    )
    apply_button.grid(row=3, column=0, columnspan=3)


def year_frequency_command(
        toplevel,
        date_entries,
        years,
        freq_frmt,
        displayed_results,
        icons
):
    min_, max_, step_ = date_entries[:]
    min_, max_, step_ = int(min_.get()), int(max_.get()), \
        int(step_.get())
    freq_frmt[0], freq_frmt[1], freq_frmt[2] = min_, max_, step_
    with open("year-frequency.txt", "w", encoding="utf-8") as f:
        year_dict = {}
        count = 0
        for i in range(min_, max_ + 1, step_):
            year_dict[
                (min_ + (count * step_),
                 min_ + (count * step_) + step_)
            ] = []
            count += 1
        for i in years:
            for keys, values in year_dict.items():
                if keys[0] <= i < keys[1]:
                    year_dict[keys[0], keys[1]] += i,
        for keys, values in year_dict.items():
            f.write(f"{keys} = {len(values)}\n")
        f.write(f"Total = {len(displayed_results)}")
        toplevel.destroy()
        MsgBox(
            title="Info",
            message=f"{len(displayed_results)} records were exported.",
            icons=icons,
            level="info"
        )
