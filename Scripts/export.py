# -*- coding: utf-8 -*-

from .messagebox import MsgBox
from .utilities import convert_coordinates, edit_coordinate


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
    with open("links.txt", "w", encoding="utf-8") as f:
        for i, j in enumerate(displayed_results):
            if len(j) == 13:
                url = j[-2]
            else:
                url = j[-4]
            f.write(f"{url}\n")
    MsgBox(
        title="Info",
        level="info",
        message=f"{len(displayed_results)} links were exported.",
        icons=icons
    )


def export_dist(widget, icons, dist):
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
    d = {}
    for item in displayed_results:
        if dist == "latitude":
            var = int(convert_coordinates(item[7]))
        elif dist == "longitude":
            var = int(convert_coordinates(item[8]))
        else:
            var = int(item[4].split()[2])
        if var in d:
            d[var] += 1
        else:
            d[var] = 1
    with open(f"{dist}_distribution.csv", "w") as f:
        f.write(f"{dist},record\n")
        if dist == "year":
            for k, v in d.items():
                f.write(f"{k},{v}\n")
        elif dist in ["latitude", "longitude"]:
            for k, v in {_k: d[_k] for _k in sorted(d)}.items():
                f.write(f"\"{edit_coordinate(k, dist)}\",{v}\n")
        MsgBox(
            title="Info",
            message=f"{len(displayed_results)} records were exported.",
            icons=icons,
            level="info"
        )
