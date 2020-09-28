# -*- coding: utf-8 -*-

from .menu import Menu
from .modules import tk, logging
from .utilities import create_image_files, load_defaults

logging.basicConfig(
    format="- %(levelname)s - %(asctime)s - %(message)s",
    level=logging.INFO,
    datefmt="%d.%m.%Y %H:%M:%S"
)


def main():
    root = tk.Tk()
    root.title("TkAstroDb")
    root.geometry("800x600")
    root.resizable(width=False, height=False)
    icons = create_image_files(path="Icons")
    Menu(master=root, icons=icons)
    load_defaults()
    root.mainloop()
