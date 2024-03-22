# TkAstroDb

**TkAstroDb** is a Python program that uses [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) to conduct statistical studies in astrology. Because of the license conditions, [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) can not be shared with third party users. Therefore those who are interested in using this program with [Astro-Databank](https://www.astro.com/astro-databank/Main_Page), should contact with the webmaster of [Astrodienst](http://www.astro.com) to get a license.

## Availability

- Windows, Linux and macOS
- Python >= 3.8

## Requirements

```commandline
pip install -r requirements.txt
```

or

```commandline
pip3 install -r requirements.txt
```
## Usage

```commandline
python3 app.py
```

or

```commandline
python app.py
```

## News

### March 2024

- Changed the graphical user interface from [TkInter](https://docs.python.org/3/library/tk.html) to [PyQt5](https://doc.qt.io/qtforpython-5/).
- Added collapsible frames that can be activated/deactivated by buttons for the main features of the program which are `Analysis`, `Calculations` and `Comparison`.
- Removed selections and calculations for `Grand Cross`, `Kite`, `Mystic Rectangle`
- Added selections to calculate `z-index` and `significance values` based on significance level.
- Added a responsive user input for graphical comparison.
- Added an intuitive carousel to make selections more easily and analyze the astrological charts quickly.
- Added a table view to discover and filter the database using different columns.

## Licenses

**TkAstroDb** is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.
