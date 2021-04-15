#!/usr/bin/python3.8
# -*- coding: utf-8 -*-

import os
import sys
import platform
import subprocess

installed_packages = [
    p.decode().split("==")[0]
    for p in subprocess.check_output(
        [sys.executable, "-m", "pip", "freeze"]
    ).split()
]

packages = [
    "numpy==1.19.2",
    "matplotlib",
    "scipy==1.4.1",
    "pandas==1.1.2",
    "XlsxWriter",
    "xlrd==1.2.0"
]


def call(p, n, f):
    os.system(
        f"{sys.executable} -m pip install {os.path.join(p, f[n])}"
    )
    

if os.name == "posix":
    packages += ["pyswisseph==2.0.0.post2"]
elif os.name == "nt":
    if "pyswisseph" not in installed_packages:
        path = os.path.join(".", "Eph", "Whl")
        files = [i for i in os.listdir(path) if "pyswisseph" in i]
        if sys.version_info.minor == 6:
            if platform.architecture()[0] == "32bit":
                call(p=path, n=0, f=files)
            elif platform.architecture()[0] == "64bit":
                call(p=path, n=1, f=files)
        elif sys.version_info.minor == 7:
            if platform.architecture()[0] == "32bit":
                call(p=path, n=2, f=files)
            elif platform.architecture()[0] == "64bit":
                call(p=path, n=3, f=files)
        elif sys.version_info.minor == 8:
            if platform.architecture()[0] == "32bit":
                call(p=path, n=4, f=files)
            elif platform.architecture()[0] == "64bit":
                call(p=path, n=5, f=files)
        elif sys.version_info.minor == 9:
            if platform.architecture()[0] == "32bit":
                call(p=path, n=6, f=files)
            elif platform.architecture()[0] == "64bit":
                call(p=path, n=7, f=files)

for package in packages:
    if package.startswith("pyswisseph"):
        if "pyswisseph" not in installed_packages:
            os.system(f"{sys.executable} -m pip install {package}")
    else:
        if package.split("==")[0] not in installed_packages:
            os.system(f"{sys.executable} -m pip install {package}")
           
if __name__ == "__main__":
    import Scripts
    Scripts.main()
