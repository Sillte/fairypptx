from pathlib import Path
import os
from setuptools import setup, find_packages

_this_folder = Path(__file__).resolve().parent


def _read_requirements(path):
    return [line for line in map(str.strip, Path(path).read_text().split("\n")) if line]


_install_requires = _read_requirements(_this_folder / "requirements.txt")

if __name__ == "__main__":
    os.chdir(_this_folder)  # Current folder is set to be this directory.
    setup(
        packages=find_packages(),
        install_requires=_install_requires,
    )
