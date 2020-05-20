import os
import subprocess
from pathlib import Path
from setuptools import Command
from distutils import errors

from .space_remover import remove_trailings

_this_folder = Path(__file__).resolve().parent


class FormatViolation(Exception):
    pass


def apply_format(path, flake_output="flake8.txt"):
    """My Invidivual Code Formatter.

    """
    remove_trailings(path)
    result = subprocess.call(f"black {path}", shell=True)
    flake_output = Path(flake_output)
    if flake_output.exists():
        os.unlink(flake_output)
    ret = subprocess.call(f"flake8 {path} --output-file {flake_output} --tee", shell=True)
    if ret:
        raise FormatViolation("Failed at flake8.")


class FormatterCommand(Command):
    """My Individual Code Formatter.
    Just warp my ``bat`` commands as Command.
    """

    description = "My individual code formatter"

    user_options = [("root=", "r", "root folder.")]

    def initialize_options(self):
        self.root = None

    def finalize_options(self):
        if not self.root:
            raise errors.DistutilsArgError("``root`` option is required.")
        self.root = Path(self.root)
        if not self.root.exists():
            raise errors.DistutilsFileError(f"``{self.root}`` does not exist.")

    def run(self):
        try:
            apply_format(self.root)
        except FormatViolation as e:
            raise errors.DistutilsExecError("[Error at Formatter]", e)
