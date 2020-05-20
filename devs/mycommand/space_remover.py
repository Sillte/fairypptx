import os
from pathlib import Path
from setuptools import Command
from distutils import errors

_this_folder = Path(__file__).resolve().parent

LINESEP = "\n"


def remove_trailings(path):
    """To prevent W291 trailing whitespace

    Note
    ----
    * This is in-place.
    * UTF8 is assumed.
    """
    path = Path(path)
    if not path.exists():
        raise ValueError("``path`` must exist.")

    if path.is_dir():
        for f in path.glob("**/*.py"):
            remove_trailings(f)
    else:
        lines = path.read_text(encoding="utf8").split(LINESEP)
        lines = [line.rstrip() for line in lines]
        output = LINESEP.join(lines)
        path.write_text(output, encoding="utf8")


class SpaceRemoverCommand(Command):
    description = "To prevent W291 trailing whitespace. "

    """ For specification, 
    refer to  https://dankeder.com/posts/adding-custom-commands-to-setup-py/
    """
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
        for f in self.root.glob("**/*.py"):
            remove_trailings(f)


def _test():
    original = """ Are you a werewolf or madman? 
It is nothing but a mistake of you to claim innocent me is werewolf...   
because I am a Hunter.
  
"""
    expected = """ Are you a werewolf or madman?
It is nothing but a mistake of you to claim innocent me is werewolf...
because I am a Hunter.

"""
    _test_path = _this_folder / "__remove_trailing.txt"
    _test_path.write_text(original)
    remove_trailings(_test_path)
    text = _test_path.read_text()
    assert expected == _test_path.read_text()
    os.unlink(_test_path)


def _compare(base, target):
    for a, b in zip(text.split(linesep), expected.split(LINESEP)):
        print(a.replace(" ", "*"), b.replace(" ", "*"))


if __name__ == "__main__":
    _test()
