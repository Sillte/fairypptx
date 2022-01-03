"""Handle registory.
In this context, registrory stores `json` or `pkl` data.

File Storage
--------------------
 `{Path.home()} /.fairypptx/registory`.

Key Specification
--------------------
* (category): Specify what kind of data is handled. 
* (key): the key  
         When the `target` is stored in File Storage, 
         `{key}.json` becomes the file name.
         Case incensitive.

Value Specification
----------------------
* (target)

* When File Storage is used, the target must be json-serializable.
* If you only use in-memory registory, then any object is valid,
but  is deriable if so. 


Normally, users are not intended to call below functions directly. 
For each class, ``register`` / ``fetch`` functions are prepared. 
"""

from pathlib import Path
from collections.abc import Mapping
from collections import defaultdict
import shutil
import subprocess
import weakref
import json
import pickle

__registory = defaultdict(dict)


def register(category, key, target, extension=None, *, disk=None):
    """Register `target` of the specific `category`  with `key`.
    Args:
        catetory: the subfolder name. 
        key (str): the stem of path.
        target (str): the registered target.
        extension: If specified, then the file is stored in `{key}{extension}`
                   Notice that this specification is related to `disk`.
        disk (bool or None): If `True`, file is stored,
                             If `False`, file is no stored, 
                             If `None`, boolean is determined whether extension is specified or not.
    Note
    --------------------------------------------------------
    If you would like to save `target` to disk, 
    you should do one of the followings.
        A. specify extension to either `.json` / `.pkl` and set `disk` to be `True` or None.
        B. setting `disk` to be `True`.
    Notice that A. is safer.   
    """
    global __registory
    if disk is None:
        disk = True if extension else False
    if disk is True:
        extension = _solve_extension(target, extension)
    else:
        extension = None
    # Internally, if extension is `None` here,  then disk is not stored.
    assert extension in {".pkl", ".json", None}

    __registory[category][key] = target
    if extension is not None:
        # Existent files are deleted.
        for path in _to_existent_paths(category, key):
            path.unlink()
        path = _to_path(category, key, extension)
        if extension == ".json":
            with open(path, "w", encoding="utf8") as fp:
                json.dump(target, fp, indent=4, ensure_ascii=False)
        elif extension == ".pkl":
            with open(path, "wb") as fp:
                pickle.dump(target, fp)
        else:
            raise NotImplementedError("Bug.")

def fetch(category, key, *, disk=True): 
    """Fetch target with (category, key).
    """
    global __registory

    if key in __registory[category]:
        return __registory[category][key]

    if disk is True:
        paths = _to_existent_paths(category, key)
        if 1 < len(paths):
            raise KeyError(f"Multiple files for {(key, category)} exist. It is a bug.")
        elif len(paths) == 0:
            raise KeyError(f"[`category`][`key`] = [{category}][{key}] is not existent in memory nor disk.")
        path = paths[0]
        extension = path.suffix
        if extension == ".json":
            with open(path, "r", encoding="utf8") as fp:
                target = json.load(fp)
        elif extension == ".pkl":
            with open(path, "rb") as fp:
                target = pickle.load(fp)
        __registory[category][key] = target
        return target
    raise KeyError(f"[`category`][`key`] = [{category}][{key}] is not existent in memory.")

def keys(category, disk=True):
    """Return the set of keys.
    """
    memory_keys = set(__registory[category].keys())
    if disk is False:
        return memory_keys
    else:
        folder = _registory_folder() / category
        folder_keys = set(p.stem for p in folder.glob("*.*"))
        return memory_keys | folder_keys


def clear(category=None, key=None, disk=False):
    """Clear the registory. 
    """
    global __registory
    if category is None:
        __registory = defaultdict(dict)
        if disk is True:
            folders = _registory_folder().glob("*/")
            if folder.exists():
                shutil.rmtree(folder)
    else:
        folder = _registory_folder() / category
        if key is None:
            __registory[category] = dict()
            if disk is True:
                if not folder.exists():
                    raise ValueError(f"`{category}` folder does not exist.")
                shutil.rmtree(folder)
        else:
            del __registory[category][key]
            if disk is True:
                paths = _to_existent_paths(category, key)
                for path in paths:
                    path.unlink()

def _to_stem(key):
    # For tuples, "$" is used for separator? 
    assert isinstance(key, str), "Current Limitation"
    assert "$" not in key, "Current Limitation"
    return str(key)


def _to_path(category, key, extension):
    folder = _registory_folder() / category
    folder.mkdir(exist_ok=True)
    stem = _to_stem(key)
    path = folder / f"{stem}{extension}"
    return path

def _to_existent_paths(category, key):
    folder = _registory_folder() / category
    stem = _to_stem(key)
    return list(folder.glob(f"{stem}.*/"))

def _registory_folder():
    folder =  Path.home() / ".fairypptx" / "registory"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _solve_extension(target, extension):
    if extension is not None:
        return extension

    # If json serializable, return `.json`.
    try:
        json.dumps(target, indent=4, ensure_ascii=False)
    except Exception:
        pass
    else:
        return  ".json"
    # The last resort.
    return ".pkl"

if __name__ == "__main__":
    _registory_folder()
    data = {"first": 1, "second":2}
    register("TestObject", "data_key", data, ".pkl", disk=False)
    clear()
    gained = fetch("TestObject", "data_key", disk=False)
    print(keys("TestObject"))
    print(gained)


