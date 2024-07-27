"""Handle registry.
In this context, registry stores `json` or `pkl` data.

File Storage
--------------------
 `{Path.home()} /.fairypptx/registry`.

Key Specification
--------------------
* (category): Specify what kind of data is handled. 
* (key): the key  
         When the `target` is stored in File Storage, 
         `{key}.json` becomes the file name.
         Case insensitive.

Value Specification
----------------------
* (target)

* When File Storage is used, the target must be json-serializable.
* If you only use in-memory registry, then any object is valid,
but  is desirable if so. 


Normally, users are not intended to call below functions directly. 
For each class, ``register`` / ``fetch`` functions are prepared. 


Additionally, this sub module provide the temporary file generation, 
Some Microsoft Object Models requires the existence of `file`. (e.g. , `FillFormat.UserPicture`).  
For these functions, it is expected to use this picture.
"""

from pathlib import Path
from contextlib import contextmanager
from collections.abc import Mapping
from collections import defaultdict
import shutil
import subprocess
import weakref
import json
import pickle

__registry = defaultdict(dict)


def register(category, key, target, extension=None, *, disk=None):
    """Register `target` of the specific `category`  with `key`.
    Args:
        category: the subfolder name. 
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
    global __registry
    if disk is None:
        disk = True if extension else False
    if disk is True:
        extension = _solve_extension(target, extension)
    else:
        extension = None
    # Internally, if extension is `None` here,  then disk is not stored.
    assert extension in {".pkl", ".json", None}

    __registry[category][key] = target
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
    global __registry

    if key in __registry[category]:
        return __registry[category][key]

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
        __registry[category][key] = target
        return target
    raise KeyError(f"[`category`][`key`] = [{category}][{key}] is not existent in memory.")

def keys(category, disk=True):
    """Return the set of keys.
    """
    memory_keys = set(__registry[category].keys())
    if disk is False:
        return memory_keys
    else:
        folder = _registry_folder() / category
        folder_keys = set(p.stem for p in folder.glob("*.*"))
        return memory_keys | folder_keys


def clear(category=None, key=None, disk=False):
    """Clear the registry. 
    """
    global __registry
    if category is None:
        __registry = defaultdict(dict)
        if disk is True:
            folders = _registry_folder().glob("*/")
            if folder.exists():
                shutil.rmtree(folder)
    else:
        folder = _registry_folder() / category
        if key is None:
            __registry[category] = dict()
            if disk is True:
                if not folder.exists():
                    raise ValueError(f"`{category}` folder does not exist.")
                shutil.rmtree(folder)
        else:
            del __registry[category][key]
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
    folder = _registry_folder() / category
    folder.mkdir(exist_ok=True)
    stem = _to_stem(key)
    path = folder / f"{stem}{extension}"
    return path

def _to_existent_paths(category, key):
    folder = _registry_folder() / category
    stem = _to_stem(key)
    return list(folder.glob(f"{stem}.*"))

def _registry_folder():
    folder =  Path.home() / ".fairypptx" / "registry"
    folder.mkdir(parents=True, exist_ok=True)
    return folder

REGISTRY_FOLDER = _registry_folder()   #  The folder into which files are generated.  

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


@contextmanager
def yield_temporary_path(memory=None, suffix: str=None):
    """Generate `temporary` file within this context.

    Args: 
        memory: the target object, it's used when you want to dump `memory` into `HDD`. 
        suffix: the suffix of the file. 
 
    This is intended to be used for calling a part of Microsoft Object Model function. 
    Hence, the type of `memory` is limited, I suppose.   

    This function assumes 2 scenes. 

    1. When you use `path` as the input of `Method of Object Model`.
        - For this usage, you should specify `memory` at call.
    2. When you use `path` as the output of `Method of Object Model`.
        - For this usage, you should specify `suffix` at call.

    Note:
        As of today (2022-01-04), the generation of `folder` is out of scope.   
    """
    import uuid
    import os
    from PIL import Image

    if memory is None and suffix is None:
        raise TypeError("Either of `memory` or `suffix` must be specified.")

    if memory is not None and suffix is not None:
        raise TypeError("Don't specify both of `memory` or `suffix`.")


    temporary_folder = _registry_folder()  / "__temporary__"
    temporary_folder.mkdir(exist_ok=True, parents=True)
    stem = uuid.uuid1()

    if memory is not None:
        if isinstance(memory, Image.Image):
            path = temporary_folder / f"{stem}.png"
            memory.save(path)
        elif isinstance(memory, bytes):
            path = temporary_folder / f"{stem}"
            path.write_bytes(memory)
        elif isinstance(memory, (str)):
            path = temporary_folder / f"{stem}"
            path.write_text(memory, "utf8")
        else:
            raise TypeError("The given memory cannot be handled. ", memory.__class__)
    elif suffix is not None:
        if not suffix.startswith("."):
            suffix = f".{suffix}"
        path = temporary_folder / "{stem}{suffix}"
    else:
        raise 

    try:
        yield path
    except Exception as e:
        if path.exists():
            os.unlink(path)
        raise e
    else:
        if path.exists():
            os.unlink(path)


if __name__ == "__main__":
    _registry_folder()
    data = {"first": 1, "second":2}
    register("TestObject", "data_key", data, ".pkl", disk=False)
    #clear()
    gained = fetch("TestObject", "data_key", disk=False)
    print(keys("TestObject"))
    print(gained)

    from PIL import Image
    image = Image.new("RGB", size=(1, 2))
    with yield_temporary_path(image) as k:
        s = k
    print(s.exists())
    print(REGISTRY_FOLDER)


