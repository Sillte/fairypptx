""" Handle's temporary storage.
"""
import os
import shutil
import uuid

_storage_folder = "__pptx_storage__"

def _is_exist_storage_folder(current_folder):
    storage = os.path.join(current_folder, _storage_folder)
    return os.path.exists(storage)

def _assure_storage_folder():
    current_folder = os.getcwd()
    storage = os.path.join(current_folder, _storage_folder)
    if not os.path.exists(storage):
        os.makedirs(storage)
    return storage

__number = 0
def get_path(suffix=""):
    global __number
    storage_folder = _assure_storage_folder()
    filepath = os.path.join(storage_folder, f"{__number}{suffix}")
    __number += 1
    return filepath


if __name__ == "__main__":
    pass
