from pathlib import Path 

def get_registry_folder() -> Path:
    folder =  Path.home() / ".fairypptx" / "registry"
    folder.mkdir(parents=True, exist_ok=True)
    return folder
