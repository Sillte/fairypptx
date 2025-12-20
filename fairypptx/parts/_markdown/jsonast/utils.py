from pathlib import Path
import subprocess
import json 


def to_jsonast(path: Path | str): 
    return _from_str_or_path(path)

def _is_existent_path(content):
    try:
        return Path(content).exists()
    except OSError:
        return False

def _from_str_or_path(content: str | Path) -> str:
    if _is_existent_path(content):
        content = Path(content).read_text("utf8")
    else:
        content = str(content)
    ret = subprocess.run("pandoc -t json",
                         text=True,
                         stdout=subprocess.PIPE, 
                         input=content, encoding="utf8")
    assert ret.returncode == 0
    return json.loads(ret.stdout)
