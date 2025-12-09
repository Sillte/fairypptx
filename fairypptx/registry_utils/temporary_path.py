from __future__ import annotations
from pathlib import Path
from contextlib import contextmanager
from typing import Iterator
import uuid
from PIL import Image

from fairypptx.registry_utils.utils import get_registry_folder


def _get_temporary_folder() -> Path:
    registry_folder = get_registry_folder()
    temp = registry_folder / "__$temporary$__"
    temp.mkdir(exist_ok=True, parents=True)
    return temp


@contextmanager
def yield_temporary_path(suffix: str = "") -> Iterator[Path]:
    """
    Create an empty temporary file that exists during the context.
    Suitable for COM API calls requiring actual files.
    """
    folder = _get_temporary_folder()
    p = folder / f"{uuid.uuid4()}{suffix}"
    p.touch()

    try:
        yield p
    finally:
        if p.exists():
            p.unlink()


@contextmanager
def yield_temporary_dump(obj: Image.Image | bytes | str, suffix: str | None = None) -> Iterator[Path]:
    """Save given memory object (PIL Image, bytes, str) into a temporary file.
    """
    folder = _get_temporary_folder()

    # Auto suffix if not provided
    if suffix is None:
        suffix = ".png" if isinstance(obj, Image.Image) else ""

    p = folder / f"{uuid.uuid1()}{suffix}"

    if isinstance(obj, Image.Image):
        obj.save(p)
    elif isinstance(obj, bytes):
        p.write_bytes(obj)
    elif isinstance(obj, str):
        p.write_text(obj, encoding="utf8")
    else:
        raise TypeError(f"Unsupported type: {type(obj)}")

    try:
        yield p
    finally:
        if p.exists():
            p.unlink()
