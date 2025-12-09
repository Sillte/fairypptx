"""Registry utilities for storing small structured artifacts on disk.

This module provides a small, extensible registry abstraction composed of three
roles:
- Serializer: converts between Python objects and storage representation.
- IDLocator: maps (category, key, file_type) to a storage identifier (e.g. Path).
- Accessor: performs read/write/delete/exists on the identifier.

The core class `StructureRegistry` composes these roles to provide a simple
API for registering, fetching, listing and removing named artifacts grouped by
category. Implementations are intentionally small and intended to be
replaced/extended in tests or other environments.
"""

import json
import importlib
from pathlib import Path
from typing import Sequence, Any, Annotated
import pickle
from functools import lru_cache 
from fairypptx.registry_utils.utils import get_registry_folder
from typing import Protocol  
from pydantic import Field, BaseModel


FileType = Annotated[str, Field(pattern=r"^[a-zA-Z0-9_]+$")]
CategoryElement = Annotated[str, Field(pattern=r"^[a-zA-Z0-9_]+$")]

# When you extend other extensions, 
type Category = tuple[CategoryElement, ...]


class SerializerProtocol[T](Protocol):
    """Protocol for serializer implementations.

    A serializer converts between Python objects and a storage representation
    (for example, JSON text or a pickle bytes object). Implementations should
    be effectively stateless (or immutable) so they are safe to reuse.

    - `file_type` should return the short file-type identifier used by the
      registry (e.g. "json", "pkl").
    - `to_atomic` converts a Python object into the storage type `T`.
    - `from_atomic` converts storage value `T` back to a Python object (or
      returns `None` when input is `None`).
    """

    @property
    def file_type(self) -> FileType:
        ...

    def to_atomic(self, data: Any) -> T:
        ...

    def from_atomic(self, atomic: T | None) -> Any | None:
        ...

class JsonSerializer(SerializerProtocol[str]):
    """Simple JSON serializer.

    This implementation uses the stdlib `json` module. It produces and
    consumes UTF-8 JSON text. Implementations that need different formatting
    options (indentation, ensure_ascii, etc.) should provide their own
    configurable instance.
    """

    @property
    def file_type(self) -> FileType: 
        return "json"

    def to_atomic(self, data: Any) -> str:
        return json.dumps(data, indent=4)

    def from_atomic(self, atomic: str | None) -> Any | None:
        if atomic:
            # json.loads does not accept an `indent` argument; just parse.
            return json.loads(atomic)
        return None

class BaseModelSerializer(SerializerProtocol[str]):
    @property
    def file_type(self) -> FileType: 
        return "basemodel"

    def to_atomic(self, data: BaseModel) -> str:
        return json.dumps({
            "__class__": data.__class__.__module__ + "." + data.__class__.__name__,
            "__data__": data.model_dump()
        }, indent=4)

    def from_atomic(self, atomic: str | None) -> BaseModel | None:
        if not atomic:
            return None
        raw = json.loads(atomic)
        class_path = raw["__class__"]
        data = raw["__data__"]

        module_name, class_name = class_path.rsplit(".", 1)
        module = importlib.import_module(module_name)
        cls = getattr(module, class_name)
        return cls.model_validate(data)
        

class IDLocatorProtocol[T](Protocol):
    """Protocol that maps category/key/file_type to storage identifiers.

    An `IDLocator` is responsible for the mapping between a logical
    (category, key, file_type) triple and a concrete storage id (for example
    a `pathlib.Path`). It also enumerates known categories and keys.
    """

    def make_id(self, category: Category, key: str, file_type: FileType) -> T:
        ...

    def to_category(self, id_: T, file_type: FileType) -> Category:
        ...

    def to_key(self, id_: T, file_type: FileType) -> str:
        ...
    
    @property
    def categories(self) -> Sequence[Category]: 
        ...

    def get_ids(self, category: Category, file_type: FileType) -> Sequence[T]:
        ...


class PathIDLocator(IDLocatorProtocol[Path]):
    """File-system based ID handler using a root registry folder.

    The handler maps a category (tuple of path segments) and key to a
    filesystem `Path` under the configured registry root folder.
    """

    def __init__(self, root_folder: Path | None = None):
        if root_folder is None:
            root_folder = get_registry_folder()
        self.root_folder = root_folder

    def make_id(self, category: Category, key: str, file_type: FileType) -> Path:
        folder = self.root_folder.joinpath(*category)
        return folder / f"{key}.{file_type}"

    def to_category(self, id_: Path, file_type: FileType) -> Category:
        _ = file_type
        return id_.parent.relative_to(self.root_folder).parts

    def to_key(self, id_: Path, file_type: FileType) -> str:
        _ = file_type
        return id_.stem
    
    def get_ids(self, category: Category, file_type: FileType) -> Sequence[Path]:
        folder = self.root_folder.joinpath(*category)
        return sorted(path for path in folder.glob(f"*.{file_type}"))

    @property
    def categories(self) -> Sequence[Category]: 
        categories = set()
        for path in self.root_folder.glob("**/*.*"):
            category = path.parent.relative_to(self.root_folder).parts
            categories.add(category)
        return list(categories)



class AccessorProtocol[T_id, T_data](Protocol):
    """Protocol that performs storage operations for a given identifier.

    Implementations should perform the minimal I/O required and raise
    exceptions for I/O failures. Methods are:
    - `read(id_) -> data` : return stored data for identifier
    - `write(id_, data)` : persist `data` under `id_`
    - `delete(id_)` : remove stored item
    - `exists(id_) -> bool` : check for presence
    """

    def read(self, id_: T_id) -> T_data:
        ...

    def write(self, id_: T_id, data: T_data) -> None:
        ...

    def delete(self, id_: T_id) -> None:
        ...
        
    def exists(self, id_: T_id) -> bool:
        ...

class PathStrAccessor(AccessorProtocol[Path, str]):
    """Accessor using filesystem text files.

    - `read`/`write` operate on text content encoded in the filesystem.
    - `write` will create parent directories as needed.
    """

    def read(self, id_: Path): 
        return id_.read_text()

    def write(self, id_: Path, data: str) -> None: 
        id_.parent.mkdir(parents=True, exist_ok=True)
        id_.write_text(data)

    def delete(self, id_: Path) -> None:
        id_.unlink()
        
    def exists(self, id_: Path) -> bool:
        return id_.exists()

 

class StructureRegistry[T_data, T_id]:
    """Registry façade that composes serializer, id-handler and accessor.

    Usage:
        registry = StructureRegistry(JsonSerializer(), PathIDLocator(), PathStrAccessor())
        registry.register(obj, ('mycategory',), 'name')
        data = registry.fetch(('mycategory',), 'name')

    The registry caches fetched items (LRU). Callers that modify the
    underlying storage via `register`/`deregister` will clear the cache
    automatically.
    """
    def __init__(self,
                 serializer: SerializerProtocol[T_data],
                 locator: IDLocatorProtocol[T_id], 
                 accessor: AccessorProtocol[T_id, T_data]):
        self.serializer = serializer
        self.locator = locator
        self.accessor = accessor

    def _to_category(self, category: Category | str) -> Category:
        if isinstance(category, str):
            return (category, )
        return category

    def put(self, obj: Any, category: Category | str, key: str) -> None:
        """Persist `obj` under `(category, key)`.

        - `category` may be a string or a Category tuple.
        - `key` is the resource name (filename stem without extension).
        This method writes through the accessor and clears the fetch cache.
        """
        category = self._to_category(category)
        id_ = self._to_storage_id(category, key)
        storage_data = self.serializer.to_atomic(obj)
        self.accessor.write(id_, storage_data)
        self._fetch_from_storage.cache_clear()

    def is_exists(self, category: Category | str, key: str) -> bool:
        """Return True when the (category, key) entry exists in storage."""
        category = self._to_category(category)
        id_ = self.locator.make_id(category, key, self.serializer.file_type)
        return self.accessor.exists(id_)


    def remove(self, category: Category | str, key: str) -> None:
        """Remove the stored entry for `(category, key)` and clear cache."""
        category = self._to_category(category)
        id_ = self._to_storage_id(category, key)
        self.accessor.delete(id_)
        self._fetch_from_storage.cache_clear()

    @property
    def categories(self) -> Sequence[Category]:
        """Return known categories as reported by the `IDLocator`."""
        return self.locator.categories

    def get_keys(self, category: Category | str) -> Sequence[str]:
        """Return a sorted list of keys for `category` (by file_type)."""
        category = self._to_category(category)
        file_type = self.serializer.file_type
        return [self.locator.to_key(id_, file_type) for id_ in self.locator.get_ids(category, file_type)]

    def fetch(self, category: Category | str, key: str) -> Any | None:
        """Fetch and return the stored object for `(category, key)`.

        Returns `None` when the item is not present. Under the hood this
        method uses an LRU-cached helper to avoid repeated deserialization.
        """
        category = self._to_category(category)
        id_ = self.locator.make_id(category, key, self.serializer.file_type)
        if self.accessor.exists(id_):
            return self._fetch_from_storage(category, key) 
        return None

    @lru_cache(maxsize=256)
    def _fetch_from_storage(self, category: Category, key: str) -> Any:
        """Internal cached read+deserialize helper.

        This method is decorated with LRU cache; callers that mutate storage
        must call `cache_clear()` on this function (the public methods
        `register`/`deregister` already do that).
        """
        id_ = self._to_storage_id(category, key)
        storage_data = self.accessor.read(id_)
        return self.serializer.from_atomic(storage_data)

    def _to_storage_id(self, category: Category, key: str) -> T_id:
        """Return the low-level storage identifier for the given logical id."""
        return self.locator.make_id(category, key, self.serializer.file_type)


JsonRegistry = StructureRegistry(JsonSerializer(), PathIDLocator(), PathStrAccessor())
BaseModelRegistry = StructureRegistry(BaseModelSerializer(), PathIDLocator(), PathStrAccessor())

