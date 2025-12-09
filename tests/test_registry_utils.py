import pytest
from pathlib import Path
from typing import Iterator

from fairypptx.registry_utils.structure_registry import (
    StructureRegistry,
    JsonSerializer,
    PathIDLocator,
    PathStrAccessor,
)


@pytest.fixture()
def root_folder(tmp_path_factory) -> Iterator[Path]:
    yield tmp_path_factory.mktemp("registry_root")


@pytest.fixture()
def json_registry(root_folder: Path) -> StructureRegistry:
    return StructureRegistry(JsonSerializer(), PathIDLocator(root_folder), PathStrAccessor())


def test_register_and_fetch(json_registry: StructureRegistry):
    category: str = "testcat"
    key: str = "item1"
    data: dict = {"a": 1, "b": "hello"}

    json_registry.put(data, category, key)
    assert json_registry.is_exists(category, key)
    fetched = json_registry.fetch(category, key)
    assert fetched == data


def test_get_keys(json_registry: StructureRegistry) -> None:
    json_registry.put({"x": 1}, "alpha", "one")
    json_registry.put({"x": 2}, "alpha", "two")
    keys = json_registry.get_keys("alpha")
    assert keys == ["one", "two"]


def test_deregister(json_registry: StructureRegistry) -> None:
    json_registry.put({"y": 5}, "cat", "k1")
    assert json_registry.is_exists("cat", "k1")

    json_registry.remove("cat", "k1")
    assert not json_registry.is_exists("cat", "k1")
    assert json_registry.fetch("cat", "k1") is None


def test_categories_property(json_registry: StructureRegistry) -> None:
    json_registry.put(1, ("a", "b"), "x")
    json_registry.put(2, ("a", "c"), "y")
    cats = set(json_registry.categories)
    assert ("a", "b") in cats
    assert ("a", "c") in cats


def test_lru_cache_clearing(json_registry: StructureRegistry) -> None:
    json_registry.put({"v": 1}, "cachetest", "obj")
    first = json_registry.fetch("cachetest", "obj")
    assert isinstance(first, dict)
    assert first["v"] == 1
    json_registry.put({"v": 2}, "cachetest", "obj")
    second = json_registry.fetch("cachetest", "obj")
    assert isinstance(second, dict)
    assert second["v"] == 2


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
