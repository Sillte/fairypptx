import pytest
from fairypptx import registry_utils

def test_registry():
    data = {"first": 1}
    category = "__test_registry_utils__"
    key = "__test"
    tuple_key = (category, key)

    registry_utils.register(category, key, data, disk=True)
    booty = registry_utils.fetch(category, key, disk=True)
    assert booty == data


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
