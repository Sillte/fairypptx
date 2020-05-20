import pytest
from fairypptx import registory_utils

def test_registory():
    data = {"first": 1}
    category = "__test_registory_utils__"
    key = "__test"
    tuple_key = (category, key)

    registory_utils.register(category, key, data, disk=True)
    booty = registory_utils.fetch(category, key, disk=True)
    assert booty == data


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
