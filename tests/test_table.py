import pytest
import numpy as np
import pandas as pd  
from fairypptx.table import Table, DFTable, Row, Rows

def test_init():
    array = np.arange(8).reshape(2, 4)
    table = Table.make(array)
    assert table[0, 0].text == str(0)
    assert table[0, 1].text == str(1)
    table[0, 1] = 10
    assert table[0, 1].text == str(10)

def test_df_table():
    df = pd.DataFrame(np.arange(12).reshape(3, 4))
    df.index = pd.MultiIndex.from_tuples([("ア", "A"), ("ア", "B"), ("ア", "C")])
    df.columns = ["W", "X", "Y", "Z"]
    table = DFTable(df)
    read_df = table.df
    assert df.equals(read_df)

def test_insert():
    """Check behavior of `insert` of `Row` / `Rows`.
    """
    # Check `insert` - for int.
    table = Table.empty((5, 2))
    shape = table.shape
    row = table.rows.insert(1)
    assert table.size == (6, 2)
    row[1].shape.text = "Hello"

    table = Table(shape)
    assert table.size == (6, 2)
    assert table[1, 1].text == "Hello"

    # Check `insert` - for Sequence of int.
    table = Table.empty((5, 2))
    shape = table.shape
    rows = table.rows.insert([4, 2])
    assert table.size == (7, 2)
    rows[0][1].shape.text = "Werewolf"
    rows[1][0].shape.text = "Human"

    table = Table(shape)
    assert table.size == (7, 2)
    assert table[4, 1].shape.text == "Werewolf"
    assert table[2, 0].text == "Human"

def test_delete():
    """Check behavior of `insert` of `Row` / `Rows`.
    """
    def _is_empty(table):
        return all(table[i, j].shape.text == "" for i in range(table.size[0]) for j in range(table.size[1]))

    # Check `delete` - for int.
    table = Table.empty((4, 2))
    assert _is_empty(table)
    table[2, 0] = "AAA"
    assert not _is_empty(table)
    table.rows.delete(2)
    assert _is_empty(table)
    assert table.size == (3, 2)

    # Check `delete` - for Sequence of `int`.
    table = Table.empty((10, 2))
    assert _is_empty(table)
    table[2, 0] = "AAA"
    table[7, 1] = "BBB"
    assert not _is_empty(table)
    table.rows.delete([2, 7])
    assert _is_empty(table)

def test_tighten():
    """Check behavior of `tighten`.
    """
    table = Table.empty((4, 2))
    table.rows[0].api.Height = 111
    assert table.rows[0].height == 111
    table.rows[0].tighten()
    assert table.rows[0].height != 111

    table = Table.empty((2, 2))
    table.rows[1].shapes[0].text = "OneLine\nTwoLine\nThreeLine"
    table.rows.tighten()
    

if __name__ == "__main__":
    pass
    # test_tighten()
    pytest.main([__file__, "--capture=no"])
