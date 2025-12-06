import numpy as np


from typing import Any
from fairypptx.core.types import COMObject


class TableApiWriter:
    """Encapsulates all writing logic for Table.
    """

    def __init__(self, table_api: COMObject):
        self.table_api = table_api

    # ---- Public API -------------------------------------------------

    @property
    def size(self) -> tuple[int, int]:
        return (len(self.table_api.Rows), len(self.table_api.Columns))

    def write_cell(self, r: int, c: int, value: Any) -> None:
        cell = self.table_api.Cell(r + 1, c + 1)
        cell.Shape.TextFrame.TextRange.Text = str(value)

    def write(self, key: tuple[int | list | slice, int | list | slice], value: np.ndarray | Any) -> None:
        """Implements the same logic as the current Table.__setitem__."""
        i_row, i_col = key

        # Case 1: scalar index
        if isinstance(i_row, int) and isinstance(i_col, int):
            r = self._normalize(i_row, axis=0)
            c = self._normalize(i_col, axis=1)
            self.write_cell(r, c, value)
            return

        # Case 2: broadcast cases
        r_idx = self._to_indices(i_row, axis=0)
        c_idx = self._to_indices(i_col, axis=1)

        value = np.array(value)

        # Broadcasting
        if value.ndim == 0:
            value = np.broadcast_to(value, (len(r_idx), len(c_idx)))
        elif value.ndim == 1:
            if len(r_idx) == 1:
                value = value[None, :]
            else:
                value = value[:, None]

        assert value.shape == (len(r_idx), len(c_idx))

        for i, rr in enumerate(r_idx):
            for j, cc in enumerate(c_idx):
                self.write_cell(rr, cc, value[i, j])

    # ---- Helpers -----------------------------------------------------

    def _normalize(self, idx, axis):
        size = self.size[axis]
        return idx + size if idx < 0 else idx

    def _is_indices(self, arg):
        return isinstance(arg, (list, tuple, slice))

    def _to_indices(self, arg, axis):
        if isinstance(arg, int):
            return [self._normalize(arg, axis)]

        if isinstance(arg, (list, tuple)):
            return [self._normalize(x, axis) for x in arg]

        if isinstance(arg, slice):
            start = arg.start or 0
            stop = arg.stop or self.size[axis]
            step = arg.step or 1
            return list(range(start, stop, step))

        raise ValueError(f"Invalid index specifier: {arg}")
