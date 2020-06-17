from typing import Any
from typing import Dict
from xlform.engine.base import Cell
from xlform.exception import XlFormArgumentException
from xlform.exception import XlFormException

__version__ = "0.1.0"


def cell_dump(cell: Cell) -> Dict[str, Any]:
    if not isinstance(cell, Cell):
        raise XlFormArgumentException()

    addr: str = cell.get_address(column_absolute=False, row_absolute=False)
    value: Dict[str, Any] = dict()
    try:
        value["formula"] = cell.get_formula()
    except XlFormException:
        pass
    try:
        value["value"] = cell.get_value()
    except XlFormException:
        pass
    try:
        value["number_format"] = cell.get_number_format()
    except XlFormException:
        pass
    try:
        value["text"] = cell.get_text()
    except XlFormException:
        pass
    return {addr: value}
