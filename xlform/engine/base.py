from pathlib import Path
from typing_extensions import final
from typing import Any
from typing import Iterator
from typing import List
from typing import Union
from xlform.exception import XlFormInternalException
from xlform.exception import XlFormNotImplementedException
import datetime

CellValue = Union[str, float, int, datetime.datetime]


def safe_cast_cell_value(value: Any) -> CellValue:
    if isinstance(value, str):
        return value
    if isinstance(value, float):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, datetime.datetime):
        return value
    raise XlFormInternalException("Unknown type: %s" % (type(value)))


class Cell(object):
    def get_formula(self) -> CellValue:
        """Get formul(

        Returns:
            CellValue: Formula
        """
        raise XlFormNotImplementedException()

    def get_value(self) -> CellValue:
        """Get value

        If there is a formula, the calculation result is returned as a value.
        Otherwise, the value is returned.

        Raise an exception if the formula cannot be analyzed.

        Returns:
            CellValue: Value
        """
        raise XlFormNotImplementedException()

    def get_number_format(self) -> str:
        """Get number format

        Returns:
            str: Number format
        """
        raise XlFormNotImplementedException()

    def get_text(self) -> str:
        """Get text

        Raise an exception if the formula cannot be analyzed.
        Raise an exception if the number format cannot be analyzed.

        Returns:
            str: Text
        """
        raise XlFormNotImplementedException()

    def get_address(
            self, column_absolute: bool = True,
            row_absolute: bool = True) -> str:
        """Get cell address

        Args:
            column_absolute (bool, optional): True if column absolute
            row_absolute (bool, optional): True if row absolute

        Returns:
            str: A1 style absolute address like '$A$1'
        """
        raise XlFormNotImplementedException()

    def set_value(self, value: CellValue) -> None:
        """Set cell value

        Args:
            value (CellValue): Cell value
        """
        raise XlFormNotImplementedException()


class Range(object):
    def get_row(self) -> int:
        """Get row count

        Returns:
            int: Row count
        """
        raise XlFormNotImplementedException()

    def get_column(self) -> int:
        """Get column count

        Returns:
            int: Column count
        """
        raise XlFormNotImplementedException()

    def get_cell(self, row: int, column: int) -> Cell:
        """Get cell

        Args:
            row (int): Row index starting from 1
            column (int): Column index starting from 1

        Returns:
            Cell: Cell
        """
        raise XlFormNotImplementedException()


class Sheet(object):
    def get_name(self) -> str:
        """Get sheet name

        Returns:
            str: Sheet name
        """
        raise XlFormNotImplementedException()

    def get_range(self, arg: str) -> Range:
        """Get range

        Args:
            arg (str): range like 'A1', 'A1:C3'

        Returns:
            Range: Range
        """
        raise XlFormNotImplementedException()

    def get_cell(self, row: int, column: int) -> Cell:
        """Get cell

        Args:
            row (int): Row index starting from 1
            column (int): Column index starting from 1

        Returns:
            Cell: Cell
        """
        raise XlFormNotImplementedException()

    def protect(self) -> None:
        """Protect"""
        raise XlFormNotImplementedException()

    def unprotect(self) -> None:
        """Unprotect"""
        raise XlFormNotImplementedException()

    def calculate(self) -> None:
        """Calculate"""
        raise XlFormNotImplementedException()


class Book(object):
    def save(self, path: Path) -> None:
        """Save book

        Args:
            path (Path): File path
        """
        raise XlFormNotImplementedException()

    def close(self) -> None:
        """Close book"""
        raise XlFormNotImplementedException()

    def iter_sheets(self) -> Iterator[Sheet]:
        """Get sheets iterator

        Returns:
            Iterator[Sheet]: Sheets
        """
        raise XlFormNotImplementedException()

    @final
    def get_sheets(self) -> List[Sheet]:
        """Get sheets

        Returns:
            List[Sheet]: Sheets
        """
        return list(self.iter_sheets())

    def add_sheet(self, name: str) -> None:
        """Add sheet

        Args:
            name (str): Sheet name
        """
        raise XlFormNotImplementedException()


class Engine(object):
    def new_book(self) -> Book:
        """New book

        Returns:
            Book: Book
        """
        raise XlFormNotImplementedException()

    def open_book(self, path: Path) -> Book:
        """Open book

        Args:
            path (Path): File path

        Returns:
            Book: Book
        """
        raise XlFormNotImplementedException()
