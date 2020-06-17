from xlform.engine.base import Book
from xlform.engine.base import Engine
from xlform.engine.base import Sheet
from xlform.engine.openpyxl import EngineOpenpyxl
from xlform.exception import XlFormArgumentException
from xlform.form import Form
from xlform.form import FormFactory
from xlform.form import FormItem
from xlform.form import FormItemCell
import unittest


class TestFormFactory(unittest.TestCase):
    def test_register_form__empty(self) -> None:
        factory: FormFactory = FormFactory()
        factory.register_form("form1", {})

    def test_register_form__multiple_items(self) -> None:
        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1", {"item1": {"cls": FormItem}, "item2": {"cls": FormItem}}
        )

    def test_register_form__item_with_kwargs(self) -> None:
        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1", {"item1": {"cls": FormItem, "kwargs": {}}}
        )

    def test_register_form__illegal_key(self) -> None:
        factory: FormFactory = FormFactory()
        with self.assertRaises(XlFormArgumentException):
            factory.register_form(
                "form1", {"item1": {"cls": FormItem, "x": {}}}
            )

    def test_new_form(self) -> None:
        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1",
            {
                "item1": {
                    "cls": FormItemCell,
                    "kwargs": {"sheet_name": "Sheet1", "range_arg": "A1"},
                }
            },
        )
        engine: Engine = EngineOpenpyxl()
        book1: Book = engine.new_book()
        book1_sheet1: Sheet = book1.get_sheets()[0]
        book1_sheet1.get_cell(1, 1).set_value(123)
        form: Form = factory.new_form("form1", book1)
        self.assertIsInstance(form, Form)


if __name__ == "__main__":
    unittest.main()
