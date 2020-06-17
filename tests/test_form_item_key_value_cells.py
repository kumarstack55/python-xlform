from xlform.engine.base import Book
from xlform.engine.base import Engine
from xlform.engine.base import Sheet
from xlform.engine.openpyxl import EngineOpenpyxl
from xlform.exception import XlFormArgumentException
from xlform.form import FormFactory
from xlform.form import FormItemKeyValueCells
import unittest


class TestFormItemKeyValueCells(unittest.TestCase):
    def test_get_form_doc(self) -> None:
        engine: Engine = EngineOpenpyxl()
        book1: Book = engine.new_book()
        book1_sheet1: Sheet = book1.get_sheets()[0]
        book1_sheet1.get_cell(1, 1).set_value("value")
        book1_sheet1.get_cell(1, 2).set_value(10)

        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1",
            {
                "item1": {
                    "cls": FormItemKeyValueCells,
                    "kwargs": {
                        "sheet_name": "Sheet1",
                        "range_arg": "A1:B1",
                        "header_value": "value",
                    },
                }
            },
        )
        form = factory.new_form("form1", book1)
        doc = form.get_form_doc()

        self.assertIsInstance(doc, dict)
        self.assertTrue("item1" in doc)
        self.assertIsInstance(doc["item1"], dict)
        self.assertTrue("_meta" in doc["item1"])
        self.assertTrue("result" in doc["item1"])
        self.assertEqual(doc["item1"]["result"], 10)

    def test_get_form_doc__validation(self) -> None:
        engine: Engine = EngineOpenpyxl()
        book1: Book = engine.new_book()
        book1_sheet1: Sheet = book1.get_sheets()[0]
        book1_sheet1.get_cell(1, 1).set_value("xyz")
        book1_sheet1.get_cell(1, 2).set_value(10)

        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1",
            {
                "item1": {
                    "cls": FormItemKeyValueCells,
                    "kwargs": {
                        "sheet_name": "Sheet1",
                        "range_arg": "A1:B1",
                        "header_value": "value",
                    },
                }
            },
        )
        with self.assertRaises(XlFormArgumentException):
            factory.new_form("form1", book1)

    def test_set_form_doc(self) -> None:
        engine: Engine = EngineOpenpyxl()
        book1: Book = engine.new_book()
        book1_sheet1: Sheet = book1.get_sheets()[0]
        book1_sheet1.get_cell(1, 1).set_value("value")
        book1_sheet1.get_cell(1, 2).set_value(10)

        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1",
            {
                "item1": {
                    "cls": FormItemKeyValueCells,
                    "kwargs": {
                        "sheet_name": "Sheet1",
                        "range_arg": "A1:B1",
                        "header_value": "value",
                    },
                }
            },
        )
        form = factory.new_form("form1", book1)
        doc = form.get_form_doc()
        doc["item1"]["result"] = 20
        form.set_form_doc(doc)

        self.assertEqual(book1_sheet1.get_cell(1, 2).get_value(), 20)


if __name__ == "__main__":
    unittest.main()
