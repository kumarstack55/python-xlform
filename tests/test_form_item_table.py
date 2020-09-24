from typing import List
from xlform.engine.base import Book
from xlform.engine.base import Engine
from xlform.engine.base import Sheet
from xlform.engine.openpyxl import EngineOpenpyxl
from xlform.form import FormFactory
from xlform.form import FormItemTable
import unittest


class TestFormItemTable(unittest.TestCase):
    def setUp(self) -> None:
        self._engine: Engine = EngineOpenpyxl()
        self._book: Book = self._engine.new_book()
        self._sheet: Sheet = self._book.get_sheets()[0]

        table: List[List[str]] = [
            ["head1", "head1", "head2"],
            ["head11", "head12", "head21"],
            ["data111", "data112", "data121"],
            ["data211", "data212", "data221"],
            ["data311", "data312", "data321"],
        ]

        for row, rows in enumerate(table, start=1):
            for col, value in enumerate(rows, start=1):
                self._sheet.get_cell(row, col).set_value(value)

    def test_get_form_doc(self) -> None:
        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1",
            {
                "item1": {
                    "cls": FormItemTable,
                    "kwargs": {"sheet_name": "Sheet1", "range_arg": "A3:C5"},
                }
            },
        )
        form = factory.new_form("form1", self._book)
        doc = form.get_form_doc()

        self.assertIsInstance(doc, dict)
        self.assertTrue("item1" in doc)
        self.assertIsInstance(doc["item1"], dict)
        self.assertTrue("_meta" in doc["item1"])
        self.assertTrue("result" in doc["item1"])
        self.assertEqual(
            doc["item1"]["result"],
            [
                ["data111", "data112", "data121"],
                ["data211", "data212", "data221"],
                ["data311", "data312", "data321"],
            ],
        )

    def test_set_form_doc(self) -> None:
        factory: FormFactory = FormFactory()
        factory.register_form(
            "form1",
            {
                "item1": {
                    "cls": FormItemTable,
                    "kwargs": {"sheet_name": "Sheet1", "range_arg": "A3:C5"},
                }
            },
        )
        form = factory.new_form("form1", self._book)
        doc = form.get_form_doc()
        doc["item1"]["result"][1][2] = "x"
        form.set_form_doc(doc)

        self.assertEqual(self._sheet.get_cell(4, 3).get_value(), "x")


if __name__ == "__main__":
    unittest.main()
