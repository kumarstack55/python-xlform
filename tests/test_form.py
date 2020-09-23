from typing import Any
from typing import Dict
from xlform.exception import XlFormNotImplementedException
from xlform.form import Form
from xlform.form import FormItem
from xlform.form import ItemDoc
import unittest


class TestForm(unittest.TestCase):
    class FormItemImpl(FormItem):
        def _validate_book(self) -> None:
            pass

        def _validate_item_doc(self, item_doc: ItemDoc) -> None:
            pass

        def _get_item_doc(self) -> ItemDoc:
            return ItemDoc(result=1)

        def _set_item_doc(self, item_doc: ItemDoc) -> None:
            raise XlFormNotImplementedException()

    def test_add_form_item(self) -> None:
        f = Form()
        form_item = self.FormItemImpl()
        f.add_form_item("item1", form_item)

    def test_get_form_doc(self) -> None:
        f = Form()
        form_item = self.FormItemImpl()
        f.add_form_item("item1", form_item)
        doc: Dict[str, Any] = f.get_form_doc()

        self.assertIsInstance(doc, dict)
        self.assertTrue("item1" in doc)
        self.assertIsInstance(doc["item1"], dict)
        self.assertTrue("_meta" in doc["item1"])
        self.assertTrue("result" in doc["item1"])
        self.assertEqual(doc["item1"]["result"], 1)


if __name__ == "__main__":
    unittest.main()
