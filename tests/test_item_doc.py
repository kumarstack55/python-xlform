from xlform.form import ItemDoc
import unittest


class TestItemDoc(unittest.TestCase):
    def test_get_meta(self) -> None:
        item_doc = ItemDoc(meta={"x": 1}, result=1)
        self.assertEqual(item_doc.get_meta(), {"x": 1})

    def test_get_result(self) -> None:
        item_doc = ItemDoc(result=1)
        self.assertEqual(item_doc.get_result(), 1)

    def test_get_dict(self) -> None:
        item_doc = ItemDoc(meta={"x": 1}, result=1)
        self.assertEqual(item_doc.get_dict(), {"_meta": {"x": 1}, "result": 1})


if __name__ == "__main__":
    unittest.main()
