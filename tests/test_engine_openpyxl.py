from typing import Any
from unittest import TestLoader
from unittest import TestSuite
from xlform.engine.openpyxl import EngineOpenpyxl
from xlform.engine.test import EngineTestCase
import xlform.engine.test
import unittest


class TestEngineOpenpyxl(EngineTestCase):
    def setUp(self) -> None:
        self._engine = EngineOpenpyxl()

    def test_cell_get_text(self) -> None:
        self.skipTest("EngineOpenpyxl doesn't support evaluation of formula.")

    def test_sheet_protect(self) -> None:
        # Although openpyxl has the ability to protect sheets,
        # it does not appear to support file saving for protection.
        # Therefore, the engine does not implement sheet protection.
        self.skipTest("not implemented")

    def test_sheet_unprotect(self) -> None:
        # Although openpyxl has the ability to protect sheets,
        # it does not appear to support file saving for protection.
        # Therefore, the engine does not implement sheet protection.
        self.skipTest("not implemented")


def load_tests(loader: TestLoader, tests: Any, patterns: Any) -> TestSuite:
    return xlform.engine.test.load_tests(loader, (TestEngineOpenpyxl,))


if __name__ == "__main__":
    unittest.main()
