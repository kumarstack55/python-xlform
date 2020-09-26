from abc import ABC
from abc import abstractmethod
from typing_extensions import final
from typing import Any
from typing import Dict
from typing import List
from typing import Optional
from xlform.engine.base import Book
from xlform.engine.base import CellValue
from xlform.engine.base import Range
from xlform.engine.base import Sheet
from xlform.exception import XlFormArgumentException
from xlform.exception import XlFormInternalException
from xlform.exception import XlFormNotImplementedException
from xlform.exception import XlFormValidationException
from xlform import cell_dump
import copy


ItemDocMeta = Dict[str, Any]
ItemDocResult = Any


class ItemDoc(object):
    @final
    def __init__(
        self, result: ItemDocResult, meta: Optional[ItemDocMeta] = None
    ) -> None:
        self._result: ItemDocResult = result

        if meta is None:
            meta = dict()
        if not isinstance(meta, dict):
            raise XlFormArgumentException("meta is not a dict type.")
        self._meta: ItemDocMeta = meta

    @final
    def get_meta(self) -> ItemDocMeta:
        """Get meta data"""
        return copy.deepcopy(self._meta)

    @final
    def get_result(self) -> ItemDocResult:
        """Get result"""
        return copy.deepcopy(self._result)

    @final
    def get_dict(self) -> Dict[str, Any]:
        """Get document"""
        return {"_meta": self.get_meta(), "result": self.get_result()}


class FormItem(ABC):
    @abstractmethod
    def _validate_book(self) -> None:
        """Validate book"""
        raise XlFormNotImplementedException()

    @abstractmethod
    def _validate_item_doc(self, item_doc: ItemDoc) -> None:
        """Validate item document"""
        raise XlFormNotImplementedException()

    @abstractmethod
    def _get_item_doc(self) -> ItemDoc:
        """Get item document from book"""
        raise XlFormNotImplementedException()

    @abstractmethod
    def _set_item_doc(self, item_doc: ItemDoc) -> None:
        """Set item document to book"""
        raise XlFormNotImplementedException()

    @final
    def get_item_doc(self) -> ItemDoc:
        """Get item document from book

        If validation fails, it raises an exception.

        Returns:
            ItemDoc: [TODO:description]
        """
        self._validate_book()
        item_doc = self._get_item_doc()
        assert item_doc is not None
        self._validate_item_doc(item_doc)
        return item_doc

    @final
    def set_item_doc(self, item_doc: ItemDoc) -> None:
        """Set item document to book

        If validation fails, it raises an exception.
        When the exception is raised, the book may have changed.

        Args:
            item_doc (ItemDoc): Item document
        """
        self._validate_item_doc(item_doc)
        self._validate_book()
        self._set_item_doc(item_doc)
        self._validate_book()


class FormItemCell(FormItem):
    def __init__(self, book: Book, sheet_name: str, range_arg: str):
        self._book = book
        self._sheet_name = sheet_name
        self._range_arg = range_arg
        try:
            self._validate_book()
        except XlFormValidationException:
            raise XlFormArgumentException()

    def _find_sheet(self, sheet_name: str) -> Sheet:
        for sheet in self._book.iter_sheets():
            if sheet.get_name() == self._sheet_name:
                return sheet
        raise XlFormArgumentException()

    def _validate_book(self) -> None:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        if r.get_rows_count() != 1 or r.get_columns_count() != 1:
            raise XlFormArgumentException()

    def _validate_item_doc(self, item_doc: ItemDoc) -> None:
        pass

    def _get_item_doc(self) -> ItemDoc:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        cell = r.get_cell(1, 1)
        return ItemDoc(meta=cell_dump(cell), result=cell.get_value())

    def _set_item_doc(self, item_doc: ItemDoc) -> None:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        r.get_cell(1, 1).set_value(item_doc.get_result())


class FormItemKeyValueCells(FormItem):
    """
    A form item with a range of two columns and one row, with the left cell
    being the header
    """

    def __init__(
        self,
        book: Book,
        sheet_name: str,
        range_arg: str,
        header_value: CellValue,
    ):
        self._book = book
        self._sheet_name = sheet_name
        self._range_arg = range_arg
        self._header_value = header_value
        try:
            self._validate_book()
        except XlFormValidationException as e:
            raise XlFormArgumentException("Illegal argument: %s" % (str(e)))

    def _find_sheet(self, sheet_name: str) -> Sheet:
        for sheet in self._book.iter_sheets():
            if sheet.get_name() == self._sheet_name:
                return sheet
        raise XlFormArgumentException()

    def _validate_book(self) -> None:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        if r.get_cell(1, 1).get_value() != self._header_value:
            raise XlFormValidationException("header_value not found.")

    def _validate_item_doc(self, item_doc: ItemDoc) -> None:
        pass

    def _get_item_doc(self) -> ItemDoc:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        meta = dict()
        cell = r.get_cell(1, 2)
        meta.update(cell_dump(cell))
        return ItemDoc(meta=meta, result=cell.get_value())

    def _set_item_doc(self, item_doc: ItemDoc) -> None:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        r.get_cell(1, 2).set_value(item_doc.get_result())


class FormItemTable(FormItem):
    def __init__(
        self,
        book: Book,
        sheet_name: str,
        range_arg: str,
        header_rows_count: int = 0,
        header_path_list: Optional[List[List[str]]] = None,
    ):
        self._book = book
        self._sheet_name = sheet_name
        self._range_arg = range_arg
        self._header_rows_count = header_rows_count
        self._header_path_list = header_path_list

        if self._header_rows_count < 0:
            raise XlFormArgumentException()
        elif self._header_rows_count > 1:
            if not isinstance(self._header_path_list, list):
                raise XlFormArgumentException()
            try:
                self._validate_book()
            except XlFormValidationException as e:
                raise XlFormArgumentException("Illegal argument: %s" % (e))
            for header_path in self._header_path_list:
                if not isinstance(header_path, list):
                    raise XlFormArgumentException(
                        "not isinstance(header_path, list)"
                    )
                if len(header_path) != self._header_rows_count:
                    raise XlFormArgumentException(
                        "len(header_path) != self._header_rows_count"
                    )

    def _find_sheet(self, sheet_name: str) -> Sheet:
        for sheet in self._book.iter_sheets():
            if sheet.get_name() == self._sheet_name:
                return sheet
        raise XlFormArgumentException()

    def _validate_book(self) -> None:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        if r.get_rows_count() <= self._header_rows_count:
            raise XlFormValidationException()
        if r.get_columns_count() <= 0:
            raise XlFormInternalException()
        if self._header_rows_count == 0:
            pass
        elif self._header_rows_count >= 1:
            if self._header_path_list is None:
                raise XlFormInternalException()
            if len(self._header_path_list) != r.get_columns_count():
                raise XlFormArgumentException(
                    "len(self._header_path_list) != r.get_columns_count()"
                )
            for col_index, header_path in enumerate(
                self._header_path_list, start=1
            ):
                for row_index, header in enumerate(header_path, start=1):
                    if r.get_cell(row_index, col_index).get_value() != header:
                        XlFormValidationException()
        else:
            raise XlFormNotImplementedException()

    def _validate_item_doc_row_list(
        self, range_: Range, row: List[CellValue], row_index: int
    ) -> None:
        if len(row) != range_.get_columns_count():
            raise XlFormValidationException(
                "len(row) != range_.get_columns_count(): (%s, %d)"
                % (row, range_.get_columns_count())
            )

    def _validate_item_doc_row_dict(
        self, range_: Range, row: Dict[str, Any], row_index: int
    ) -> None:
        assert isinstance(self._header_path_list, list)
        for col_index, header_path in enumerate(self._header_path_list):
            dic = row
            for header_path_index in range(0, len(header_path) - 1):
                path_part = header_path[header_path_index]
                if path_part not in dic:
                    raise XlFormValidationException(
                        "path_part not in dic: header_path=%s, row=%s"
                        % (header_path, row)
                    )
                dic = dic[path_part]
                if not isinstance(dic, dict):
                    raise XlFormValidationException(
                        "not isinstance(dic, dict): row=%s, header_path=%s"
                        % (row, header_path)
                    )
            if not header_path[-1] in dic:
                raise XlFormValidationException()

    def _validate_item_doc(self, item_doc: ItemDoc) -> None:
        result = item_doc.get_result()

        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        if not isinstance(result, list):
            raise XlFormValidationException()

        data_rows_count = r.get_rows_count() - self._header_rows_count
        if len(result) != data_rows_count:
            raise XlFormValidationException(
                "len(result) != data_rows_count: %d, %d"
                % (len(result), data_rows_count)
            )
        for row_index in range(0, data_rows_count):
            row = result[row_index]
            if isinstance(row, list):
                self._validate_item_doc_row_list(r, row, row_index)
            elif isinstance(row, dict):
                self._validate_item_doc_row_dict(r, row, row_index)
            else:
                raise XlFormValidationException()

    def _get_item_doc_row_list(
        self, meta: Dict[str, Any], range_: Range
    ) -> ItemDoc:
        result_list: List[List[CellValue]] = list()
        for row_index in range(1, range_.get_rows_count() + 1):
            row_list: List[CellValue] = list()
            for col_index in range(1, range_.get_columns_count() + 1):
                cell = range_.get_cell(row_index, col_index)
                meta.update(cell_dump(cell))
                row_list.append(cell.get_value())
            result_list.append(row_list)
        return ItemDoc(meta=meta, result=result_list)

    def _get_item_doc_row_dict(
        self, meta: Dict[str, Any], range_: Range
    ) -> ItemDoc:
        result_list: List[Dict[str, CellValue]] = list()
        start = 1 + self._header_rows_count
        for row_index in range(start, range_.get_rows_count() + 1):
            row_dict: Dict[str, CellValue] = dict()
            for col_index in range(1, range_.get_columns_count() + 1):
                cell = range_.get_cell(row_index, col_index)
                meta.update(cell_dump(cell))
                assert self._header_path_list is not None
                header_path = self._header_path_list[col_index - 1]

                dic: Dict[str, Any] = row_dict
                for header_path_index in range(0, len(header_path) - 1):
                    path_part = header_path[header_path_index]
                    if path_part not in dic:
                        dic[path_part] = dict()
                    if not isinstance(dic[path_part], dict):
                        raise XlFormInternalException()
                    dic = dic[path_part]
                dic[header_path[-1]] = cell.get_value()
            result_list.append(row_dict)
        return ItemDoc(meta=meta, result=result_list)

    def _get_item_doc(self) -> ItemDoc:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)
        meta: Dict[str, Any] = dict()
        if self._header_rows_count == 0:
            return self._get_item_doc_row_list(meta, r)
        elif self._header_rows_count >= 1:
            return self._get_item_doc_row_dict(meta, r)
        else:
            raise XlFormInternalException()

    def _set_item_doc(self, item_doc: ItemDoc) -> None:
        sheet = self._find_sheet(self._sheet_name)
        r = sheet.get_range(self._range_arg)

        result = item_doc.get_result()
        if isinstance(result, list):
            rows_count = len(result)
            if r.get_rows_count() - self._header_rows_count != rows_count:
                raise XlFormArgumentException()
            for row in result:
                if len(row) != r.get_columns_count():
                    raise XlFormArgumentException()

            for row_index, row in enumerate(result, start=1):
                for col_index, value in enumerate(row, start=1):
                    r.get_cell(row_index, col_index).set_value(value)

        elif isinstance(result, dict):
            raise XlFormNotImplementedException()


class Form(object):
    @final
    def __init__(self) -> None:
        self._form_item_dic: Dict[str, FormItem] = dict()

    @final
    def add_form_item(self, name: str, form_item: FormItem) -> None:
        self._form_item_dic[name] = form_item

    @final
    def get_form_doc(self) -> Dict[str, Any]:
        dic: Dict[str, Any] = dict()
        for form_item_name, form_item in self._form_item_dic.items():
            item_doc = form_item.get_item_doc()
            dic[form_item_name] = item_doc.get_dict()
        return dic

    @final
    def set_form_doc(self, doc: Dict[str, Any]) -> None:
        form_item_name_list: List[str] = list(doc.keys())

        for form_item_name in form_item_name_list:
            if form_item_name not in self._form_item_dic:
                raise XlFormArgumentException()

        for form_item_name in form_item_name_list:
            form_item = self._form_item_dic[form_item_name]
            if not isinstance(doc[form_item_name], dict):
                raise XlFormArgumentException()
            if "result" not in doc[form_item_name]:
                raise XlFormArgumentException()
            result = doc[form_item_name]["result"]
            item_doc = ItemDoc(result=result)
            form_item.set_item_doc(item_doc)


class FormFactory(object):
    def __init__(self) -> None:
        self._form_dic: Dict[str, Dict[str, Any]] = dict()

    def register_form(
        self, name: str, form_item_cls_kwargs_dic: Dict[str, Dict[str, Any]]
    ) -> None:
        """Registering the FormItem classes that constitutes the Form

        Args:
            name (str): Form name
            form_item_cls_kwargs_dic (Dict[str, Dict[str, Any]]): cls and
            the constructor arguments
        """
        for form_item, cls_kwargs_dic in form_item_cls_kwargs_dic.items():
            KEY_CLS = "cls"
            if KEY_CLS not in cls_kwargs_dic:
                raise XlFormArgumentException(
                    "Key %s does not exist." % (KEY_CLS)
                )

            keys_exists = {key: True for key in cls_kwargs_dic.keys()}
            keys_exists.pop(KEY_CLS)
            keys_may_exists = ["kwargs"]
            for k in keys_may_exists:
                if k in keys_exists:
                    keys_exists.pop(k)
            if len(keys_exists) > 0:
                raise XlFormArgumentException(
                    "Unknown keys: %s" % (keys_exists)
                )

        self._form_dic[name] = form_item_cls_kwargs_dic

    def new_form(self, name: str, book: Book) -> Form:
        """Create a new form

        Args:
            name (str): Form name
            book (Book): Book

        Returns:
            Form: Form object
        """
        form_item_cls_kwargs_dic = self._form_dic[name]

        form: Form = Form()
        for form_item_name, cls_kwargs_dic in form_item_cls_kwargs_dic.items():
            cls, kwargs = cls_kwargs_dic["cls"], cls_kwargs_dic["kwargs"]
            form_item = cls(book=book, **kwargs)
            form.add_form_item(form_item_name, form_item)
        return form
