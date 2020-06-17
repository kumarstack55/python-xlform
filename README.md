[![codecov](https://codecov.io/gh/kumarstack55/python-xlform/branch/master/graph/badge.svg?token=ETG7NIKFW8)](https://codecov.io/gh/kumarstack55/python-xlform)

# python-xlform

A package for reading and writing spreadsheet such like xlsx files.

https://github.com/kumarstack55/python-xlform

## What is python-xlform for?

Spreadsheets, such as Microsoft Excel, are so versatile that they are sometimes used not only as the original spreadsheet, but also as input forms, views for table references, and outline editors for printing. On the other hand, spreadsheets are often used by the system to generate reports mechanically. Thus, spreadsheets exist between people and systems and are often difficult data formats to parse as people make changes to them.

This package provides a layer of validation so that you can avoid unintentional changes when reading and writing values and formulas in a spreadsheet.

This package allows you to perform the following tasks more securely:

* Work with spreadsheets such as Excel.
* Validate the spreadsheet.
    * Validate the sheet structure of a spreadsheet.
    * Validate that there is an expected structure at a specified location in the sheet.
    * Validate that the structure of the book or sheet is.
* Read data from any range of the spreadsheet.
    * Verify the structure of the spreadsheet's tables, including the headers.
    * Extract values from a single cell, two cells consisting of key and value, a table with no header, and a table with a header.
    * The library user defines an arbitrary structure and extracts values, formulas, formats, etc.
* Convert the data in the spreadsheet.
    * To convert to JSON format.
    * To convert to your own format.
* Write data in spreadsheets.
    * To write data in a spreadsheet with changes to some of the data you read from the spreadsheet.

## What is not python-xlform for?

* python-xlform is not a module for dealing with Form Control or AciveX controls such as Button or Checkbox in Excel spreadsheets.
* At the moment, it works with complex spreadsheets and there is still room for improvement in terms of performance.

## Why python-xlform?

* Validation
    * Validate whether a spreadsheet is in the expected format.
    * Add columns, rows, and rename headers that you don't expect
    * Handle values, formulas and formats if engine supports them
* Extensibility
    * You can implement and use your own engines for spreadsheet input and output. For example, you can use openpyxl as a backend for python-xlform. You can also use win32com as a backend to access files via COM.

## What is the difference with openpyxl?

* python-xlform uses openpyxl.
* python-xlform makes it easier to validate data.

## What is the difference with pyexcel?

* python-xlform can be used for any table structure, including values and expressions.
* pyexcel is specialized for table structures. It may have performance advantages.


## Example

### Example1: Employee

```
       1234567890123456 1234567890123456 1234567890123456 1234567890123456
-  A B C                D                E                F
1      Employee_details
2
3      Emp_ID           5
4
5      Full_Name        Adams_Vanessa
6
7      Date_Of_Joining  Thu, 25-Sep-1980
8
9      Division         HFD              Age              43
10
11     Ranking          HFD              Salary           38038
```

TODO

### Example: Table

```
- A      B     C     D      E
1 table1             table2
2
3 head1  head2       dataA  dataB
4 data11 data12      dataC  dataD
5 data21 data22
```

TODO

### Example: Key-Value

```
  A     B
1 value 10
2
```

```python
from xlform.engine.openpyxl import EngineOpenpyxl
from xlform.form import FormFactory
from xlform.form import FormItemKeyValueCells


engine = EngineOpenpyxl()
book1 = engine.open_book(...)
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
```

## Software requimenets

Python 3.6 or greater

## Installation

TODO

## Architecture

You can change the engine you are using to read and write xlsx spreadsheets
and other features that are not supported by openpyxl.

As another engine, I am developing an Excel operation by COM.
Another idea is to develop an engine that can manipulate Google Spreadsheet.

```
+----------------------------------------------------+
| xlform                                             |
+----------------------------------------------------+
+----------------------------------------------------+
| xlform engine interface                            |
+----------------------------------------------------+
+--------------+ +--------------+ +--------------+
| engine impl. | | engine impl. | | engine impl. | ...
| openpyxl     | | X            | | Y            |
+--------------+ +--------------+ +--------------+
```

## License

MIT

## See also

* [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
* [pyexcel](http://docs.pyexcel.org/en/latest/)
* [pywin32](https://github.com/mhammond/pywin32)
* [xlwings](https://www.xlwings.org/)
