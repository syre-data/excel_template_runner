from typing import Union, Any, Optional
from enum import Enum
from dataclasses import dataclass
import openpyxl
import openpyxl.worksheet.worksheet
from syre import _LEGACY_

DataFormatType = Enum("DataFormatType", ["SPREADSHEET", "EXCEL_WORKBOOK"])
"""Identifier for the type of input data to expect.

SPREADSHEET: A simple spreadsheeet. (e.g. csv)
EXCEL_WORKBOOK: An Excel workbook.
"""

WorkbookSheetId = Enum("WorkbookSheetId", ["INDEX", "TITLE"])
"""Identifier kind for an Excel workbook sheet.

INDEX: 0-based index of the sheet.
TITLE: Name of the sheet.
"""

ColumnId = Enum("ColumnId", ["HEADER", "INDEX"])
"""Identifier kind for columns.

HEADER: Select based on header labels.
INDEX: Select by column index (0-based). e.g. 0 -> A, 1 -> B
"""

InputColumnsRangeType = Enum(
    "InputColumnsRangeType", ["SPECIFIED", "RANGE_UNTIL_BREAK"]
)
"""How to determinate data break from data source.

SPECIFIED: Start and end are specified explicitly.
RANGE_UNTIL_BREAK: Start is specified explicitly, data is copied until the first empty column.
"""

HeaderAction = Enum("HeaderAction", ["NONE", "INSERT", "REPLACE"])
"""How header info from data files should be manipulated.

NONE: Data is copied in as-is from the input source.
INSERT: Inserts the file name of the data source as an additional header above all others.
REPLACE: Replaces all data headers with the file name.
"""


@dataclass
class SpreadsheetDataArgs:
    """Data arguments for spreadsheet style data.
    See `pandas.read_csv` for more info.

    Args:
        skip_rows (int, optional): Number of rows to skip until the first header or data. Defaults to 0.
        comment (Optional[str], optional): Empty string or single character indicating a line is a comment and should be ignored. Defaults to None.
    """

    skip_rows: int = 0
    comment: Optional[str] = ""

    def __post_init__(self):
        if self.comment is not None and len(self.comment) != 1:
            raise ValueError("`comment` must be a single character")


WorksheetId = Union[str, int]
"""Worksheet identifier. Can be either the 0-based index of the worksheet within the workbook, or the worksheet's label.
"""


@dataclass
class ExcelDataArgs:
    """Data arguments for spreadsheet style data.

    Args:
        sheet (int, str): Worksheet id. 0-based index or sheet label.
        skip_rows (int, optional): Number of rows to skip until the first header or data. Defaults to 0.
    """

    sheet: WorksheetId
    skip_rows: int = 0


Worksheet = openpyxl.worksheet.worksheet.Worksheet
DataFormatArgs = Union[SpreadsheetDataArgs, ExcelDataArgs]

if _LEGACY_:
    from typing import List, Dict, Tuple

    Metadata = Dict[str, Any]
    MetadataArgs = List[str]
    ColumnSelection = Union[List[List[str]], List[str], List[int]]
    AssetFilter = Dict[str, Any]
    AssetProperties = Dict[str, Any]
    DataSelectionArgs = List[str]
    ReplaceRange = Tuple[int, int]

else:
    Metadata = dict[str, Any]
    MetadataArgs = list[str]
    ColumnSelection = Union[list[list[str]], list[str], list[int]]
    AssetFilter = dict[str, Any]
    AssetProperties = dict[str, Any]
    DataSelectionArgs = list[str]
    ReplaceRange = tuple[int, int]
