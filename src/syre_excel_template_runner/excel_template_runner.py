from typing import Optional
import os

import pandas
import openpyxl as xl
from openpyxl import utils as xl_utils
from openpyxl.utils import exceptions as xl_exceptions
from openpyxl.formula.tokenizer import Tokenizer, Token
from openpyxl.formula.translate import Translator
import formulas
import syre

from . import utils
from .types import (
    DataFormatType,
    DataFormatArgs,
    WorksheetId,
    ColumnId,
    Worksheet,
    ColumnSelection,
    HeaderAction,
    AssetFilter,
    AssetProperties,
    ReplaceRange,
)


def selection_type(selector: ColumnSelection) -> ColumnId:
    """Determines the type of selector.

    Args: s
        selector (ColumnSelection): Selector to classify.

    Raises:
        ValueError: If the selector is invalid.

    Returns:
        SelectionType:
    """
    if type(selector) is not list:
        raise ValueError("Selector must be a list")

    if len(selector) == 0:
        raise ValueError("Empty selector")

    rep = selector[0]
    if type(rep) is int:
        return ColumnId.INDEX
    elif type(rep) is list:
        if len(rep) == 0:
            raise ValueError("Empty selector")

        if type(rep[0]) is str:
            return ColumnId.HEADER
        else:
            raise ValueError("Invalid selector")
    else:
        raise ValueError("Invalid selector")


def get_worksheet(workbook: xl.Workbook, sheet: WorksheetId) -> Worksheet:
    """Get a worksheet from a workbook.

    Args:
        workbook (xl.Workbook):
        sheet (Union[str, int]):

    Raises:
        ValueError: If worksheet is not found.

    Returns:
        xl.worksheet.worksheet.Worksheet:
    """
    if type(sheet) is int and sheet < (workbook.sheets):
        return workbook.worksheets[sheet]
    elif type(sheet) is str and sheet in workbook:
        return workbook[sheet]
    else:
        raise ValueError("Invalid input worksheet")


def canonicalize_db_root_path(db: syre.Database) -> str:
    """Canonicalizes teh database's root path.

    Args:
        db (syre.Database): Database

    Returns:
        str: Canonicalized root path of the database.
    """
    root_path = os.path.realpath(db._root_path)
    if os.name == "nt":
        # windows, ensure UNC
        if not root_path.startswith(UNC_PATH):
            root_path = UNC_PATH + path

    return root_path


def insert_data_from_excel(
    asset: syre.Asset,
    worksheet: Worksheet,
    data_worksheet: WorksheetId,
    data_columns: ColumnSelection,
    header_action: HeaderAction,
    current_column: int,
    data_headers: int = 0,
    skip_rows: int = 0,
    asset_path: Optional[str] = None,
):
    """Insert data into a worksheet from an Excel workbook.

    Args:
        asset (syre.Asset): Asset representing the data resource.
        worksheet (Worksheet): Template worksheet in which to insert the data.
        data_worksheet (WorksheetId): Worksheet id containing the data.
        data_columns (ColumnSelection): List of column headers or labels identifying the data to be copied into the template. [data]
        header_action (HeaderAction): How to label data. [template]
        current_column (int): Current column index of the template manipulation.
        data_headers (int, optional): Number of headers in the data. Defaults to 0.
        skip_rows (int, optional): Number of rows to skip until the first header or data. Defaults to 0.
        asset_path (str): Relative path to the asset file.
            Used to label data if `data-label` is `HeaderAction.INSERT` or `HeaderAction.REPLACE`.
            Defaults to None.
    """
    asset_file_name = os.path.basename(asset.file)
    xl_model = formulas.ExcelModel().loads(asset.file).finish()
    calculated_data = xl_model.calculate()
    input_wb = xl.load_workbook(asset.file, data_only=True)

    if isinstance(data_worksheet, str):
        input_ws = input_wb.get_sheet_by_name(data_worksheet)
    elif isinstance(data_worksheet, int):
        input_ws = input_wb.worksheets[data_worksheet]
    else:
        raise TypeError("Invalid `data_worksheet`")

    context_key = f"'[{asset_file_name}]{input_ws.title.upper()}'"
    worksheet.insert_cols(current_column, len(data_columns))
    for label in data_columns:
        data_row_start = skip_rows + 1
        if (
            header_action is HeaderAction.INSERT
            or header_action is HeaderAction.REPLACE
        ):
            worksheet.cell(row=1, column=current_column).value = asset_path
            data_row_start += 1

        data_cells = input_ws[label]
        if header_action is HeaderAction.REPLACE:
            data_cells = data_cells[data_headers:]

        for row, data_cell in enumerate(data_cells, start=data_row_start):
            template_cell = worksheet.cell(row=row, column=current_column)
            if type(data_cell.value) is str:
                tok = Tokenizer(data_cell.value)
                tokens = tok.items
                if len(tokens) == 0:
                    continue
                elif tokens[0].type == Token.LITERAL:
                    template_cell.value = data_cell.value
                else:
                    calculated_cell = calculated_data.get(
                        f"{context_key}!{data_cell.coordinate}"
                    )
                    template_cell.value = calculated_cell.value[0, 0]
            else:
                template_cell.value = data_cell.value


def insert_data_from_csv(
    asset: syre.Asset,
    worksheet: Worksheet,
    data_selector: ColumnSelection,
    header_action: HeaderAction,
    current_column: int,
    skip_rows: int = 0,
    comment: Optional[str] = None,
    asset_path: Optional[str] = None,
):
    """Insert data into a worksheet from an Excel workbook.

    Args:
        asset (syre.Asset): Asset representing the data resource.
        worksheet (Worksheet): Template worksheet in which to insert the data.
        data_selector (ColumnSelection): List of column headers or labels identifying the data to be copied into the template. [data]
        header_action (HeaderAction): How to construct data headers. [template]
        current_column (int): Current column index of the template manipulation.
        skip_rows (int, optional): Number of rows to skip until the first header or data. Defaults to 0.
        comment (Optional[str (length 1)], optional): Comment character to ignore lines. Defaults to None.
        asset_path (str): Relative path to the asset file.
            Used to label data if `data-label` is `HeaderAction.INSERT` or `HeaderAction.REPLACE`.
            Defaults to None.
    """
    data_selector_type = selection_type(data_selector)
    data = pandas.read_csv(asset.file, skiprows=skip_rows, comment=comment)
    if data_selector_type is ColumnId.INDEX:
        df = data.iloc[:, data_selector]
    elif data_selector_type is ColumnId.HEADER:
        raise NotImplementedError("todo")
        # df = data.loc[:, data_selector]
    else:
        raise RuntimeError("Invalid data selector")

    worksheet.insert_cols(current_column, len(data_selector))
    for col_label, col in df.items():
        data_row_start = 1
        if (
            header_action is HeaderAction.INSERT
            or header_action is HeaderAction.REPLACE
        ):
            worksheet.cell(row=1, column=current_column).value = asset_path
            data_row_start += 1

        if header_action is not HeaderAction.REPLACE:
            if type(col_label) is tuple:
                for row, label in enumerate(col_label, start=data_row_start):
                    worksheet.cell(row=row, column=current_column).value = label

                data_row_start += len(col_label)
            else:
                worksheet.cell(row=data_row_start, column=current_column).value = (
                    col_label
                )
                data_row_start += 1

        for row, value in enumerate(col, start=data_row_start):
            worksheet.cell(row=row, column=current_column, value=value)


def translate_tokenizer_formula(
    tokenizer: Tokenizer,
    replace_range: ReplaceRange,
    column_shift: int,
    header_action: HeaderAction,
    insertion_break_column: int,
):
    """Translates a Tokenizer's formula, adusting for column shifts.

    Args:
        tokenizer (Tokenizer):
        replace_range (int, int): Replaced columns in the template.
        column_shift (int): How many columns data was shifted due to insertions.
        header_action (HeaderAction): How data headers were inserted into the template.
        insertion_break_column (int): Column after the final data insertion column.
            e.g. If the last data was inserted in column B (index 2), this would be column C (index 3).

    Raises:
        ValueError: If formula range is invalid.
    """
    # NB: `column_shift` could be calculated from `replace-range` and `insertion_break_column`,
    # but is passed in for efficiency as this funciton may be called often and the column shift will be constant.
    for token in tokenizer.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            col_start, _, col_end, _ = xl_utils.range_boundaries(token.value)
            if col_end == replace_range[1]:
                # expand range to include new data
                formula_range = token.value.split(":")
                if len(formula_range) == 1:
                    token.value = Translator.translate_col(
                        formula_range[0], column_shift
                    )
                elif len(formula_range) == 2:
                    formula_range[1] = Translator.translate_col(
                        formula_range[1], column_shift
                    )
                    token.value = ":".join(formula_range)

                else:
                    raise ValueError(f"Invalid range `{token.value}`")

            if col_start > replace_range[1]:
                token.value = Translator.translate_range(
                    token.value, rdelta=0, cdelta=column_shift
                )

            if header_action is HeaderAction.INSERT:
                formula_range = token.value.split(":")
                if len(formula_range) == 1:
                    if replace_range[0] <= col_start <= insertion_break_column:
                        try:
                            token.value = Translator.translate_row(formula_range[0], 1)
                        except ValueError:
                            pass
                elif len(formula_range) == 2:
                    if (
                        replace_range[0] <= col_start <= insertion_break_column
                        and replace_range[0] <= col_end <= insertion_break_column
                    ):
                        try:
                            token.value = Translator.translate_row(token.value, 1)
                        except ValueError:
                            pass
                else:
                    raise ValueError(f"Invalid range `{token.value}`")


def translate_worksheet_formulas(
    workbook: xl.Workbook,
    replace_range: ReplaceRange,
    insertion_break_column: int,
    header_action: HeaderAction,
):
    """Translates formulas that reference cells within the replaced ranged.

    Args:
        workbook (xl.Workbook): Template.
        replace_range (ReplaceRange): Replaced columns in the template.
        insertion_break_column (int): Column after the final data insertion column.
            e.g. If the last data was inserted in column B (index 2), this would be column C (index 3).
        column_shift (int): How many columns data was shifted due to insertions.
        header_action (HeaderAction): How data headers were inserted into the template.
    """
    column_shift = utils.column_shift(replace_range[1], insertion_break_column)
    if column_shift == 0 and header_action is not HeaderAction.INSERT:
        return

    for ws in workbook.worksheets:
        for col, column in enumerate(ws.iter_cols(), start=1):
            if replace_range[0] <= col <= insertion_break_column:
                # in replaced range or index
                continue

            for cell in column:
                if type(cell.value) is not str:
                    continue

                tok = Tokenizer(cell.value)
                tokens = tok.items
                if len(tokens) == 0 or tokens[0].type == Token.LITERAL:
                    continue

                translate_tokenizer_formula(
                    tok,
                    replace_range,
                    column_shift,
                    header_action,
                    insertion_break_column,
                )
                cell.value = tok.render()


def main(
    template_path: str,
    worksheet: WorksheetId,
    replace_range: (int, int),
    data_format_type: DataFormatType,
    column_selection: ColumnSelection,
    header_action: HeaderAction,
    output_path: str,
    data_format_args: DataFormatArgs = {},
    asset_filter: AssetFilter = {},
    data_headers: int = 0,
    output_properties: AssetProperties = {},
):
    """Use an Excel file as a template.

    Args:
        template_path (str): Aboslute path to the Excel template.
        worksheet (WorksheetId): Name or index of the worksheet. [template]
        replace_range (int, int): Column index range to replace in the template. [template]
        data_format_type: (DataFormatType): Type of data format to expect.
        column_selection (ColumnSelection): List of column indexes or labels identifying the data to be copied into the template. [data]
        header_action (HeaderAction): How to modify data headers when inserting.
        output_path (str): Relative path to the saved output file.
        data_format_args (DataFormatArgs): Arguments relevant for the expected data format.
        asset_filter (AssetFilter): Filter to search for input data.
        data_headers (int, optional): Number of headers in the data. Defaults to 0.
        output_properties (AssetProperties): Asset properties assigned to the output Asset. Defaults to {}.

    Raises:
        ValueError: If `data_selector` is empty.
        RuntimeError: If `asset_filter` returns no assets.
    """
    if len(column_selection) == 0:
        raise ValueError("Empty data selector")

    db = syre.Database()
    template = xl.load_workbook(filename=template_path)
    ws = get_worksheet(template, worksheet)

    delete_columns_count = replace_range[1] - replace_range[0] + 1
    ws.delete_cols(replace_range[0], delete_columns_count)

    assets = db.find_assets(**asset_filter)
    if len(assets) == 0:
        raise RuntimeError("No matching assets")

    if header_action is HeaderAction.INSERT:
        ws.insert_rows(0)

    db_root_path = canonicalize_db_root_path(db)
    current_col = replace_range[0]
    if data_format_type is DataFormatType.EXCEL_WORKBOOK:
        for asset in assets:
            asset_path = os.path.relpath(asset.file, db_root_path)
            insert_data_from_excel(
                asset,
                ws,
                data_format_args.sheet,
                column_selection,
                header_action,
                current_col,
                data_headers=data_headers,
                skip_rows=data_format_args.skip_rows,
                asset_path=asset_path,
            )
            current_col += 1

    elif data_format_type is DataFormatType.SPREADSHEET:
        for asset in assets:
            asset_path = os.path.relpath(asset.file, db_root_path)
            insert_data_from_csv(
                asset,
                ws,
                column_selection,
                header_action,
                current_col,
                skip_rows=data_format_args.skip_rows,
                comment=data_format_args.comment,
                asset_path=asset_path,
            )
            current_col += 1
    else:
        raise ValueError("Invalid data format type")

    translate_worksheet_formulas(
        template,
        replace_range,
        current_col,
        header_action,
    )

    path = db.add_asset(output_path, **output_properties)
    template.save(path)
