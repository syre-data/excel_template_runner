import os
import openpyxl.utils as xl_utils
from .types import (
    DataFormatType,
    Metadata,
    MetadataArgs,
    ColumnSelection,
    DataSelectionArgs,
    HeaderAction,
)


def parse_data_format_type_arg(arg: str) -> DataFormatType:
    """Parses a string in to a data format type.

    Args:
        arg (str): String to parse.

    Raises:
        ValueError: If the string is not valid.

    Returns:
        DataFormatType: Corresponding data format type.
    """
    arg = arg.lower()
    if arg == "spreadsheet":
        return DataFormatType.SPREADSHEET
    elif arg == "excel":
        return DataFormatType.EXCEL_WORKBOOK

    raise ValueError("Invalid data format type string")


def parse_header_action(arg: str) -> HeaderAction:
    """Parses a string in to a header action.

    Args:
        arg (str): String to parse.

    Raises:
        ValueError: If the string is not valid.

    Returns:
        HeaderAction: Corresponsing header action.
    """
    arg = arg.lower()
    if arg == "none":
        return HeaderAction.NONE
    elif arg == "insert":
        return HeaderAction.INSERT
    elif arg == "replace":
        return HeaderAction.REPLACE

    raise ValueError("Invalid header action string")


def parse_metadata_args(args: MetadataArgs) -> Metadata:
    """Parses argument from command line input into a dictionary of metadata.
    Nested metadata should use dot (.) notation.

    Args:
        args (MetadataArgs): Metadata arguments.

    Returns:
        (Metadata)
    """
    metadata = {}
    for value in args:
        key, val = value.split("=")
        try:
            val = int(val)
        except ValueError:
            val = float(val)

        metadata[key] = val

    return metadata


def parse_column_selection_args(args: DataSelectionArgs) -> ColumnSelection:
    """Parses argument from command line interface as a column selection.

    Args:
        args (DataSelectionArgs): Column selection arguments.

    Raises:
        ValueError: If argument could not be parsed.

    Returns:
        ColumnSelection: Column selection.
    """
    try:
        return [int(arg) for arg in args]
    except ValueError:
        pass

    try:
        return list(map(xl_utils.column_index_from_string, args))
    except ValueError:
        pass

    headers = [list(map(str.strip, arg.split(","))) for arg in args]
    header_lengths = [len(header) for header in headers]
    if not all([length == header_lengths[0] for length in header_lengths]):
        raise ValueError("Could not parse data selection")

    return headers


def is_excel_file(path: str) -> bool:
    if os.path.splitext(path)[1] == ".xlsx":
        return True

    return False


def column_shift(replace_range_end: int, insertion_break_column: int) -> int:
    """Calculates how many columns the input data shifted the replace range.

    e.g. If the replace range was [B, C] and the input data took 4 columns
    -- i.e. [B, E], `insertion_break_column` = 6 -- the columns shift was 2.

    e.g. If the replace range was [B, D] and the input data took 2 columns
    -- i.e. [B, C], `insertion_break_column` = 4 -- the columns shift was -1.

    Args:
        replace_range_end (int): End column of the replace range.
        insertion_break_column (int): Column after the final data insertion column.
            e.g. If the last data was inserted in column B (index 2), this would be column C (index 3).

    Returns:
        int: Number of columns that data shifted.
    """
    return insertion_break_column - replace_range_end - 1
