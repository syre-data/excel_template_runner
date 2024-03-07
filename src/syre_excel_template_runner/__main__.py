import argparse
from .types import DataFormatType, HeaderAction, ExcelDataArgs, SpreadsheetDataArgs
from . import utils
from . import excel_template_runner

parser = argparse.ArgumentParser(
    prog="Syre Excel Template Runner",
    description="Runs Excel templates for a Syre project.",
)

# required args
parser.add_argument("template", help="Path to the template.")

parser.add_argument(
    "worksheet", help="Worksheet id for the template where data should be inserted."
)

parser.add_argument(
    "--replace-start",
    type=int,
    required=True,
    help="Start of the range (0-based index) where input data should replace.",
)

parser.add_argument(
    "--replace-end",
    type=int,
    required=True,
    help="End of the range (0-based index) where data should replace.",
)

parser.add_argument(
    "--data-format-type",
    choices=["spreadsheet", "excel"],
    required=True,
    help="The type of data format to expect. `spreadsheet` indicates a simple spreadsheet format (e.g. CSV). `excel` indicates an Excel workbook.",
)

parser.add_argument(
    "--data-columns",
    nargs="+",
    required=True,
    help="Columns within each input Asset that should be copied into the template.\
            Columns may be headers, labels, or indices.\
            For headers use the syntax h1l1,h1l2 h2l1,h2l2.\
            For labels use the syntax A B C.\
            For indices (0-based) use the syntax 0 1 2.",
)

parser.add_argument(
    "--header-action",
    choices=["none", "insert", "replace"],
    default="none",
    required=True,
    help="How to insert header information. `none` does not modify headers at all. `insert` appends the asset path as an additional header above the rest. `replace` will replace the data headers with the asset path.",
)

parser.add_argument(
    "--output", required=True, help="Relative path for the output Asset file."
)

# data format args
# --- spreadsheet
parser.add_argument(
    "--comment-character",
    help="Comment character to skip lines. Only available if `--data-format-type=spreadsheet`.",
)

# --- excel
parser.add_argument(
    "--excel-sheet",
    help="Worksheet id. Required for `--data-format-type=excel`. May be a 0-based index or a string.",
)

# optional args
parser.add_argument(
    "--skip-rows", type=int, default=0, help="How many rows to skip when reading data."
)

# asset filter
parser.add_argument("--filter-name", help="Asset name filter for input data.")

parser.add_argument("--filter-type", help="Asset type filter for input data.")

parser.add_argument(
    "--filter-tags", nargs="*", default=[], help="Asset tags filter for input data."
)

parser.add_argument(
    "--filter-metadata",
    action="append",
    help="Asset metadata filter for input data. Nested metadata is incdicated using dot (.) notation.",
)

# output asset
parser.add_argument("--output-name", help="Asset name for the output data.")
parser.add_argument("--output-type", help="Asset type for the output data.")
parser.add_argument(
    "--output-tags", nargs="*", default=[], help="Asset tags for the output data."
)
parser.add_argument(
    "--output-metadata",
    action="append",
    help="Asset metadata for the output data. Nested metadata is incdicated using dot (.) notation.",
)

# parse
args = parser.parse_args()
if args.replace_start < 0:
    raise ValueError("`--replace-start` must be non-negative.")

if args.replace_end < 0:
    raise ValueError("`--replace-end` must be non-negative.")

replace_range = (args.replace_start, args.replace_end)

data_format_type = utils.parse_data_format_type_arg(args.data_format_type)
if data_format_type is DataFormatType.SPREADSHEET:
    data_format_args = SpreadsheetDataArgs(
        skip_rows=args.skip_rows, comment=args.comment_character
    )
elif data_format_type is DataFormatType.EXCEL_WORKBOOK:
    if args.excel_sheet is None:
        parser.error("--excel-sheet is required if --data-format-type=excel")
    try:
        excel_sheet = int(args.excel_sheet)
    except:
        excel_sheet = args.excel_sheet

    data_format_args = ExcelDataArgs(sheet=excel_sheet, skip_rows=args.skip_rows)
else:
    parser.error(
        "Could not parse value of `--data-format-type` into a data format type."
    )

data_columns = utils.parse_column_selection_args(args.data_columns)
header_action = utils.parse_header_action(args.header_action)
asset_filter = {
    "name": args.filter_name,
    "type": args.filter_type,
    "tags": args.filter_tags,
}

if args.filter_metadata is not None:
    metadata = utils.parse_metadata_args(args.filter_metadata)
    asset_filter["metadata"] = metadata

output_properties = {
    "name": args.output_name,
    "type": args.output_type,
    "tags": args.output_tags,
}

if args.output_metadata is not None:
    metadata = utils.parse_metadata_args(args.output_metadata)
    output_properties["metadata"] = metadata

excel_template_runner.main(
    args.template,
    args.worksheet,
    replace_range,
    data_format_type,
    data_columns,
    header_action,
    args.output,
    data_format_args=data_format_args,
    asset_filter=asset_filter,
    output_properties=output_properties,
)
