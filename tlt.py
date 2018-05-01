import argparse
import csv
import json
import openpyxl
import os
import sys

DEFAULT_FORMAT_DEFINITION_FILE = "format.json"


def main():
    parser = setup_parser()
    args = parser.parse_args()

    src_file = args.SOURCE
    dst_file = args.dest

    if not os.path.isfile(src_file):
        sys.exit("'{}' does not exist".format(src_file))

    try:
        src = load_original_file(src_file)
    except:
        sys.exit("'{}' is an unsupported file format".format(src_file))

    dst = renderTemplate(src)

    # output
    if dst_file:
        with open(dst_file, "a") as f:
            for line in dst:
                f.write(line + "\n")
    else:
        for line in dst:
            print(line)


def setup_parser():
    parser = argparse.ArgumentParser(
        prog="tlt",
        usage="tlt [<optional arguments>] SOURCE",
        description="TODO: write",
        add_help=True,
    )
    parser.add_argument(
        "SOURCE",
        help="Source file path. Currently supported formats are .csv and .xlsx",
    )
    parser.add_argument(
        "-d",
        "--dest",
        help="Destination file path. If not specified, it is output to stdout.",
    )

    return parser


def load_original_file(file):
    ext = os.path.splitext(file)[-1]  # get a file extention
    if ext == ".csv":
        return load_csv_file(file)
    elif ext == ".xlsx":
        return load_excel_file(file)
    else:
        raise


def load_csv_file(file):
    with open(file, "r") as f:
        return list(csv.reader(f))


def load_excel_file(file):
    book = openpyxl.load_workbook(file, read_only=True, keep_vba=False)
    sheet = book.worksheets[0]

    ret = []
    for cols in sheet.rows:
        ret.append([str(col.value or '') for col in cols])

    return ret


def renderTemplate(src):
    fmt_def = load_format_definition(DEFAULT_FORMAT_DEFINITION_FILE)
    ret = []
    for row in src:
        formatted = fmt_def["templates"]["master-thesis"]
        for i, fmt in enumerate(fmt_def["columns"]):
            cell = row[i]
            label = fmt["label"]
            if "replaces" in fmt:
                for repl in fmt["replaces"]:
                    cell = cell.replace(repl["from"], repl["to"])
            formatted = formatted.replace("{" + label + "}", cell)
        ret.append(formatted.replace("\n", ""))  # trim newline

    return ret


def load_format_definition(file):
    with open(file, "r") as f:
        return json.load(f)


if __name__ == "__main__":
    main()
