import argparse
from lib.script import (
    handle_excel_command, 
)

def main() -> None:
    parser = argparse.ArgumentParser(description="Automation tools CLI")
    subparser = parser.add_subparsers(dest="command", help="Available commands")

    read_parser = subparser.add_parser("read", help="Read an Excel file")
    read_parser.add_argument("-d", "--file-path", type=str, help="Path to the Excel file")
    read_parser.add_argument("-s", "--sheet-name", type=str, help="Name of the sheet to read", default=None)

    read_parser.add_argument("-f","--from", dest="cell_from", type=str, help="Start cell (e.g., A1)")
    read_parser.add_argument("-t","--to", dest="cell_to", type=str, help="End cell (e.g., D10)")
    read_parser.add_argument("-o","--output", dest="output_cell", type=str, nargs=2, help="Two output start cell (e.g., A1, B1)")
    read_parser.add_argument("-os","--output-sheet-name", dest="output_sheet_name", type=str, help="Name of the sheet to write output", default=None)

    # add headless option
    read_parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")


    args = parser.parse_args()

    match args.command:
        case "read":
            if bool(args.cell_from) ^ bool(args.cell_to):
                parser.error("Both --from and --to must be provided together.")
            if not bool(args.output_cell):
                parser.error("--output cells must be provided")

            print(f"File path: {args.file_path}")
            print(f"Sheet name: {args.sheet_name if args.sheet_name else 'Active Sheet'}")
            print(f"Cell range: {args.cell_from}:{args.cell_to}")
            print(f"Output start from: {args.output_cell}")
            print(f"Output sheet name: {args.output_sheet_name if args.output_sheet_name else 'Active Sheet'}\n")

            result = handle_excel_command(args.file_path, (args.cell_from, args.cell_to), args.output_cell, args.sheet_name, args.output_sheet_name, args.headless)
            if result is not None:
                print(result)
        case _:
            parser.print_help()


if __name__ == "__main__":
    main()