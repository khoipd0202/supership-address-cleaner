from __future__ import annotations

import argparse
from pathlib import Path

from .cleaner import AddressCleaner


def _default_output(input_path: str) -> str:
    path = Path(input_path)
    return str(path.with_name(f"{path.stem}_cleaned.xlsx"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="vn-address-clean",
        description="Clean SuperShip-style Vietnamese address Excel files.",
    )
    parser.add_argument("input", help="Input .xlsx file")
    parser.add_argument("-o", "--output", help="Output .xlsx file")
    parser.add_argument(
        "--include-empty-rows",
        action="store_true",
        help="Keep rows without POI/street/level4 instead of filtering them out.",
    )
    parser.add_argument(
        "--combined-row",
        action="store_true",
        help="Keep POI/street/level4 in one output row instead of splitting them into separate rows.",
    )
    parser.add_argument("--sheet-name", help="Worksheet name. Defaults to active sheet.")
    args = parser.parse_args(argv)

    output = args.output or _default_output(args.input)
    stats = AddressCleaner().clean_excel(
        args.input,
        output,
        include_empty_rows=args.include_empty_rows,
        split_components=not args.combined_row,
        sheet_name=args.sheet_name,
    )
    print(f"Input rows:  {stats.input_n}")
    print(f"Output rows: {stats.output_n}")
    print(f"Removed:     {stats.removed}")
    print(f"Saved:       {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
