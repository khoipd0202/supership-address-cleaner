from __future__ import annotations

import argparse
from pathlib import Path

from .cleaner import AddressCleaner


def _default_output(input_path: str) -> str:
    path = Path(input_path)
    return str(path.with_name(f"{path.stem}_cleaned.xlsx"))


def main(argv: list[str] | None = None) -> int:
    try:
        from dotenv import load_dotenv
    except ImportError:
        pass
    else:
        load_dotenv()

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
    parser.add_argument(
        "--cerebras",
        action="store_true",
        help="Enable Cerebras LLM parser for ambiguous rows.",
    )
    parser.add_argument(
        "--cerebras-model",
        help="Cerebras model to use (default: gpt-oss-120b).",
    )
    parser.add_argument(
        "--cerebras-api-key",
        help="Cerebras API Key (overrides CEREBRAS_API_KEY environment variable).",
    )
    parser.add_argument(
        "--queue-all",
        action="store_true",
        help="Hàng chờ tuần tự: xử lý TOÀN BỘ dòng mơ hồ qua LLM (không cap), "
             "tự giãn nhịp theo rate limit và backoff khi bị chặn. "
             "Dùng cho file lớn (5k-10k dòng). Ngụ ý --cerebras.",
    )
    args = parser.parse_args(argv)

    output = args.output or _default_output(args.input)
    stats = AddressCleaner(
        use_cerebras=args.cerebras or args.queue_all,
        cerebras_api_key=args.cerebras_api_key,
        cerebras_model=args.cerebras_model,
        queue_all=args.queue_all,
    ).clean_excel(
        args.input,
        output,
        include_empty_rows=args.include_empty_rows,
        split_components=not args.combined_row,
        sheet_name=args.sheet_name,
    )
    print(f"Tổng dòng đầu vào:    {stats.input_n}")
    print(f"Dòng trùng lặp đã bỏ: {stats.duplicates}")
    print(f"Dòng xuất ra:         {stats.output_n}")
    print(f"Dòng bị loại:         {stats.removed}")
    print(f"Dòng cần kiểm tra:    {stats.review_n}")
    print(f"Đã lưu:               {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
