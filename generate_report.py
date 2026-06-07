from __future__ import annotations

import argparse
import shutil
import warnings
from datetime import date, datetime, time
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

warnings.filterwarnings("ignore", message="pkg_resources is deprecated as an API.*", category=UserWarning)

from docxtpl import DocxTemplate
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
DEFAULT_EXCEL = ROOT / "EXCEL" / "data.xlsx"
DEFAULT_TEMPLATE = ROOT / "TEMPLATE" / "TEMPLATE.docx"
DEFAULT_OUTPUT = ROOT / "OUTPUT" / "thermal_report.docx"
DEFAULT_CHUNK_SIZE = 10

# Temporary report-level defaults. If these headers are later added to Excel,
# non-blank Excel values will override these constants row by row.
DEFAULT_CONSTANTS: dict[str, Any] = {
    "ptu.no": "PTU 09",
    "model": "",
    "rating": "415V",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render one Word report page per Excel row, preserving the source DOCX page objects."
    )
    parser.add_argument("--excel", type=Path, default=DEFAULT_EXCEL, help="Input Excel data file.")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="One-page DOCX template.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Base output path for chunked DOCX parts.")
    parser.add_argument("--sheet", default=None, help="Worksheet name. Defaults to the active sheet.")
    parser.add_argument(
        "--keep-pages",
        action="store_true",
        help="Keep intermediate rendered page DOCX files next to the final output for debugging.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=DEFAULT_CHUNK_SIZE,
        help="Maximum pages per merged output part.",
    )
    args = parser.parse_args()
    if args.chunk_size < 1:
        parser.error("--chunk-size must be at least 1")
    return args


def format_value(header: str, value: Any) -> Any:
    if value is None:
        return ""

    if isinstance(value, datetime):
        if header == "date":
            return value.strftime("%d-%m-%Y")
        if header == "time":
            return value.strftime("%H:%M")
        return value.isoformat(sep=" ")

    if isinstance(value, date) and header == "date":
        return value.strftime("%d-%m-%Y")

    if isinstance(value, time):
        return value.strftime("%H:%M")

    if header == "humidity" and isinstance(value, (int, float)):
        return f"{value:.0%}" if 0 <= value <= 1 else str(value)

    return value


def set_nested(context: dict[str, Any], dotted_key: str, value: Any) -> None:
    parts = dotted_key.split(".")
    target = context
    for part in parts[:-1]:
        current = target.get(part)
        if not isinstance(current, dict):
            current = {}
            target[part] = current
        target = current
    target[parts[-1]] = value


def build_context(row: dict[str, Any]) -> dict[str, Any]:
    context: dict[str, Any] = {}

    for key, value in DEFAULT_CONSTANTS.items():
        set_nested(context, key, value)

    for key, value in row.items():
        if key is None:
            continue
        normalized_key = str(key).strip()
        if not normalized_key:
            continue

        formatted = format_value(normalized_key, value)
        if formatted == "" and normalized_key in DEFAULT_CONSTANTS:
            continue
        set_nested(context, normalized_key, formatted)

    return context


def load_rows(excel_path: Path, sheet_name: str | None) -> list[dict[str, Any]]:
    workbook = load_workbook(excel_path, data_only=True)
    worksheet = workbook[sheet_name] if sheet_name else workbook.active

    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    rows: list[dict[str, Any]] = []

    for row_idx in range(2, worksheet.max_row + 1):
        values = [worksheet.cell(row_idx, col).value for col in range(1, worksheet.max_column + 1)]
        if not any(value is not None for value in values):
            continue
        rows.append(dict(zip(headers, values)))

    if not rows:
        raise ValueError(f"No data rows found in {excel_path}")

    return rows


def render_page(template_path: Path, context: dict[str, Any], output_path: Path) -> None:
    template = DocxTemplate(str(template_path))
    template.render(context)
    template.save(output_path)


def combine_pages_with_word(page_paths: list[Path], output_path: Path) -> None:
    if not page_paths:
        raise ValueError("No rendered pages to combine.")

    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Microsoft Word merge requires pywin32/win32com.") from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    master_path = output_path.parent / f"{output_path.stem}.word_merge_working.docx"
    if master_path.exists():
        master_path.unlink()
    if output_path.exists():
        output_path.unlink()
    shutil.copy2(page_paths[0], master_path)

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    document = None
    try:
        document = word.Documents.Open(str(master_path.resolve()), ReadOnly=False, AddToRecentFiles=False)
        for page_path in page_paths[1:]:
            source_document = None
            insert_range = document.Range(document.Content.End - 1, document.Content.End - 1)
            insert_range.InsertBreak(7)  # wdPageBreak
            insert_range = document.Range(document.Content.End - 1, document.Content.End - 1)
            try:
                source_document = word.Documents.Open(
                    str(page_path.resolve()),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                )
                source_document.Content.Copy()
                insert_range.Paste()
            finally:
                if source_document is not None:
                    try:
                        source_document.Close(SaveChanges=False)
                    except Exception:
                        pass

        document.SaveAs2(str(output_path.resolve()), FileFormat=12)  # wdFormatXMLDocument
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                # Word can disconnect the COM object after SaveAs2 while the file is still saved.
                pass
        try:
            word.Quit()
        except Exception:
            pass
        if master_path.exists():
            master_path.unlink()

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError(f"Word merge did not create a valid output file: {output_path}")


def output_part_path(output_path: Path, part_number: int) -> Path:
    return output_path.with_name(f"{output_path.stem}_part_{part_number:02d}{output_path.suffix}")


def clear_existing_parts(output_path: Path) -> None:
    for existing_part in output_path.parent.glob(f"{output_path.stem}_part_*.docx"):
        existing_part.unlink()


def chunked(values: list[Path], chunk_size: int) -> list[list[Path]]:
    return [values[idx : idx + chunk_size] for idx in range(0, len(values), chunk_size)]


def combine_pages_in_chunks(page_paths: list[Path], output_path: Path, chunk_size: int) -> list[Path]:
    clear_existing_parts(output_path)
    chunks = chunked(page_paths, chunk_size)
    output_paths = []
    for part_number, chunk in enumerate(chunks, start=1):
        part_path = output_part_path(output_path, part_number)
        combine_pages_with_word(chunk, part_path)
        output_paths.append(part_path)
    return output_paths


def generate_report(
    excel_path: Path,
    template_path: Path,
    output_path: Path,
    sheet_name: str | None,
    keep_pages: bool,
    chunk_size: int,
) -> list[Path]:
    rows = load_rows(excel_path, sheet_name)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if keep_pages:
        pages_dir = output_path.with_suffix("")
        if pages_dir.exists():
            shutil.rmtree(pages_dir)
        pages_dir.mkdir(parents=True)
        page_paths = []
        for idx, row in enumerate(rows, start=1):
            page_path = pages_dir / f"page_{idx:03d}.docx"
            render_page(template_path, build_context(row), page_path)
            page_paths.append(page_path)
        return combine_pages_in_chunks(page_paths, output_path, chunk_size)

    with TemporaryDirectory(dir=output_path.parent) as temp_dir:
        temp_path = Path(temp_dir)
        page_paths = []
        for idx, row in enumerate(rows, start=1):
            page_path = temp_path / f"page_{idx:03d}.docx"
            render_page(template_path, build_context(row), page_path)
            page_paths.append(page_path)
        return combine_pages_in_chunks(page_paths, output_path, chunk_size)


def main() -> None:
    args = parse_args()
    output_paths = generate_report(
        excel_path=args.excel,
        template_path=args.template,
        output_path=args.output,
        sheet_name=args.sheet,
        keep_pages=args.keep_pages,
        chunk_size=args.chunk_size,
    )
    print("Generated:")
    for output_path in output_paths:
        print(f"- {output_path}")


if __name__ == "__main__":
    main()
