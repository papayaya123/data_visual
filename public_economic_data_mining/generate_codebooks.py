#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import re
from collections import Counter
import json
from pathlib import Path
from typing import Iterable, List, Optional, Sequence
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import csv


XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
COLUMN_PRIORITY = {
    "datetime": 6,
    "date": 5,
    "float": 4,
    "integer": 3,
    "boolean": 2,
    "string": 1,
}
TYPE_LABELS = {
    "datetime": "日期時間 (datetime)",
    "date": "日期 (date)",
    "float": "浮點數 (float)",
    "integer": "整數 (integer)",
    "boolean": "布林值 (boolean)",
    "string": "文字 (string)",
    "empty": "空欄 (empty)",
}
SAMPLE_LIMIT = 5
UNIQUE_LIMIT = 1000


def iter_csv_rows(path: Path) -> Iterable[List[str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        for row in reader:
            yield [cell.strip() for cell in row]


def iter_xlsx_rows(path: Path) -> Iterable[List[str]]:
    with ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        ns = {"a": XLSX_NS, "r": REL_NS}
        sheets = workbook.findall("a:sheets/a:sheet", namespaces=ns)
        if not sheets:
            return
        first_sheet = sheets[0]
        rel_id = first_sheet.attrib.get(f"{{{REL_NS}}}id")
        sheet_path = "xl/worksheets/sheet1.xml"
        if rel_id:
            rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rel_ns = {"rel": REL_NS}
            for rel in rels_root.findall("rel:Relationship", namespaces=rel_ns):
                if rel.attrib.get("Id") == rel_id:
                    target = rel.attrib.get("Target", "")
                    if target.startswith("/"):
                        sheet_path = target.lstrip("/")
                    elif target.startswith("xl/"):
                        sheet_path = target
                    else:
                        sheet_path = f"xl/{target}"
                    break
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in shared_root.findall("a:si", namespaces={"a": XLSX_NS}):
                text_parts = []
                for node in si.iterfind(".//a:t", namespaces={"a": XLSX_NS}):
                    text_parts.append(node.text or "")
                shared_strings.append("".join(text_parts))
        with zf.open(sheet_path) as sheet_file:
            for event, elem in ET.iterparse(sheet_file, events=("end",)):
                if elem.tag.endswith("}row"):
                    row_data: List[Optional[str]] = []
                    for cell in elem:
                        if not cell.tag.endswith("}c"):
                            continue
                        ref = cell.attrib.get("r")
                        if ref:
                            idx = column_reference_to_index(ref)
                        else:
                            idx = len(row_data)
                        while len(row_data) <= idx:
                            row_data.append(None)
                        row_data[idx] = read_cell_value(cell, shared_strings)
                    while row_data and row_data[-1] is None:
                        row_data.pop()
                    yield [
                        ("" if cell is None else str(cell).strip()) for cell in row_data
                    ]
                    elem.clear()


COLUMN_RE = re.compile(r"([A-Z]+)")


def column_reference_to_index(ref: str) -> int:
    match = COLUMN_RE.match(ref)
    if not match:
        return len(ref)
    letters = match.group(1)
    total = 0
    for char in letters:
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total - 1


def read_cell_value(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find(f"{{{XLSX_NS}}}v")
    if cell_type == "s" and value_node is not None:
        idx_text = value_node.text or "0"
        try:
            idx = int(idx_text)
        except ValueError:
            return value_node.text or ""
        if 0 <= idx < len(shared_strings):
            return shared_strings[idx]
        return shared_strings[-1] if shared_strings else ""
    if cell_type == "inlineStr":
        text_parts = []
        is_node = cell.find(f"{{{XLSX_NS}}}is")
        if is_node is not None:
            for node in is_node.iterfind(f".//{{{XLSX_NS}}}t"):
                text_parts.append(node.text or "")
        return "".join(text_parts)
    if cell_type == "b":
        if value_node is None:
            return ""
        return "TRUE" if value_node.text == "1" else "FALSE"
    if value_node is None:
        return ""
    return value_node.text or ""


class ColumnStats:
    def __init__(self, name: str, index: int):
        self.name = name or f"Column_{index+1}"
        self.total = 0
        self.non_null = 0
        self.type_counts: Counter[str] = Counter()
        self.notes: set[str] = set()
        self.samples: list[str] = []
        self.unique_values: set[str] = set()
        self.unique_overflow = False
        self.numeric_min: Optional[float] = None
        self.numeric_min_text: Optional[str] = None
        self.numeric_max: Optional[float] = None
        self.numeric_max_text: Optional[str] = None
        self.temporal_min: Optional[dt.datetime] = None
        self.temporal_min_text: Optional[str] = None
        self.temporal_max: Optional[dt.datetime] = None
        self.temporal_max_text: Optional[str] = None

    def observe(self, raw_value: Optional[str]):
        self.total += 1
        text = (raw_value or "").strip()
        if not text:
            return
        self.non_null += 1
        dtype, parsed, note = detect_type(text)
        self.type_counts[dtype] += 1
        if note:
            self.notes.add(note)
        if dtype in {"integer", "float"}:
            value = float(parsed)
            if self.numeric_min is None or value < self.numeric_min:
                self.numeric_min = value
                self.numeric_min_text = text
            if self.numeric_max is None or value > self.numeric_max:
                self.numeric_max = value
                self.numeric_max_text = text
        elif dtype in {"date", "datetime"}:
            assert isinstance(parsed, dt.datetime)
            if self.temporal_min is None or parsed < self.temporal_min:
                self.temporal_min = parsed
                self.temporal_min_text = text
            if self.temporal_max is None or parsed > self.temporal_max:
                self.temporal_max = parsed
                self.temporal_max_text = text
        if text not in self.samples and len(self.samples) < SAMPLE_LIMIT:
            self.samples.append(text)
        if not self.unique_overflow:
            self.unique_values.add(text)
            if len(self.unique_values) > UNIQUE_LIMIT:
                self.unique_values.clear()
                self.unique_overflow = True

    def resolved_type(self) -> str:
        if not self.type_counts:
            return "empty"
        best_type = None
        best_score = (-1, -1)
        for dtype, count in self.type_counts.items():
            priority = COLUMN_PRIORITY.get(dtype, 0)
            candidate = (count, priority)
            if candidate > best_score:
                best_score = candidate
                best_type = dtype
        return best_type or "string"

    def unique_count_display(self) -> str:
        if self.unique_overflow:
            return f">{UNIQUE_LIMIT}"
        return str(len(self.unique_values))

    def min_display(self, dtype: str) -> str:
        if dtype in {"integer", "float"}:
            return self.numeric_min_text or "-"
        if dtype in {"date", "datetime"}:
            return self.temporal_min_text or "-"
        return "-"

    def max_display(self, dtype: str) -> str:
        if dtype in {"integer", "float"}:
            return self.numeric_max_text or "-"
        if dtype in {"date", "datetime"}:
            return self.temporal_max_text or "-"
        return "-"

    def sample_display(self) -> str:
        if not self.samples:
            return "-"
        if len(self.samples) == 1:
            return self.samples[0]
        return "; ".join(self.samples)

    def notes_display(self) -> str:
        notes = list(self.notes)
        notes.sort()
        extra = []
        if self.unique_overflow:
            extra.append(f"唯一值超過 {UNIQUE_LIMIT} 筆，僅顯示下限")
        notes.extend(extra)
        if not notes:
            return ""
        return "；".join(notes)


DATETIME_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M",
    "%Y%m%dT%H%M%S",
]
DATE_FORMATS = [
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y%m%d",
]
INT_RE = re.compile(r"^[+-]?\d+$")
FLOAT_RE = re.compile(r"^[+-]?(?:\d+\.\d*|\.\d+|\d+\.\d+)$")
BOOL_VALUES = {
    "true": True,
    "false": False,
    "yes": True,
    "no": False,
    "y": True,
    "n": False,
    "是": True,
    "否": False,
}


def detect_type(text: str) -> tuple[str, object, Optional[str]]:
    datetime_value, fmt = try_parse_formats(text, DATETIME_FORMATS)
    if datetime_value:
        return "datetime", datetime_value, f"格式 {fmt}"
    date_value, fmt = try_parse_formats(text, DATE_FORMATS)
    if date_value:
        return "date", date_value, f"格式 {fmt}"
    lower = text.lower()
    if lower in BOOL_VALUES:
        return "boolean", BOOL_VALUES[lower], None
    if INT_RE.fullmatch(text):
        try:
            return "integer", int(text), None
        except ValueError:
            pass
    if FLOAT_RE.fullmatch(text):
        try:
            return "float", float(text), None
        except ValueError:
            pass
    return "string", text, None


def try_parse_formats(text: str, formats: Sequence[str]) -> tuple[Optional[dt.datetime], Optional[str]]:
    for fmt in formats:
        try:
            return dt.datetime.strptime(text, fmt), fmt
        except ValueError:
            continue
    return None, None


def analyze_file(path: Path) -> dict:
    if path.suffix.lower() == ".csv":
        row_iter = iter_csv_rows(path)
    elif path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        row_iter = iter_xlsx_rows(path)
    else:
        raise ValueError(f"Unsupported file type: {path}")

    iterator = iter(row_iter)
    try:
        headers = next(iterator)
    except StopIteration:
        return {
            "path": path,
            "headers": [],
            "rows": 0,
            "columns": [],
        }
    normalized_headers = [header.strip() for header in headers]
    stats = [ColumnStats(name, idx) for idx, name in enumerate(normalized_headers)]
    row_count = 0
    for row in iterator:
        row_count += 1
        if len(row) < len(stats):
            row = row + [""] * (len(stats) - len(row))
        elif len(row) > len(stats):
            for idx in range(len(stats), len(row)):
                column_name = f"Column_{idx+1}"
                stats.append(ColumnStats(column_name, idx))
                normalized_headers.append(column_name)
        for idx, value in enumerate(row[: len(stats)]):
            stats[idx].observe(value)
    return {
        "path": path,
        "headers": normalized_headers,
        "rows": row_count,
        "columns": stats,
    }


def format_codebook(report: dict, dataset_alias: Optional[dict], data_prefix: str) -> str:
    path = report["path"]
    rows = report["rows"]
    headers = report["headers"]
    stats: List[ColumnStats] = report["columns"]
    alias = dataset_alias or {}
    display_title = alias.get("title")
    dataset_name = path.name
    if display_title:
        title_line = f"# {display_title}（{dataset_name}）Codebook"
    else:
        title_line = f"# {dataset_name} Codebook"
    display_path = alias.get("path")
    if not display_path:
        prefix = data_prefix.rstrip("/")
        display_path = f"{prefix}/{dataset_name}" if prefix else dataset_name
    column_aliases: dict[str, str] = alias.get("columns", {})
    lines = [
        title_line,
        "",
        f"- 檔案路徑: `{display_path}`",
        f"- 總列數 (不含欄位列): {rows}",
        f"- 欄位數: {len(headers)}",
        "",
        "| 欄位名稱 | 推測型態 | 非空值/總列 | 唯一值數 | 最小值 | 最大值 | 範例值 | 備註 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for stat in stats:
        dtype = stat.resolved_type()
        type_label = TYPE_LABELS.get(dtype, dtype)
        non_null = stat.non_null
        total = stat.total
        unique_display = stat.unique_count_display()
        min_value = stat.min_display(dtype)
        max_value = stat.max_display(dtype)
        samples = stat.sample_display()
        notes = stat.notes_display()
        column_name = column_aliases.get(stat.name, stat.name)
        lines.append(
            f"| {column_name} | {type_label} | {non_null}/{total} | {unique_display} | {min_value} | {max_value} | {samples} | {notes} |"
        )
    return "\n".join(lines) + "\n"


def write_codebook(report: dict, output_path: Path, dataset_alias: Optional[dict], data_prefix: str):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(format_codebook(report, dataset_alias, data_prefix), encoding="utf-8")


def load_alias_config(path: Optional[str]) -> dict:
    if not path:
        return {}
    alias_path = Path(path)
    if not alias_path.exists():
        raise SystemExit(f"找不到欄位別名設定檔: {alias_path}")
    with alias_path.open(encoding="utf-8") as handle:
        data = json.load(handle)
    return data


def main():
    parser = argparse.ArgumentParser(description="產生資料集欄位摘要 (codebook)")
    parser.add_argument("--data-dir", default="data", help="來源資料夾 (預設: data)")
    parser.add_argument("--output-dir", default="codebook", help="輸出資料夾 (預設: codebook)")
    parser.add_argument("--alias-config", help="欄位與標題別名設定檔 (選用)")
    args = parser.parse_args()
    data_dir = Path(args.data_dir)
    output_dir = Path(args.output_dir)
    if not data_dir.exists():
        raise SystemExit(f"資料夾不存在: {data_dir}")
    alias_config = load_alias_config(args.alias_config)
    data_prefix = Path(args.data_dir).as_posix()
    processed = 0
    for dataset in sorted(data_dir.iterdir()):
        if dataset.suffix.lower() not in {".csv", ".xlsx", ".xlsm", ".xls"}:
            continue
        report = analyze_file(dataset)
        target = output_dir / f"{dataset.stem}_codebook.md"
        dataset_alias = alias_config.get(dataset.stem)
        write_codebook(report, target, dataset_alias, data_prefix)
        processed += 1
        print(f"已產生: {target}")
    if processed == 0:
        print("未找到可處理的資料檔案")


if __name__ == "__main__":
    main()
