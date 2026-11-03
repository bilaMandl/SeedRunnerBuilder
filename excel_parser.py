#!/usr/bin/env python3
"""
Import practice questions from Excel workbooks into SQLite.

Excel conventions (columns B–D, 1-based):
  C=$  part of the question (scenario / context)
  C=#  question line (the actual question)
  C=A-D  answer option (D=! marks the correct answer)
  C=*  explanation for the correct answer
  D=1|2  difficulty on the # or $ row (1 -> Easy, 2 -> Hard)
  D=% on the # or $ row  question belongs to the exam category (מבחנים)

Rows where column C is not a known marker and the row contains data
are treated as tables and stored as JSON.
"""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl

MARKERS = {"$", "#", "A", "B", "C", "D", "*"}
ANSWER_KEYS = ("A", "B", "C", "D")
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL = SCRIPT_DIR / "תרגול" / "ביטוח לאומי" / "חישובים.xlsx"
DEFAULT_SCHEMA = SCRIPT_DIR / "init.sql"
DEFAULT_DB = SCRIPT_DIR / "practice.db"

# Default Excel / Windows placeholder names to skip (empty sheets/folders only).
_DEFAULT_SHEET_PATTERNS = (
    re.compile(r"^גיליון\s*\d+$", re.IGNORECASE),
    re.compile(r"^sheet\s*\d+$", re.IGNORECASE),
    re.compile(r"^גליון עבודה של Microsoft Excel חדש(?:\s*\(\d+\))?$", re.IGNORECASE),
)
_DEFAULT_FOLDER_PATTERNS = (
    re.compile(r"^תיקיה חדשה(?:\s*\(\d+\))?$", re.IGNORECASE),
    re.compile(r"^new folder(?:\s*\(\d+\))?$", re.IGNORECASE),
)


def is_default_placeholder_name(name: str, *, kind: str = "sheet") -> bool:
    """Return True for Excel/Windows default names like 'גיליון1' or 'תיקיה חדשה (2)'."""
    text = name.strip()
    if not text:
        return True
    patterns = _DEFAULT_FOLDER_PATTERNS if kind == "folder" else _DEFAULT_SHEET_PATTERNS
    return any(pattern.fullmatch(text) for pattern in patterns)


@dataclass
class ContentPart:
    type: str
    text: str = ""
    table: dict[str, Any] | None = None

    def to_dict(self) -> dict[str, Any]:
        if self.type == "table" and self.table is not None:
            return {"type": "table", "table": self.table}
        return {"type": "text", "text": self.text}


@dataclass
class ParsedQuestion:
    parts: list[ContentPart] = field(default_factory=list)
    question_line: str = ""
    difficulty: str = "Easy"
    category: str = "practice"
    answers: dict[str, str] = field(default_factory=dict)
    correct_key: str | None = None
    explanation_parts: list[str] = field(default_factory=list)
    source_row: int = 0

    @property
    def question_text(self) -> str:
        blocks: list[str] = []
        for part in self.parts:
            if part.type == "text" and part.text.strip():
                blocks.append(part.text.strip())
        if self.question_line.strip():
            blocks.append(self.question_line.strip())
        return "\n\n".join(blocks)

    @property
    def content_json(self) -> str | None:
        payload = [
            part.to_dict()
            for part in self.parts
            if part.type == "table" or (part.type == "text" and part.text.strip())
        ]
        if not payload:
            return None
        return json.dumps(payload, ensure_ascii=False)

    @property
    def explanation(self) -> str:
        return "\n".join(p.strip() for p in self.explanation_parts if p and p.strip())

    def is_complete(self) -> bool:
        if not self.question_text or not self.question_line.strip():
            return False
        if len(self.answers) != 4:
            return False
        if not all(self.answers.get(k, "").strip() for k in ANSWER_KEYS):
            return False
        return self.correct_key in ANSWER_KEYS


def cell_str(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def marker(value: Any) -> str:
    return cell_str(value)


def row_text(ws, row: int) -> str:
    prefix = cell_str(ws.cell(row, 1).value)
    body = cell_str(ws.cell(row, 2).value)
    if prefix and body:
        return f"{prefix} {body}".strip()
    return body or prefix


def row_cells(ws, row: int, max_col: int) -> list[str]:
    return [cell_str(ws.cell(row, col).value) for col in range(1, max_col + 1)]


def is_content_row(cells: list[str]) -> bool:
    return any(cell for cell in cells)


def trim_trailing_empty(cells: list[str]) -> list[str]:
    trimmed = list(cells)
    while trimmed and not trimmed[-1]:
        trimmed.pop()
    return trimmed


def is_table_row(ws, row: int, max_col: int) -> bool:
    c = marker(ws.cell(row, 3).value)
    if c in MARKERS:
        return False
    cells = row_cells(ws, row, max_col)
    if not is_content_row(cells):
        return False
    if c:
        return True
    filled = sum(1 for cell in cells if cell)
    return filled >= 2


def parse_cell_value(value: str) -> str | int | float:
    text = value.strip()
    if not text:
        return ""
    normalized = text.replace("\u00a0", "").replace("₪", "").replace(",", "").strip()
    if re.fullmatch(r"-?\d+", normalized):
        return int(normalized)
    if re.fullmatch(r"-?\d+\.\d+", normalized):
        return float(normalized)
    return text


def _is_percentage_format(number_format: str) -> bool:
    return bool(re.search(r"0(\.0+)?%", number_format or ""))


def _format_percentage(value: float, number_format: str) -> str:
    pct = value * 100
    match = re.search(r"0\.(0+)%", number_format or "")
    if match:
        decimals = len(match.group(1))
        text = f"{pct:.{decimals}f}".rstrip("0").rstrip(".")
        return f"{text}%"
    if abs(pct - round(pct)) < 1e-9:
        return f"{int(round(pct))}%"
    text = f"{pct:.10f}".rstrip("0").rstrip(".")
    return f"{text}%"


def parse_excel_cell(cell) -> str | int | float:
    value = cell.value
    if value is None:
        return ""
    number_format = str(cell.number_format or "")
    if isinstance(value, (int, float)) and _is_percentage_format(number_format):
        return _format_percentage(float(value), number_format)
    if isinstance(value, (int, float)) and not number_format.strip().endswith("%"):
        # Plain numeric cell – keep as number when not explicitly text.
        if float(value).is_integer():
            return int(value)
        return float(value)
    return parse_cell_value(cell_str(value))


def normalize_table_value(value: str | int | float) -> str | int | float:
    if isinstance(value, str):
        return parse_cell_value(value)
    return value


def strip_leading_empty_columns(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return rows
    while rows and all(not row[0] for row in rows):
        rows = [row[1:] for row in rows]
    return rows


def strip_marker_columns(row: list[str]) -> list[str]:
    row = trim_trailing_empty(row)
    while row and row[-1] in {"$", "%"}:
        row = row[:-1]
    return row


def table_rows_to_json(raw_rows: list[list[str]]) -> dict[str, Any]:
    rows = [strip_marker_columns(row) for row in raw_rows]
    rows = [row for row in rows if any(cell for cell in row)]
    rows = strip_leading_empty_columns(rows)
    if not rows:
        return {"columnHeaders": [], "rows": []}

    if len(rows) >= 1 and all(len(r) == 2 for r in rows) and all(_looks_like_row_label(r[0]) for r in rows):
        return {
            "columnHeaders": [""],
            "rows": [{"header": row[0], "cells": [parse_cell_value(row[1])]} for row in rows],
        }

    max_width = max(len(row) for row in rows)
    rows = [row + [""] * (max_width - len(row)) for row in rows]

    header_row = rows[0]
    data_rows = rows[1:]
    column_headers = [normalize_table_value(cell) for cell in header_row]

    if not data_rows:
        return {"columnHeaders": column_headers, "rows": []}

    parsed_rows: list[list[str | int | float]] = []
    for row in data_rows:
        row = trim_trailing_empty(row)
        if not row:
            continue
        cells = row[: len(header_row)]
        while len(cells) < len(header_row):
            cells.append("")
        parsed_rows.append([normalize_table_value(cell) for cell in cells])

    return {"columnHeaders": column_headers, "rows": parsed_rows}


def _looks_like_row_label(value: str) -> bool:
    text = value.strip()
    if not text:
        return False
    if re.search(r"\d", text):
        return False
    if "₪" in text or "%" in text:
        return False
    return True


def _data_rows_have_side_headers(data_rows: list[list[str]]) -> bool:
    candidates = [row for row in data_rows if row and row[0].strip()]
    if not candidates:
        return False
    if len(candidates[0]) < 2:
        return False
    return all(
        _looks_like_row_label(row[0]) and any(cell.strip() for cell in row[1:])
        for row in candidates
    )


def map_difficulty(value: Any) -> str:
    text = cell_str(value)
    if text in {"2", "2.0"}:
        return "Hard"
    if text in {"1", "1.0"}:
        return "Easy"
    return "Easy"


def is_exam_flag(value: str) -> bool:
    return value == "%"


def is_intro_text_row(ws, row: int) -> bool:
    a = cell_str(ws.cell(row, 1).value)
    b = cell_str(ws.cell(row, 2).value)
    c = marker(ws.cell(row, 3).value)
    return a.isdigit() and bool(b) and c not in MARKERS


def intro_row_table_header(ws, row: int) -> list[str] | None:
    """Intro row with column headers in C–F (employer table layout)."""
    if not is_intro_text_row(ws, row):
        return None
    header_values = [cell_str(ws.cell(row, col).value) for col in (3, 4, 5, 6)]
    if not any(header_values[1:]):
        return None
    numeric_headers = sum(1 for value in header_values[1:] if value and re.search(r"\d", value))
    if numeric_headers >= 2:
        return None
    return header_values


def is_employer_table_row(ws, row: int) -> bool:
    """Table row with employer name in C and values in D–F."""
    if cell_str(ws.cell(row, 1).value) or cell_str(ws.cell(row, 2).value):
        return False
    label = cell_str(ws.cell(row, 3).value)
    if not label or label in MARKERS:
        return False
    return bool(cell_str(ws.cell(row, 4).value))


def employer_table_row_cells(ws, row: int) -> list[str | int | float]:
    return [parse_excel_cell(ws.cell(row, col)) for col in (3, 4, 5, 6)]


def table_row_values(ws, row: int, max_col: int) -> list[str | int | float]:
    return [parse_excel_cell(ws.cell(row, col)) for col in range(1, max_col + 1)]


def apply_question_flags(question: ParsedQuestion, d_str: str) -> None:
    if is_exam_flag(d_str):
        question.category = "exam"
    elif d_str in {"1", "1.0", "2", "2.0"}:
        question.difficulty = map_difficulty(d_str)


def parse_sheet(ws) -> tuple[list[ParsedQuestion], list[str]]:
    max_col = max(ws.max_column, 4)
    questions: list[ParsedQuestion] = []
    warnings: list[str] = []

    current: ParsedQuestion | None = None
    table_buffer: list[list[str]] = []
    pending_table_buffer: list[list[str]] = []
    pending_text_parts: list[str] = []
    after_question_line = False

    def flush_table() -> None:
        nonlocal table_buffer
        if current is not None and table_buffer:
            table_json = table_rows_to_json(table_buffer)
            current.parts.append(ContentPart(type="table", table=table_json))
        table_buffer = []

    def flush_pending_into(question: ParsedQuestion) -> None:
        nonlocal pending_table_buffer, pending_text_parts
        for text in pending_text_parts:
            question.parts.append(ContentPart(type="text", text=text))
        pending_text_parts = []
        if pending_table_buffer:
            table_json = table_rows_to_json(pending_table_buffer)
            question.parts.append(ContentPart(type="table", table=table_json))
            pending_table_buffer = []

    def flush_question() -> None:
        nonlocal current, after_question_line
        flush_table()
        if current is not None:
            if current.is_complete():
                questions.append(current)
            else:
                warnings.append(
                    f"Sheet '{ws.title}' row {current.source_row}: incomplete question skipped"
                )
        current = None
        after_question_line = False

    def start_question(row: int) -> ParsedQuestion:
        nonlocal current
        if current is not None and current.is_complete():
            flush_question()
        if current is None:
            current = ParsedQuestion(source_row=row)
            flush_pending_into(current)
        return current

    for row in range(1, ws.max_row + 1):
        c = marker(ws.cell(row, 3).value)
        d = ws.cell(row, 4).value
        d_str = cell_str(d)

        if is_intro_text_row(ws, row):
            flush_table()
            if current is not None and current.is_complete():
                flush_question()
            pending_text_parts.append(row_text(ws, row))
            table_header = intro_row_table_header(ws, row)
            if table_header:
                target_buffer = pending_table_buffer if current is None else table_buffer
                target_buffer.append(table_header)
            continue

        if is_employer_table_row(ws, row):
            row_data = employer_table_row_cells(ws, row)
            if current is not None and current.is_complete():
                flush_question()
            if current is None:
                pending_table_buffer.append(row_data)
                continue
            if not after_question_line:
                table_buffer.append(row_data)
                continue
            warnings.append(
                f"Sheet '{ws.title}' row {row}: table row after question line ignored"
            )
            continue

        if is_table_row(ws, row, max_col):
            row_data = table_row_values(ws, row, max_col)
            if current is not None and current.is_complete():
                flush_question()
            if current is None:
                pending_table_buffer.append(row_data)
                continue
            if not after_question_line:
                table_buffer.append(row_data)
                continue
            warnings.append(
                f"Sheet '{ws.title}' row {row}: table row after question line ignored"
            )
            continue

        flush_table()

        if c == "$":
            question = start_question(row)
            text = row_text(ws, row)
            if text:
                question.parts.append(ContentPart(type="text", text=text))
            apply_question_flags(question, d_str)
            continue

        if c == "#":
            if current is not None and current.is_complete():
                flush_question()
            question = start_question(row)
            question.question_line = row_text(ws, row)
            apply_question_flags(question, d_str)
            after_question_line = True
            continue

        if c in ANSWER_KEYS:
            if current is None or not after_question_line:
                warnings.append(f"Sheet '{ws.title}' row {row}: answer outside question block")
                continue
            answer_text = row_text(ws, row)
            if c in current.answers:
                warnings.append(f"Sheet '{ws.title}' row {row}: duplicate answer '{c}'")
            current.answers[c] = answer_text
            if d_str == "!":
                if current.correct_key is not None:
                    warnings.append(f"Sheet '{ws.title}' row {row}: multiple correct answers")
                current.correct_key = c
            continue

        if c == "*":
            if current is None:
                warnings.append(f"Sheet '{ws.title}' row {row}: explanation outside question block")
                continue
            text = row_text(ws, row)
            if text:
                current.explanation_parts.append(text)
            continue

        if not c and not is_content_row(row_cells(ws, row, max_col)):
            continue

        if c and c not in MARKERS:
            if current is not None and not after_question_line:
                text = row_text(ws, row)
                if text:
                    current.parts.append(ContentPart(type="text", text=text))
            else:
                warnings.append(f"Sheet '{ws.title}' row {row}: unrecognized marker '{c}'")

    flush_question()
    return questions, warnings


def parse_workbook(path: Path) -> dict[str, list[ParsedQuestion]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    parsed: dict[str, list[ParsedQuestion]] = {}
    for sheet_name in wb.sheetnames:
        parsed[sheet_name] = parse_sheet(wb[sheet_name])[0]
    return parsed


def parsed_to_seed_item(question: ParsedQuestion) -> dict[str, Any]:
    return {
        "question": question.question_text,
        "answers": [question.answers[key] for key in ANSWER_KEYS],
        "correct": ANSWER_KEYS.index(question.correct_key),
        "explanation": question.explanation,
        "difficulty": question.difficulty,
        "category": question.category,
        "content": json.loads(question.content_json) if question.content_json else [],
    }


def parse_excel_to_seed(path: Path) -> dict[str, list[dict[str, Any]]]:
    workbook = parse_workbook(path)
    return {
        sheet_name: [parsed_to_seed_item(question) for question in questions]
        for sheet_name, questions in workbook.items()
    }


def init_database(conn: sqlite3.Connection, schema_path: Path) -> None:
    conn.executescript(schema_path.read_text(encoding="utf-8"))


def get_or_create_topic(
    conn: sqlite3.Connection,
    cache: dict[tuple[str | None, str], int],
    name: str,
    parent_id: int | None,
    level: int,
) -> int:
    key = (name, parent_id)
    if key in cache:
        return cache[key]
    conn.execute(
        "INSERT INTO topics (name, parent_id, level) VALUES (?, ?, ?)",
        (name, parent_id, level),
    )
    topic_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    cache[key] = topic_id
    return topic_id


def _questions_columns(conn: sqlite3.Connection) -> set[str]:
    return {column[1] for column in conn.execute("PRAGMA table_info(questions)").fetchall()}


def insert_questions(
    conn: sqlite3.Connection,
    topic_id: int,
    questions: list[ParsedQuestion],
) -> int:
    columns = _questions_columns(conn)
    inserted = 0
    for question in questions:
        values: dict[str, Any] = {
            "topic_id": topic_id,
            "question_text": question.question_text,
            "difficulty": question.difficulty,
        }
        field_names = ["topic_id", "question_text", "difficulty"]
        if "category" in columns:
            values["category"] = question.category
            field_names.append("category")
        if "explanation" in columns:
            values["explanation"] = question.explanation or None
            field_names.append("explanation")
        if "content_json" in columns:
            values["content_json"] = question.content_json
            field_names.append("content_json")

        placeholders = ", ".join("?" for _ in field_names)
        conn.execute(
            f"INSERT INTO questions ({', '.join(field_names)}) VALUES ({placeholders})",
            tuple(values[name] for name in field_names),
        )
        question_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        for key in ANSWER_KEYS:
            conn.execute(
                """
                INSERT INTO answers (question_id, text, is_correct)
                VALUES (?, ?, ?)
                """,
                (
                    question_id,
                    question.answers[key],
                    1 if question.correct_key == key else 0,
                ),
            )
        inserted += 1
    return inserted


def import_excel(
    excel_path: Path,
    db_path: Path,
    schema_path: Path,
    subject_name: str | None = None,
    reset: bool = True,
) -> dict[str, int | list[str]]:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    if not schema_path.exists():
        raise FileNotFoundError(f"Schema file not found: {schema_path}")

    if reset and db_path.exists():
        db_path.unlink()

    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")

    try:
        init_database(conn, schema_path)
        topic_cache: dict[tuple[str | None, str], int] = {}

        subject = subject_name or excel_path.parent.name
        file_topic = excel_path.stem

        subject_id = get_or_create_topic(conn, topic_cache, subject, None, 1)
        file_id = get_or_create_topic(conn, topic_cache, file_topic, subject_id, 2)

        workbook = parse_workbook(excel_path)
        stats: dict[str, int | list[str]] = {
            "topics": 2 + len(workbook),
            "questions": 0,
            "answers": 0,
            "warnings": [],
        }

        for sheet_name, questions in workbook.items():
            sheet_id = get_or_create_topic(conn, topic_cache, sheet_name, file_id, 3)
            count = insert_questions(conn, sheet_id, questions)
            stats["questions"] = int(stats["questions"]) + count
            stats["answers"] = int(stats["answers"]) + count * 4

        conn.commit()
        return stats
    finally:
        conn.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Import Excel questions into SQLite")
    parser.add_argument(
        "--excel",
        type=Path,
        default=DEFAULT_EXCEL,
        help="Path to the Excel workbook",
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=DEFAULT_DB,
        help="Output SQLite database path",
    )
    parser.add_argument(
        "--schema",
        type=Path,
        default=DEFAULT_SCHEMA,
        help="Path to init.sql schema",
    )
    parser.add_argument(
        "--subject",
        type=str,
        default=None,
        help="Top-level topic name (defaults to the Excel parent folder name)",
    )
    parser.add_argument(
        "--keep-db",
        action="store_true",
        help="Do not delete an existing database file before import",
    )
    return parser


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    args = build_parser().parse_args()
    stats = import_excel(
        excel_path=args.excel,
        db_path=args.db,
        schema_path=args.schema,
        subject_name=args.subject,
        reset=not args.keep_db,
    )

    print(f"Database: {args.db}")
    print(f"Topics:   {stats['topics']}")
    print(f"Questions:{stats['questions']}")
    print(f"Answers:  {stats['answers']}")

    warnings = stats.get("warnings") or []
    if warnings:
        print("\nWarnings:")
        for warning in warnings:
            print(f"  - {warning}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
