from __future__ import annotations

from pathlib import Path
from pprint import pformat
from typing import Any
import ast

from excel_parser import parse_excel_to_seed


def find_workbooks(folder: Path) -> list[Path]:
    if not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")
    return sorted(
        path for path in folder.rglob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    )


def build_seed(folder: Path) -> tuple[dict[str, Any], list[Path], list[str]]:
    seed: dict[str, Any] = {}
    errors: list[str] = []
    workbooks = find_workbooks(folder)

    for workbook in workbooks:
        try:
            sheets = parse_excel_to_seed(workbook)
        except Exception as exc:
            errors.append(f"{workbook.name}: {exc}")
            continue

        subject = workbook.parent.name
        subject_seed = seed.setdefault(subject, {})
        subject_seed[workbook.stem] = sheets

    return seed, workbooks, errors


RUNNER_TEMPLATE = '''from __future__ import annotations

import json
import sqlite3
from typing import Any

SEED: dict[str, dict[str, dict[str, list[dict[str, Any]]]]] = __SEED__


def _questions_columns(conn: sqlite3.Connection) -> set[str]:
    return {row[1] for row in conn.execute("PRAGMA table_info(questions)").fetchall()}


def run_seed(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    columns = _questions_columns(conn)

    for subject_name, files in SEED.items():
        cur.execute(
            "INSERT INTO topics (name, parent_id, level) VALUES (?, NULL, 1)",
            (subject_name,),
        )
        subject_id = cur.lastrowid

        for file_name, sheets in files.items():
            cur.execute(
                "INSERT INTO topics (name, parent_id, level) VALUES (?, ?, 2)",
                (file_name, subject_id),
            )
            file_id = cur.lastrowid

            for sheet_name, questions in sheets.items():
                cur.execute(
                    "INSERT INTO topics (name, parent_id, level) VALUES (?, ?, 3)",
                    (sheet_name, file_id),
                )
                sheet_id = cur.lastrowid

                for question in questions:
                    content = question.get("content") or []
                    content_json = json.dumps(content, ensure_ascii=False) if content else None
                    values = {
                        "topic_id": sheet_id,
                        "question_text": question["question"],
                        "difficulty": question.get("difficulty", "Easy"),
                    }
                    names = ["topic_id", "question_text", "difficulty"]
                    if "category" in columns:
                        values["category"] = question.get("category", "practice")
                        names.append("category")
                    if "explanation" in columns:
                        values["explanation"] = (question.get("explanation") or "").strip() or None
                        names.append("explanation")
                    if "content_json" in columns:
                        values["content_json"] = content_json
                        names.append("content_json")

                    placeholders = ", ".join("?" for _ in names)
                    cur.execute(
                        f"INSERT INTO questions ({", ".join(names)}) VALUES ({placeholders})",
                        tuple(values[name] for name in names),
                    )
                    question_id = cur.lastrowid

                    for index, answer_text in enumerate(question["answers"]):
                        cur.execute(
                            "INSERT INTO answers (question_id, text, is_correct) VALUES (?, ?, ?)",
                            (question_id, answer_text, 1 if index == question["correct"] else 0),
                        )

    conn.commit()
'''


def write_runner(seed: dict[str, Any], target: Path) -> None:
    target = target.resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    seed_text = pformat(seed, width=120, sort_dicts=False)
    ast.parse(seed_text)
    target.write_text(
        RUNNER_TEMPLATE.replace("__SEED__", seed_text),
        encoding="utf-8",
    )
