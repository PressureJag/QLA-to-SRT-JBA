"""
Parses maths exam paper Word documents (.doc / .docx) to extract a question bank.

Q-number labels (Q1, Q2, …) must match the topic column headers in the QLA
spreadsheet for direct enrichment — no mapping step needed.

Paper 2 keys are prefixed with 'P2_' to avoid collision with Paper 1.
"""

import os
import re
import subprocess
import tempfile
from pathlib import Path

try:
    from docx import Document
except ImportError:
    Document = None

_Q_START = re.compile(r"^Q(\d+)\b", re.IGNORECASE)
_MARKS   = re.compile(r"\((\d+)\s*marks?\)", re.IGNORECASE)


def _to_docx(path):
    """
    Return a path to a .docx version of the file.
    If the file is already .docx, returns (original_path, False).
    If .doc, converts via macOS textutil and returns (tmp_path, True).
    Caller is responsible for deleting the tmp file when is_temp=True.
    """
    if Path(str(path)).suffix.lower() == ".docx":
        return str(path), False
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp.close()
    try:
        subprocess.run(
            ["textutil", "-convert", "docx", "-output", tmp.name, str(path)],
            check=True, capture_output=True, timeout=30,
        )
        return tmp.name, True
    except Exception:
        os.unlink(tmp.name)
        raise


def _extract_questions(source_path):
    """
    Walk every paragraph in a Word document and group text under Q-number headers.
    Accepts both .docx and .doc (the latter is converted via macOS textutil).
    Returns {Q1: {text, marks}, Q2: {text, marks}, ...}
    """
    if Document is None:
        raise ImportError("python-docx is required: pip install python-docx")

    docx_path, is_temp = _to_docx(source_path)
    try:
        doc = Document(str(docx_path))
        bank = {}
        current_key = None
        current_parts = []

        def _flush():
            if not current_key or not current_parts:
                return
            full = " ".join(p.strip() for p in current_parts if p.strip())
            m = _MARKS.search(full)
            bank[current_key] = {
                "text":  full,
                "marks": int(m.group(1)) if m else None,
            }

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            m = _Q_START.match(text)
            if m:
                _flush()
                current_key = f"Q{m.group(1)}"
                current_parts = [text]
            elif current_key:
                current_parts.append(text)

        _flush()

        # Also scan table cells — some papers embed questions in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text = para.text.strip()
                        if not text:
                            continue
                        m = _Q_START.match(text)
                        if m:
                            key = f"Q{m.group(1)}"
                            if key not in bank:
                                bank[key] = {"text": text, "marks": None}
                                m2 = _MARKS.search(text)
                                if m2:
                                    bank[key]["marks"] = int(m2.group(1))

        return bank
    finally:
        if is_temp and os.path.exists(docx_path):
            os.unlink(docx_path)


def parse_papers(paper1_path=None, paper2_path=None):
    """
    Parse one or both exam paper Word documents.
    Returns a combined question bank dict.
    Paper 2 keys are prefixed with 'P2_' to avoid collision with Paper 1.
    Returns {} if no valid paths supplied.
    """
    bank = {}

    if paper1_path and Path(str(paper1_path)).exists():
        try:
            bank.update(_extract_questions(paper1_path))
        except Exception:
            pass

    if paper2_path and Path(str(paper2_path)).exists():
        try:
            for key, val in _extract_questions(paper2_path).items():
                bank[f"P2_{key}"] = val
        except Exception:
            pass

    return bank
