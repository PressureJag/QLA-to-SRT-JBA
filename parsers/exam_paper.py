"""
Parses maths exam paper Word documents (.docx) to extract a question bank.

Q-number labels (Q1, Q2, …) must match the topic column headers in the QLA
spreadsheet for direct enrichment — no mapping step needed.

Paper 2 keys are prefixed with 'P2_' to avoid collision with Paper 1.
"""

import re
from pathlib import Path

try:
    from docx import Document
except ImportError:
    Document = None

_Q_START = re.compile(r"^Q(\d+)\b", re.IGNORECASE)
_MARKS   = re.compile(r"\((\d+)\s*marks?\)", re.IGNORECASE)


def _extract_questions(docx_path):
    """
    Walk every paragraph in a .docx and group text under Q-number headers.
    Returns {Q1: {text, marks}, Q2: {text, marks}, ...}
    """
    if Document is None:
        raise ImportError("python-docx is required: pip install python-docx")

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
