"""بحث بسيط في نصوص المعايير اللي إنت حاططها في فولدر standards/.

مهم: نصوص معايير المراجعة الدولية (ISA) والمعايير المصرية محمية بحقوق نشر،
فحط نسخك المرخّصة بس. دي نسخة بسيطة بتعتمد على الكلمات المفتاحية؛
لما الملفات تكبر استبدلها بـ embeddings + vector DB (مثلاً ChromaDB).
"""
from __future__ import annotations

import os
import re
from pathlib import Path

_TOKEN = re.compile(r"[\w؀-ۿ]+", re.UNICODE)


def _chunks(text: str, max_chars: int = 900) -> list[str]:
    out, cur = [], ""
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        if len(cur) + len(para) > max_chars and cur:
            out.append(cur)
            cur = para
        else:
            cur = f"{cur}\n\n{para}" if cur else para
    if cur:
        out.append(cur)
    return out


def search_standards(query: str, k: int = 4, standards_dir: str | None = None) -> list[dict]:
    base = Path(standards_dir or os.environ.get("AUDIT_STANDARDS_DIR", "standards"))
    q = {t.lower() for t in _TOKEN.findall(query) if len(t) > 2}
    if not q or not base.exists():
        return []

    scored = []
    for f in sorted(base.rglob("*")):
        if f.suffix.lower() not in (".md", ".txt"):
            continue
        text = f.read_text(encoding="utf-8", errors="ignore")
        for n, ch in enumerate(_chunks(text), start=1):
            tokens = [t.lower() for t in _TOKEN.findall(ch)]
            score = sum(tokens.count(w) for w in q)
            if score:
                scored.append((score, f.name, n, ch))

    scored.sort(key=lambda x: -x[0])
    return [
        {"source_file": name, "chunk": n, "score": s, "text": ch}
        for s, name, n, ch in scored[:k]
    ]
