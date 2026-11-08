"""توحيد النص العربي للمقارنة (بدون تشكيل، وهمزات وتاء مربوطة موحّدة)."""
import re

_DIAC = re.compile(r"[ؐ-ًؚ-ٰٟۖ-ۭـ]")
_MAP = str.maketrans(
    {"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه", "ؤ": "و", "ئ": "ي"}
)


def norm(s) -> str:
    s = "" if s is None else str(s)
    s = _DIAC.sub("", s.strip().lower()).translate(_MAP)
    return re.sub(r"\s+", " ", s)
