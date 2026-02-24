# utils/text_utils.py
import re
from typing import Optional

from config import LOWERCASE_WORDS


def sanitize(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name or "").strip("_")


def short_text(s: Optional[str], limit: int = 900) -> str:
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s.strip())
    return (s[: limit - 1] + "\u2026") if len(s) > limit else s


def norm_ref(x: Optional[str]) -> str:
    return re.sub(r"\s+", "", (x or "")).upper()


def norm_txt(s: str) -> str:
    if not s:
        return ""
    # normalize NBSP, dashes, whitespace
    s = s.replace("\u00A0", " ").replace("\xa0", " ")
    s = s.replace("–", "-").replace("—", "-")
    return " ".join(s.split())


def title_case(s: str) -> str:
    if not s:
        return ""
    return s.title()


def title_with_article(name: str) -> str:
    if not name:
        return ""
    return name if name.lower().startswith("the ") else f"The {name}"


def no_article(name: str) -> str:
    """Return the name without 'The' article."""
    if not name:
        return ""
    clean = name.strip()
    # Remove "The " if it exists at the beginning
    if clean.lower().startswith("the "):
        return clean[4:].strip()
    return clean


def title_with_acronyms(s: str, keep_all_caps=True) -> str:
    # simple title-caser with stop-words; preserves ALL-CAPS tokens and acronyms in ( )
    lowers = {"and", "or", "the", "of", "for", "to", "in", "on", "by", "with", "a", "an"}
    parts = []
    for word in s.split():
        base = word.strip()
        if keep_all_caps and base.isupper() and len(base) > 1:
            parts.append(base)
        else:
            w = base.lower()
            if w in lowers:
                parts.append(w)
            else:
                parts.append(w.capitalize())
    return " ".join(parts)


def allcaps(s: str) -> str:
    return (s or "").strip().upper()


def with_The_allcaps(name: str) -> str:
    # "The CITY OF ..." - ensure capital T + no double "The"
    raw = (name or "").strip()
    core = raw[4:].strip() if raw.lower().startswith("the ") else raw
    return f"The {allcaps(core)}"


def with_the_allcaps(name: str) -> str:
    # "the REHMANN ROBSON LLC" - ensure lowercase "the" + ALL CAPS entity
    raw = (name or "").strip()
    core = raw[4:].strip() if raw.lower().startswith("the ") else raw
    return f"the {allcaps(core)}"


def format_name_standard_case(name: str) -> str:
    """
    Format name in standard title case, removing 'The' article if present.
    Example: "CITY OF ANN ARBOR, MICHIGAN" -> "City Of Ann Arbor, Michigan"
    """
    if not name:
        return ""

    clean = name.strip()

    # Remove "The" or "the" if it exists at the beginning
    if clean.lower().startswith("the "):
        clean = clean[4:].strip()

    # If input is ALL CAPS (or mostly caps), normalize first
    letters = [ch for ch in clean if ch.isalpha()]
    if letters and sum(ch.isupper() for ch in letters) / len(letters) > 0.8:
        clean = clean.lower()

    titled = title_case(clean)

    # Lowercase connector words unless first word
    parts = titled.split(" ")
    for i, w in enumerate(parts):
        if i > 0 and w.lower() in LOWERCASE_WORDS:
            parts[i] = w.lower()

    return " ".join(parts)
