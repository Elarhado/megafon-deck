"""Shared helpers for build_deck / lint_deck / render: dependency check, assets, typesetting, fonts."""
import glob
import json
import os
import re
import sys

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
PLUGIN_DIR = os.path.dirname(SCRIPTS_DIR)
ASSETS = os.path.join(PLUGIN_DIR, "assets")
INSTALL_CMD = "python3 -m pip install --user python-pptx pillow"


def require(*modules):
    """Import modules by name; on failure print install hint and exit 3."""
    import importlib
    for name in modules:
        try:
            importlib.import_module(name)
        except ImportError:
            pkg = {"pptx": "python-pptx", "PIL": "pillow"}.get(name, name)
            print("Missing Python package '%s' (import %s). Install with:\n  %s" % (pkg, name, INSTALL_CMD),
                  file=sys.stderr)
            sys.exit(3)


def load_json(name):
    with open(os.path.join(ASSETS, name), encoding="utf-8") as f:
        return json.load(f)


def palette():
    """Return (all allowed hex set, gradient-only hex set, name->hex map)."""
    b = load_json("brand.json")
    named = {}
    named.update(b["colors"]["main"])
    named.update(b["colors"]["additional"])
    named = {k: v.upper() for k, v in named.items()}
    grad = {named[n] for n in b["colors"]["gradient_only"]}
    return set(named.values()), grad, named


# ---------------------------------------------------------------- typesetting
NBSP = "\u00a0"
NNBSP = "\u202f"
_URL = re.compile(r"(https?://\S+|www\.\S+|\b[\w.+-]+@[\w-]+\.[\w.-]+)")
_BRAND = re.compile(r"\b(мегафон|megafon)(\w*)", re.I)


def _brand_fix(m):
    word, end = m.group(1), m.group(2)
    stem = "МегаФон" if re.match(r"[а-яё]", word, re.I) else "MegaFon"
    if m.group(0).isupper() and end:
        end = end.lower()
    return stem + end


def _quotes(s):
    out, depth = [], 0
    for i, ch in enumerate(s):
        if ch != '"':
            out.append(ch)
            continue
        prev = s[i - 1] if i else " "
        opening = prev.isspace() or prev in "([{«„—–-/" or i == 0
        if opening:
            out.append("«" if depth == 0 else "„")
            depth += 1
        else:
            depth = max(depth - 1, 0)
            out.append("»" if depth == 0 else "“")
    return "".join(out)


def _group(num):
    return f"{int(num):,}".replace(",", NNBSP)


def _typeset_plain(s):
    s = _quotes(s)
    s = re.sub(r"[ \t]+-[ \t]+", NBSP + "— ", s)
    s = re.sub(r"[ \t]+—[ \t]+", NBSP + "— ", s)
    # ranges 5-6 -> 5–6, but not 3-group strings (dates, phones: 2026-10-14, 8-800-555)
    s = re.sub(r"(?<![\d-])(\d+)-(\d+)(?![\d-])", r"\1–\2", s)
    # numbers already grouped with plain/nbsp spaces (12 500, 10 000 000) -> narrow no-break space
    s = re.sub(r"(?<![\d.,])(\d{1,3}(?:[ \u00a0]\d{3})+)(?![\d]|[.,]\d)",
               lambda m: re.sub(r"[ \u00a0]", NNBSP, m.group(1)), s)
    s = re.sub(r"(?<![\w.,])(\d{5,})(?![\d]|[.,]\d)", lambda m: _group(m.group(1)), s)
    s = re.sub(r"(?<![\w.,])(\d{4})(?=\s?₽|\s+руб)", lambda m: _group(m.group(1)), s)
    s = re.sub(r"\s*₽", NNBSP + "₽", s)
    s = _BRAND.sub(_brand_fix, s)
    s = re.sub(r"(?<=\S) {2,}", " ", s)
    return s


def typeset(text):
    """Apply MegaFon brandbook typesetting rules to a string (URLs untouched)."""
    if not text:
        return text
    parts = _URL.split(text)
    # quotes must be tracked across URL boundaries: process plain parts together via placeholders
    marks, buf = [], []
    for i, p in enumerate(parts):
        if i % 2:
            marks.append(p)
            buf.append("\ue000")
        else:
            buf.append(p)
    joined = _typeset_plain("".join(buf))
    it = iter(marks)
    return re.sub("\ue000", lambda m: next(it), joined)


def typeset_issues(text):
    """Return list of human-readable typesetting violations found in text."""
    if not text:
        return []
    plain = _URL.sub("\ue000", text)
    out = []
    if '"' in plain:
        out.append('straight double quotes (use «»)')
    if re.search(r"\s-\s", plain):
        out.append('hyphen with spaces " - " (use non-breaking space + em dash)')
    if re.search(r"(?<![\d-])\d+-\d+(?![\d-])", plain):
        out.append("digit range with hyphen (use en dash 5–6)")
    for m in _BRAND.finditer(plain):
        if _brand_fix(m) != m.group(0):
            out.append("brand spelling '%s' (use %s)" % (m.group(0), _brand_fix(m)))
    if re.search(r"(?<=\S) {2,}", plain):
        out.append("double spaces")
    if re.search(r"(?<![\w.,])\d{5,}(?![\d]|[.,]\d)", plain):
        out.append("number without narrow-space thousands grouping")
    if re.search(r"[ \u00a0]₽", plain):
        out.append("regular space before ₽ (use U+202F)")
    return out


# ---------------------------------------------------------------------- fonts
def find_font(bold=False):
    """Path to an Arial-metric font (Arial or Liberation Sans) or None."""
    home = os.path.expanduser("~")
    reg = ["/System/Library/Fonts/Supplemental/Arial.ttf", "/Library/Fonts/Arial.ttf",
           home + "/Library/Fonts/Arial.ttf", "C:/Windows/Fonts/arial.ttf",
           "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
           "/usr/share/fonts/liberation/LiberationSans-Regular.ttf",
           "/usr/share/fonts/TTF/LiberationSans-Regular.ttf"]
    bld = ["/System/Library/Fonts/Supplemental/Arial Bold.ttf", "/Library/Fonts/Arial Bold.ttf",
           home + "/Library/Fonts/Arial Bold.ttf", "C:/Windows/Fonts/arialbd.ttf",
           "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
           "/usr/share/fonts/liberation/LiberationSans-Bold.ttf",
           "/usr/share/fonts/TTF/LiberationSans-Bold.ttf"]
    for p in (bld if bold else reg):
        if os.path.exists(p):
            return p
    pat = "LiberationSans-Bold.ttf" if bold else "LiberationSans-Regular.ttf"
    for root in ("/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts")):
        hit = glob.glob(os.path.join(root, "**", pat), recursive=True)
        if hit:
            return hit[0]
    return None
