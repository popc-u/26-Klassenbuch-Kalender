#!/usr/bin/env python3
"""Klassenbuch (Excel-Jahresraster) -> iCalendar-Dateien (.ics).

Liest das Klassenbuch der Klasse (Monatsblöcke mit Tag | Wochentag | Eintrag | KW)
und erzeugt zwei Kalenderdateien:

  klassenbuch-apple.ics      für Apple Kalender (Mac, iPhone, iPad)
  klassenbuch-microsoft.ics  für Microsoft Outlook (Outlook.com, Outlook im Web,
                             neues und klassisches Outlook)

Was im Kalender landet:
  * jeder Eintrag als ganztägiger Termin (mehrzeilige Zellen werden pro Zeile bzw.
    pro Fach aufgeteilt)
  * Einträge mit Uhrzeit ("17 Uhr", "12:30") als Termin zur angegebenen Zeit
  * Ferien (im Klassenbuch grün eingefärbt) als mehrtägige Termine
  * Sonderwochen (eingefärbt, am ersten Tag mit "Sonderwoche" beschriftet)
  * Randnotizen mit Datumsbereich (z. B. "Besuch aus Flöha 13.3.-19.3.")

Benötigt nur openpyxl:  pip install openpyxl
Aufruf:                 python klassenbuch_ics.py Klassenbuch.xlsx --out-dir ausgabe
"""
from __future__ import annotations

import argparse
import colorsys
import datetime as dt
import hashlib
import re
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

# --------------------------------------------------------------------------- #
# Konstanten
# --------------------------------------------------------------------------- #
MONTHS = {
    "januar": 1, "jan": 1, "jänner": 1, "februar": 2, "feb": 2, "märz": 3,
    "maerz": 3, "mär": 3, "mrz": 3, "april": 4, "apr": 4, "mai": 5, "juni": 6,
    "jun": 6, "juli": 7, "jul": 7, "august": 8, "aug": 8, "september": 9,
    "sep": 9, "sept": 9, "oktober": 10, "okt": 10, "november": 11, "nov": 11,
    "dezember": 12, "dez": 12,
}
WEEKDAYS_DE = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
HEADER_RE = re.compile(r"^\s*([A-Za-zÄÖÜäöü]+)\.?\s+(\d{4})\s*$")

# Fächer, an denen eine Zeile mit mehreren Einträgen aufgeteilt wird
# ("... (chapters 29-59) Geographie: Karte ..." -> zwei Termine).
SUBJECTS = [
    "Deutsch", "Englisch", "English", "Französisch", "Franz", "Francais",
    "Français", "Latein", "Italienisch", "Spanisch", "Mathematik", "Mathe",
    "Physik", "Chemie", "Biologie", "Biology", "Bio", "Geschichte",
    "Geographie", "Geografie", "Gg", "WR", "Wirtschaft", "Informatik", "Musik",
    "BG", "Sport", "Philosophie", "Religion", "FCE", "CAE",
]
SUBJECT_SPLIT_RE = re.compile(
    r"\s+(?=(?:%s)\s*:)" % "|".join(re.escape(s) for s in SUBJECTS)
)

# Feiertage (helfen, Ferien über freie Tage hinweg zusammenzufassen)
HOLIDAY_WORDS = [
    "neujahr", "berchtold", "karfreitag", "ostermontag", "tag der arbeit",
    "himmelfahrt", "auffahrt", "pfingstmontag", "weihnachtstag", "stephanstag",
    "fasnachtsmontag", "sechseläuten", "knabenschiessen", "bundesfeier",
]
HOLIDAY_FILLS = {"rgb:FFFFD9D9"}          # Feiertagsfarbe im Klassenbuch
TODAY_FILLS = {"rgb:FFFFFF00"}            # gelbe "Heute"-Markierung
NEUTRAL_FILLS = {None, "rgb:FFFFFFFF", "rgb:00000000", "theme:0:0.0"}

TIME_RANGE_RE = re.compile(
    r"(?<![\d.:])(\d{1,2})(?:[:.](\d{2}))?\s*(?:-|–|bis)\s*(\d{1,2})(?:[:.](\d{2}))?\s*Uhr",
    re.I,
)
TIME_COLON_RANGE_RE = re.compile(
    r"(?<![\d.:])(\d{1,2}):(\d{2})\s*(?:-|–|bis)\s*(\d{1,2}):(\d{2})(?![\d])"
)
TIME_COLON_RE = re.compile(r"(?<![\d.:])(\d{1,2}):(\d{2})(?![\d])")
TIME_UHR_RE = re.compile(r"(?<![\d.:])(\d{1,2})(?:\.(\d{2}))?\s*Uhr\b", re.I)
DEFAULT_DURATION = dt.timedelta(minutes=60)

DATE_RANGE_RE = re.compile(
    r"(\d{1,2})\.\s*(\d{1,2})\.?\s*(\d{2,4})?\s*(?:-|–|bis)\s*"
    r"(\d{1,2})\.\s*(\d{1,2})\.?\s*(\d{2,4})?"
)
SINGLE_DATE_RE = re.compile(r"(?<![\d.])(\d{1,2})\.\s*(\d{1,2})\.(?:\s*(\d{2,4}))?")


# --------------------------------------------------------------------------- #
# Datenstrukturen
# --------------------------------------------------------------------------- #
@dataclass
class Day:
    date: dt.date
    text: str = ""
    text_fill: str | None = None
    day_fill: str | None = None
    block_rows: tuple[int, int] = (0, 0)


@dataclass
class Event:
    title: str
    start: dt.date | dt.datetime
    end: dt.date | dt.datetime          # exklusiv (ganztägig: Folgetag)
    description: str = ""
    kind: str = "eintrag"               # eintrag | termin | ferien | sonderwoche | notiz
    uid_seed: str = ""

    @property
    def all_day(self) -> bool:
        return not isinstance(self.start, dt.datetime)


@dataclass
class Klassenbuch:
    name: str
    title: str
    days: dict[dt.date, Day] = field(default_factory=dict)
    notes: list[tuple[str, int, int]] = field(default_factory=list)  # (Text, Zeile, Spalte)
    month_blocks: list[tuple[int, int, int, int]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Farben
# --------------------------------------------------------------------------- #
def _theme_colors(wb) -> list[str]:
    """RGB-Werte der Designfarben in Excel-Reihenfolge (lt1, dk1, lt2, dk2, accent1..6)."""
    default = ["FFFFFF", "000000", "E8E8E8", "0E2841", "156082", "E97132",
               "196B24", "0F9ED5", "A02B93", "4EA72E"]
    raw = getattr(wb, "loaded_theme", None)
    if not raw:
        return default
    try:
        root = ET.fromstring(raw)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:clrScheme", ns)
        found = {}
        for child in scheme:
            tag = child.tag.split("}")[1]
            col = child[0]
            found[tag] = (col.get("lastClr") or col.get("val") or "").upper()
        order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6"]
        return [found.get(k) or d for k, d in zip(order, default)]
    except Exception:  # pragma: no cover - robust gegen exotische Dateien
        return default


def _apply_tint(rgb: str, tint: float) -> str:
    r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


class FillReader:
    def __init__(self, wb):
        self.theme = _theme_colors(wb)

    def key(self, cell) -> str | None:
        f = cell.fill
        if f is None or f.fill_type != "solid" or f.fgColor is None:
            return None
        c = f.fgColor
        if c.type == "rgb" and isinstance(c.rgb, str):
            return "rgb:" + c.rgb.upper()
        if c.type == "theme" and c.theme is not None:
            return "theme:%d:%s" % (c.theme, round(c.tint or 0.0, 2))
        if c.type == "indexed" and c.indexed is not None:
            return "idx:%d" % c.indexed
        return None

    def rgb(self, key: str | None) -> str | None:
        if not key:
            return None
        kind, _, rest = key.partition(":")
        if kind == "rgb":
            return rest[-6:]
        if kind == "theme":
            idx, _, tint = rest.partition(":")
            idx = int(idx)
            base = self.theme[idx] if idx < len(self.theme) else None
            return _apply_tint(base, float(tint)) if base else None
        if kind == "idx":
            from openpyxl.styles.colors import COLOR_INDEX
            i = int(rest)
            return COLOR_INDEX[i][-6:] if i < len(COLOR_INDEX) else None
        return None

    def is_green(self, key: str | None) -> bool:
        """Kräftiges Grün = Ferienfarbe im Klassenbuch."""
        rgb = self.rgb(key)
        if not rgb:
            return False
        r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
        h, s, v = colorsys.rgb_to_hsv(r, g, b)
        return 75 <= h * 360 <= 165 and s >= 0.35 and v >= 0.3


# --------------------------------------------------------------------------- #
# Einlesen
# --------------------------------------------------------------------------- #
def _int_value(v) -> int | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)) and float(v).is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return None


def _text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).replace("\r\n", "\n").replace("\r", "\n").strip()


def _find_sheet(wb, wanted: str | None):
    if wanted:
        return wb[wanted]
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 80)):
            for c in row:
                if isinstance(c.value, str):
                    m = HEADER_RE.match(c.value)
                    if m and m.group(1).lower() in MONTHS:
                        return ws
    return wb.worksheets[0]


def read_klassenbuch(path: str | Path, sheet: str | None = None) -> Klassenbuch:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _find_sheet(wb, sheet)
    fills = FillReader(wb)

    # Titel ("Klassenbuch 2eW Schuljahr 26/27") -> Kalendername
    title = ""
    for row in ws.iter_rows(min_row=1, max_row=3):
        for c in row:
            if isinstance(c.value, str) and "klassenbuch" in c.value.lower():
                title = " ".join(c.value.split())
                break
        if title:
            break
    m = re.search(r"Klassenbuch\s+(\S+)", title, re.I)
    name = f"Klassenbuch {m.group(1)}" if m else (title or "Klassenbuch")
    kb = Klassenbuch(name=name, title=title or name)

    # Monatsüberschriften finden ("Juli 2026", "August 2026", ...)
    headers = []
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                hm = HEADER_RE.match(c.value)
                if hm and hm.group(1).lower() in MONTHS:
                    headers.append((c.row, c.column, int(hm.group(2)), MONTHS[hm.group(1).lower()]))
    if not headers:
        raise SystemExit("Keine Monatsüberschriften (z. B. 'August 2026') gefunden.")

    block_cols: set[int] = set()
    header_rows_by_col: dict[int, list[int]] = {}
    for r, c, _, _ in headers:
        header_rows_by_col.setdefault(c, []).append(r)
        block_cols.update(range(c, c + 4))
    header_rows = sorted({r for r, *_ in headers})

    for r0, c0, year, month in headers:
        later = [r for r in header_rows_by_col[c0] if r > r0]
        r_end = (min(later) - 1) if later else ws.max_row
        kb.month_blocks.append((r0, r_end, year, month))
        for r in range(r0 + 1, r_end + 1):
            day_no = _int_value(ws.cell(r, c0).value)
            if day_no is None or not 1 <= day_no <= 31:
                continue
            try:
                date = dt.date(year, month, day_no)
            except ValueError:
                continue
            wd_cell = _text(ws.cell(r, c0 + 1).value)
            if wd_cell and wd_cell[:2] != WEEKDAYS_DE[date.weekday()]:
                kb.warnings.append(f"Wochentag passt nicht: {date:%d.%m.%Y} ist {WEEKDAYS_DE[date.weekday()]}, Zelle sagt '{wd_cell}'")
            text_cell = ws.cell(r, c0 + 2)
            parts = [_text(text_cell.value)]
            kw_val = ws.cell(r, c0 + 3).value
            if isinstance(kw_val, str) and kw_val.strip() and not kw_val.strip().isdigit():
                parts.append(_text(kw_val))       # Text versehentlich in KW-Spalte
            text = "\n".join(p for p in parts if p)
            day = Day(date=date, text=text, text_fill=fills.key(text_cell),
                      day_fill=fills.key(ws.cell(r, c0)), block_rows=(r0, r_end))
            if date in kb.days:                   # doppelter Tag: Texte zusammenführen
                old = kb.days[date]
                old.text = "\n".join(t for t in (old.text, day.text) if t)
            else:
                kb.days[date] = day

    # Randnotizen ausserhalb der Monatsblöcke (z. B. Spalte O), ohne Titelzeile
    first_header_row = header_rows[0]
    note_cells: dict[int, list[tuple[int, str]]] = {}
    for row in ws.iter_rows(min_row=first_header_row):
        for c in row:
            if c.column in block_cols or c.value is None:
                continue
            t = _text(c.value)
            if t and t.lower() != "heute":
                note_cells.setdefault(c.column, []).append((c.row, t))
    for col, cells in note_cells.items():
        cells.sort()
        group: list[tuple[int, str]] = []
        for r, t in cells + [(10**9, "")]:
            if group and r - group[-1][0] > 2:
                kb.notes.append((" ".join(x for _, x in group), group[0][0], col))
                group = []
            if t:
                group.append((r, t))

    kb._fills = fills  # type: ignore[attr-defined]
    return kb


# --------------------------------------------------------------------------- #
# Einträge -> Termine
# --------------------------------------------------------------------------- #
def split_entries(text: str) -> list[str]:
    out = []
    for line in text.split("\n"):
        line = " ".join(line.split())
        if not line:
            continue
        for part in SUBJECT_SPLIT_RE.split(line):
            part = part.strip(" ;,")
            if part:
                out.append(part)
    return out


def _valid_time(h, m) -> bool:
    return 0 <= h <= 23 and 0 <= m <= 59


def find_time(line: str) -> tuple[dt.time, dt.time | None] | None:
    for rx in (TIME_COLON_RANGE_RE, TIME_RANGE_RE):
        m = rx.search(line)
        if m:
            h1, m1, h2, m2 = (int(x) if x else 0 for x in m.groups())
            if _valid_time(h1, m1) and _valid_time(h2, m2) and (h2, m2) > (h1, m1):
                return dt.time(h1, m1), dt.time(h2, m2)
    for rx in (TIME_COLON_RE, TIME_UHR_RE):
        m = rx.search(line)
        if m:
            h1, m1 = int(m.group(1)), int(m.group(2) or 0)
            if _valid_time(h1, m1) and h1 >= 6:
                return dt.time(h1, m1), None
    return None


def _neutral(day: Day | None, fills: FillReader) -> bool:
    """Tage, über die hinweg eine Ferien-/Sonderwochen-Periode weiterläuft."""
    if day is None:
        return True
    if day.date.weekday() >= 5:
        return True
    if day.day_fill in HOLIDAY_FILLS or day.text_fill in HOLIDAY_FILLS:
        return True
    low = day.text.lower()
    return any(w in low for w in HOLIDAY_WORDS)


def _runs(core: list[dt.date], kb: Klassenbuch, fills: FillReader, max_gap: int = 5):
    runs: list[list[dt.date]] = []
    for d in core:
        if runs:
            prev = runs[-1][-1]
            gap = (d - prev).days - 1
            between = [prev + dt.timedelta(days=i) for i in range(1, gap + 1)]
            if gap <= max_gap and all(
                _neutral(kb.days.get(x), fills)
                or (kb.days.get(x) and kb.days[x].text_fill in TODAY_FILLS)
                for x in between
            ):
                runs[-1].append(d)
                continue
        runs.append([d])
    return runs


def ferien_name(start: dt.date) -> str:
    m = start.month
    if m in (6, 7, 8):
        return "Sommerferien"
    if m in (9, 10, 11):
        return "Herbstferien"
    if m in (12, 1):
        return "Weihnachtsferien"
    if m in (2, 3):
        return "Sportferien"
    return "Frühlingsferien"


def build_events(kb: Klassenbuch) -> list[Event]:
    fills: FillReader = kb._fills  # type: ignore[attr-defined]
    events: list[Event] = []
    days = kb.days

    # --- Ferien (grün eingefärbte Schultage) -------------------------------
    core = sorted(d for d, day in days.items() if d.weekday() < 5 and fills.is_green(day.text_fill))
    for run in _runs(core, kb, fills):
        start, end = run[0], run[-1]
        for _ in range(4):        # über angrenzendes Wochenende / Feiertage ausdehnen
            prev = start - dt.timedelta(days=1)
            if prev in days and _neutral(days[prev], fills):
                start = prev
            else:
                break
        for _ in range(4):
            nxt = end + dt.timedelta(days=1)
            if nxt in days and _neutral(days[nxt], fills):
                end = nxt
            else:
                break
        name = ferien_name(run[0])
        events.append(Event(
            title=name, start=start, end=end + dt.timedelta(days=1), kind="ferien",
            description=f"{name} laut {kb.name} ({start:%d.%m.%Y} – {end:%d.%m.%Y}).",
            uid_seed=f"ferien|{start}|{end}",
        ))

    # --- Sonderwochen (Farbe des mit "Sonderwoche" beschrifteten Tages) -----
    special_keys = {
        day.text_fill for day in days.values()
        if "sonderwoche" in day.text.lower() and day.text_fill not in NEUTRAL_FILLS
        and day.text_fill not in TODAY_FILLS and not fills.is_green(day.text_fill)
    }
    special_runs = []
    if special_keys:
        core = sorted(d for d, day in days.items() if d.weekday() < 5 and day.text_fill in special_keys)
        special_runs = _runs(core, kb, fills, max_gap=3)
    covered_titles: dict[dt.date, set[str]] = {}
    for run in special_runs:
        labels = [ln for d in run for ln in split_entries(days[d].text) if "sonderwoche" in ln.lower()]
        title = labels[0] if labels else "Sonderwoche"
        if not labels:
            continue
        start, end = run[0], run[-1]
        events.append(Event(
            title=title, start=start, end=end + dt.timedelta(days=1), kind="sonderwoche",
            description=f"{title} laut {kb.name} ({start:%d.%m.%Y} – {end:%d.%m.%Y}).",
            uid_seed=f"sonderwoche|{start}|{end}",
        ))
        for d in run:
            covered_titles.setdefault(d, set()).add(title.lower())

    # --- Tageseinträge ----------------------------------------------------
    for d in sorted(days):
        day = days[d]
        entries = split_entries(day.text)
        if not entries:
            continue
        seen: dict[str, int] = {}
        for line in entries:
            if line.lower() in covered_titles.get(d, set()):
                continue
            n = seen.get(line, 0)
            seen[line] = n + 1
            desc = f"{kb.name} · {WEEKDAYS_DE[d.weekday()]}, {d:%d.%m.%Y}"
            if len(entries) > 1:
                desc += "\n\nGanzer Eintrag an diesem Tag:\n" + "\n".join(entries)
            t = find_time(line)
            if t:
                start = dt.datetime.combine(d, t[0])
                end = dt.datetime.combine(d, t[1]) if t[1] else start + DEFAULT_DURATION
                if not t[1]:
                    desc += "\n\n(Endzeit steht nicht im Klassenbuch – 1 Stunde eingetragen.)"
                events.append(Event(title=line, start=start, end=end, description=desc,
                                    kind="termin", uid_seed=f"{d}|{line}|{n}"))
            else:
                events.append(Event(title=line, start=d, end=d + dt.timedelta(days=1),
                                    description=desc, kind="eintrag", uid_seed=f"{d}|{line}|{n}"))

    # --- Randnotizen mit Datum ---------------------------------------------
    all_dates = sorted(days)
    if all_dates:
        lo, hi = all_dates[0], all_dates[-1]
        for text, row, _col in kb.notes:
            block = next(((y, m) for r0, r1, y, m in kb.month_blocks if r0 <= row <= r1), None)

            def pick(dd: int, mm: int, yy: str | None) -> dt.date | None:
                if yy:
                    y = int(yy) + (2000 if len(yy) == 2 else 0)
                    try:
                        return dt.date(y, mm, dd)
                    except ValueError:
                        return None
                cands = []
                for y in range(lo.year, hi.year + 1):
                    try:
                        c = dt.date(y, mm, dd)
                    except ValueError:
                        continue
                    if lo <= c <= hi:
                        cands.append(c)
                if not cands:
                    return None
                if block:
                    same_row_band = [r0 for r0, r1, *_ in kb.month_blocks if r0 <= row <= r1]
                    band = [(y, m) for r0, r1, y, m in kb.month_blocks if r0 in same_row_band]
                    for c in cands:
                        if (c.year, c.month) in band:
                            return c
                return cands[0]

            m = DATE_RANGE_RE.search(text)
            if m:
                d1 = pick(int(m.group(1)), int(m.group(2)), m.group(3))
                d2 = pick(int(m.group(4)), int(m.group(5)), m.group(6))
                rest = (text[:m.start()] + " " + text[m.end():]).strip(" -–:,")
            else:
                m = SINGLE_DATE_RE.search(text)
                if not m:
                    kb.warnings.append(f"Notiz ohne Datum übersprungen (Zeile {row})")
                    continue
                d1 = d2 = pick(int(m.group(1)), int(m.group(2)), m.group(3))
                rest = (text[:m.start()] + " " + text[m.end():]).strip(" -–:,")
            if not d1 or not d2 or d2 < d1:
                kb.warnings.append(f"Notiz mit unklarem Datum übersprungen (Zeile {row})")
                continue
            title = " ".join(rest.split()) or "Notiz"
            events.append(Event(
                title=title, start=d1, end=d2 + dt.timedelta(days=1), kind="notiz",
                description=f"Randnotiz im {kb.name}: {' '.join(text.split())}",
                uid_seed=f"notiz|{d1}|{d2}|{title}",
            ))

    events.sort(key=lambda e: (e.start if isinstance(e.start, dt.datetime)
                               else dt.datetime.combine(e.start, dt.time.min), e.kind != "ferien"))
    return events


# --------------------------------------------------------------------------- #
# iCalendar schreiben
# --------------------------------------------------------------------------- #
VTIMEZONE_ZURICH = [
    "BEGIN:VTIMEZONE",
    "TZID:Europe/Zurich",
    "X-LIC-LOCATION:Europe/Zurich",
    "BEGIN:DAYLIGHT",
    "TZOFFSETFROM:+0100",
    "TZOFFSETTO:+0200",
    "TZNAME:CEST",
    "DTSTART:19700329T020000",
    "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
    "END:DAYLIGHT",
    "BEGIN:STANDARD",
    "TZOFFSETFROM:+0200",
    "TZOFFSETTO:+0100",
    "TZNAME:CET",
    "DTSTART:19701025T030000",
    "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
    "END:STANDARD",
    "END:VTIMEZONE",
]


def esc(text: str) -> str:
    return (text.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r\n", "\n").replace("\n", "\\n"))


def fold(line: str) -> str:
    """Zeilen nach RFC 5545 auf 75 Bytes umbrechen, ohne UTF-8-Zeichen zu zerteilen."""
    out, cur, limit = [], b"", 75
    for ch in line:
        b = ch.encode("utf-8")
        if len(cur) + len(b) > limit:
            out.append(cur)
            cur, limit = b" " + b, 75
        else:
            cur += b
    out.append(cur)
    return "\r\n".join(x.decode("utf-8") for x in out)


def to_ics(kb: Klassenbuch, events: list[Event], flavor: str, stamp: dt.datetime) -> str:
    assert flavor in ("apple", "microsoft")
    slug = re.sub(r"[^a-z0-9]+", "-", kb.name.lower()).strip("-") or "klassenbuch"
    dtstamp = stamp.astimezone(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//Klassenbuch-Kalender//{kb.name}//DE",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{esc(kb.name)}",
        f"X-WR-CALDESC:{esc('Hausaufgaben, Prüfungen, Termine und Ferien aus dem ' + kb.title + '. Wird stündlich aus dem Klassenbuch aktualisiert.')}",
        "X-WR-TIMEZONE:Europe/Zurich",
        "REFRESH-INTERVAL;VALUE=DURATION:PT1H",
        "X-PUBLISHED-TTL:PT1H",
    ]
    lines += VTIMEZONE_ZURICH
    for e in events:
        uid = hashlib.sha1(e.uid_seed.encode("utf-8")).hexdigest()[:20] + f"@{slug}"
        lines += ["BEGIN:VEVENT", f"UID:{uid}", f"DTSTAMP:{dtstamp}"]
        if e.all_day:
            lines += [f"DTSTART;VALUE=DATE:{e.start:%Y%m%d}", f"DTEND;VALUE=DATE:{e.end:%Y%m%d}"]
        else:
            lines += [f"DTSTART;TZID=Europe/Zurich:{e.start:%Y%m%dT%H%M%S}",
                      f"DTEND;TZID=Europe/Zurich:{e.end:%Y%m%dT%H%M%S}"]
        lines.append(f"SUMMARY:{esc(e.title)}")
        if e.description:
            lines.append(f"DESCRIPTION:{esc(e.description)}")
        cat = {"ferien": "Ferien", "sonderwoche": "Sonderwoche", "notiz": "Termin",
               "termin": "Termin", "eintrag": "Klassenbuch"}[e.kind]
        lines += [f"CATEGORIES:{cat}", "TRANSP:TRANSPARENT", "SEQUENCE:0", "STATUS:CONFIRMED"]
        if flavor == "microsoft":
            lines += [
                "X-MICROSOFT-CDO-BUSYSTATUS:FREE",
                "X-MICROSOFT-CDO-INTENDEDSTATUS:FREE",
                f"X-MICROSOFT-CDO-ALLDAYEVENT:{'TRUE' if e.all_day else 'FALSE'}",
                "X-MICROSOFT-CDO-IMPORTANCE:1",
            ]
        lines.append("END:VEVENT")
    lines.append("END:VCALENDAR")
    return "\r\n".join(fold(l) for l in lines) + "\r\n"


# --------------------------------------------------------------------------- #
# Kommandozeile
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", help="Klassenbuch als .xlsx")
    ap.add_argument("--out-dir", default=".", help="Zielordner für die .ics-Dateien")
    ap.add_argument("--prefix", default="klassenbuch", help="Dateiname-Anfang (Standard: klassenbuch)")
    ap.add_argument("--sheet", help="Tabellenblatt (Standard: das erste mit Monatsüberschriften)")
    ap.add_argument("--name", help="Kalendername (Standard: aus Titelzeile, z. B. 'Klassenbuch 2eW')")
    ap.add_argument("--list", action="store_true", help="alle Termine anzeigen (nicht in öffentlichen Logs verwenden)")
    args = ap.parse_args(argv)

    kb = read_klassenbuch(args.xlsx, args.sheet)
    if args.name:
        kb.name = args.name
    events = build_events(kb)
    if not events:
        print("Keine Termine gefunden – Datei wird nicht überschrieben.", file=sys.stderr)
        return 2

    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now(dt.timezone.utc).replace(microsecond=0)
    for flavor in ("apple", "microsoft"):
        (out / f"{args.prefix}-{flavor}.ics").write_text(
            to_ics(kb, events, flavor, stamp), encoding="utf-8", newline="")

    kinds: dict[str, int] = {}
    for e in events:
        kinds[e.kind] = kinds.get(e.kind, 0) + 1
    print(f"{kb.name}: {len(events)} Termine " + ", ".join(f"{k}={v}" for k, v in sorted(kinds.items())))
    for w in kb.warnings:
        print("Hinweis:", w, file=sys.stderr)
    if args.list:
        for e in events:
            wd = WEEKDAYS_DE[e.start.weekday()]
            if e.all_day:
                last = e.end - dt.timedelta(days=1)
                when = f"{wd} {e.start:%d.%m.%Y}" + (
                    f" – {WEEKDAYS_DE[last.weekday()]} {last:%d.%m.%Y}" if last != e.start else "")
            else:
                when = f"{wd} {e.start:%d.%m.%Y %H:%M}–{e.end:%H:%M}"
            print(f"  [{e.kind:11}] {when:34} {e.title}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
