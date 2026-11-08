"""Wandelt das E-Mail-Handbook (Docx) in Markdown für den Agent um.

Das Docx bleibt das Original und wird nie verändert. Die Markdown-Datei wird
nur von diesem Skript erzeugt und nie von Hand bearbeitet.

    python scripts/handbook.py convert   # Docx -> Markdown, danach Vollständigkeitsprüfung
    python scripts/handbook.py check     # Exit 1, wenn die Markdown-Datei nicht zum Docx passt

Nur Python-Standardbibliothek, damit es ohne Zusatzpakete auf dem VPS läuft.
"""

import hashlib
import html
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCX = ROOT / "handbook" / "email-handbook.docx"
MD = ROOT / "handbook" / "email-handbook.md"

HEADER_RE = re.compile(r"^<!-- handbook-source: (\S+) sha256=([0-9a-f]{64}) version=(\S+) -->$")

# Farbe des linken Rahmens einer Hinweisbox (das Zeichen ⚠ ✓ ✗ ℹ steht zusätzlich im Text)
CALLOUT_COLORS = {"B8860B", "2E7D32", "1565C0", "C62828"}
HEADING_LEVEL = {"Heading1": 1, "Heading2": 2, "Heading3": 3}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_document(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def ensure_supported(xml: str) -> None:
    """Bricht ab, wenn das Docx Elemente enthält, die der Umwandler nicht kennt.

    So geht bei einer späteren Formatänderung nichts still verloren.
    """
    unsupported = {
        "w:drawing": "Bilder",
        "w:hyperlink": "Links",
        "w:footnoteReference": "Fussnoten",
        "w:commentReference": "Kommentare",
        "w:ins ": "Änderungsverfolgung",
        "w:del ": "Änderungsverfolgung",
        "w:sdt>": "Inhaltssteuerelemente",
        "w:gridSpan": "verbundene Tabellenzellen",
        "w:vMerge": "verbundene Tabellenzellen",
        "w:br": "manuelle Zeilenumbrüche",
        "w:tab/": "Tabulatoren",
        "w:strike": "Durchstreichung",
        "w:u ": "Unterstreichung",
        "w:vertAlign": "Hoch-/Tiefstellung",
    }
    found = sorted({label for tag, label in unsupported.items() if f"<{tag}" in xml})
    if found:
        sys.exit(f"Abbruch: Das Docx enthält {', '.join(found)}. Der Umwandler muss dafür erweitert werden.")

    levels = set(re.findall(r'<w:ilvl w:val="(\d+)"', xml))
    if levels - {"0"}:
        sys.exit("Abbruch: Das Docx enthält verschachtelte Aufzählungen. Der Umwandler muss dafür erweitert werden.")

    styles = set(re.findall(r'<w:pStyle w:val="([^"]+)"', xml)) - set(HEADING_LEVEL) - {"ListParagraph"}
    if styles:
        sys.exit(f"Abbruch: Unbekannte Absatzformate {sorted(styles)}. Der Umwandler muss dafür erweitert werden.")


def flag_on(rpr: str, tag: str) -> bool:
    match = re.search(rf"<w:{tag}(?: w:val=\"([^\"]*)\")?/>", rpr)
    return bool(match) and match.group(1) not in ("false", "0", "off")


def escape(text: str, in_table: bool) -> str:
    text = text.replace("\\", "\\\\").replace("*", "\\*").replace("_", "\\_").replace("~", "\\~")
    if in_table:
        text = text.replace("|", "\\|")
    return text


def runs_to_markdown(paragraph: str, in_table: bool) -> str:
    """Fett und Kursiv als ** und *; benachbarte Runs mit gleicher Formatierung werden zusammengefasst."""
    segments: list[tuple[bool, bool, str]] = []
    for run in re.findall(r"<w:r>(.*?)</w:r>", paragraph, re.S):
        rpr_match = re.search(r"<w:rPr>(.*?)</w:rPr>", run, re.S)
        rpr = rpr_match.group(1) if rpr_match else ""
        text = html.unescape("".join(re.findall(r"<w:t(?: [^>]*)?>(.*?)</w:t>", run, re.S)))
        if not text:
            continue
        key = (flag_on(rpr, "b"), flag_on(rpr, "i"))
        if segments and segments[-1][:2] == key:
            segments[-1] = (*key, segments[-1][2] + text)
        else:
            segments.append((*key, text))

    out = []
    for bold, italic, text in segments:
        core = text.strip()
        if not core:
            out.append(text)
            continue
        lead = text[: len(text) - len(text.lstrip())]
        trail = text[len(text.rstrip()) :]
        marker = ("**" if bold else "") + ("*" if italic else "")
        out.append(f"{lead}{marker}{escape(core, in_table)}{marker[::-1]}{trail}")
    return "".join(out).strip()


def paragraph_to_markdown(paragraph: str) -> str:
    ppr_match = re.search(r"<w:pPr>(.*?)</w:pPr>", paragraph, re.S)
    ppr = ppr_match.group(1) if ppr_match else ""
    text = runs_to_markdown(paragraph, in_table=False)
    if not text:
        return ""

    style = re.search(r'<w:pStyle w:val="([^"]+)"', ppr)
    style = style.group(1) if style else None

    if style in HEADING_LEVEL:
        # Überschriften ohne Fett/Kursiv-Marker, die Ebene trägt die Bedeutung
        plain = re.sub(r"(?<!\\)\*+", "", text).strip()
        return "#" * HEADING_LEVEL[style] + " " + plain
    if "<w:numPr>" in ppr:
        return "- " + text
    border = re.search(r'<w:left w:val="single" w:color="(\w+)"', ppr)
    if border and border.group(1) in CALLOUT_COLORS:
        return "> " + text
    if re.match(r"(\d+[.)]\s|[-+]\s|>|#)", text):
        text = "\\" + text  # Kein ungewolltes Listen- oder Überschriften-Markup
    return text


def table_to_markdown(table: str) -> str:
    rows = []
    for row in re.findall(r"<w:tr[ >].*?</w:tr>", table, re.S):
        cells = []
        for cell in re.findall(r"<w:tc>.*?</w:tc>", row, re.S):
            parts = [runs_to_markdown(p, in_table=True) for p in re.findall(r"<w:p[ >].*?</w:p>", cell, re.S)]
            cells.append("<br>".join(p for p in parts if p))
        rows.append(cells)

    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    header = [re.sub(r"(?<!\\)\*+", "", c) for c in rows[0]]  # Kopfzeile ist ohnehin fett
    lines = ["| " + " | ".join(header) + " |", "|" + "---|" * width]
    lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(lines)


def convert_xml(xml: str) -> str:
    body = re.search(r"<w:body>(.*)</w:body>", xml, re.S).group(1)
    blocks = []
    for match in re.finditer(r"<w:tbl>.*?</w:tbl>|<w:p[ >].*?</w:p>|<w:p/>", body, re.S):
        chunk = match.group(0)
        block = table_to_markdown(chunk) if chunk.startswith("<w:tbl>") else paragraph_to_markdown(chunk)
        if block:
            blocks.append(block)

    # Aufeinanderfolgende Listenpunkte ohne Leerzeile, damit sie eine Liste bilden
    out = []
    for block in blocks:
        if out and block.startswith("- ") and out[-1].startswith("- "):
            out[-1] += "\n" + block
        else:
            out.append(block)
    return "\n\n".join(out) + "\n"


# ---- Prüfung --------------------------------------------------------------


def words_from_docx(xml: str) -> list[str]:
    """Unabhängig vom Umwandler: aller Text in Dokumentreihenfolge."""
    body = re.search(r"<w:body>(.*)</w:body>", xml, re.S).group(1)
    paragraphs = re.findall(r"<w:p[ >].*?</w:p>", body, re.S)
    text = " ".join(
        html.unescape("".join(re.findall(r"<w:t(?: [^>]*)?>(.*?)</w:t>", p, re.S))) for p in paragraphs
    )
    return text.split()


def words_from_markdown(md: str) -> list[str]:
    words = []
    for line in md.splitlines():
        if HEADER_RE.match(line) or line.startswith("<!-- ") or re.fullmatch(r"\|(?:---\|)+", line):
            continue
        line = re.sub(r"^(#{1,6} |> |- )", "", line)
        if line.startswith("|") and line.endswith("|"):
            line = re.sub(r"(?<!\\)\|", " ", line).replace("<br>", " ")
        line = re.sub(r"(?<!\\)\*+", "", line)
        line = re.sub(r"\\(.)", r"\1", line)
        words += line.split()
    return words


def structure(xml: str) -> dict[str, int]:
    return {
        "headings": len(re.findall(r'<w:pStyle w:val="Heading\d"', xml)),
        "list_items": xml.count("<w:numPr>"),
        "tables": xml.count("<w:tbl>"),
        "table_rows": len(re.findall(r"<w:tr[ >]", xml)),
    }


def structure_md(md: str) -> dict[str, int]:
    lines = md.splitlines()
    return {
        "headings": sum(bool(re.match(r"#{1,6} ", l)) for l in lines),
        "list_items": sum(l.startswith("- ") for l in lines),
        "tables": sum(bool(re.fullmatch(r"\|(?:---\|)+", l)) for l in lines),
        "table_rows": sum(l.startswith("| ") for l in lines),
    }


def verify(xml: str, md: str) -> None:
    expected, actual = words_from_docx(xml), words_from_markdown(md)
    if expected != actual:
        for i, (a, b) in enumerate(zip(expected, actual)):
            if a != b:
                context = " ".join(expected[max(0, i - 8) : i + 8])
                sys.exit(f"Abbruch: Text weicht bei Wort {i} ab ({a!r} -> {b!r}). Stelle im Docx: …{context}…")
        sys.exit(f"Abbruch: Wortanzahl weicht ab (Docx {len(expected)}, Markdown {len(actual)}).")
    if structure(xml) != structure_md(md):
        sys.exit(f"Abbruch: Struktur weicht ab. Docx {structure(xml)}, Markdown {structure_md(md)}.")


def version_of(xml: str) -> str:
    match = re.search(r"Cold Email Handbook (v[\d.]+)", html.unescape(xml))
    return match.group(1) if match else "unbekannt"


def convert() -> None:
    xml = read_document(DOCX)
    ensure_supported(xml)
    md = convert_xml(xml)
    verify(xml, md)

    header = (
        f"<!-- handbook-source: {DOCX.name} sha256={sha256(DOCX)} version={version_of(xml)} -->\n"
        "<!-- Automatisch aus dem Docx erzeugt. Nicht von Hand bearbeiten: scripts/handbook.py convert -->\n\n"
    )
    full = header + md
    verify(xml, full)  # Auch die geschriebene Datei samt Kopfzeilen prüfen
    MD.write_text(full, encoding="utf-8", newline="\n")
    print(f"OK: {MD.relative_to(ROOT)} erzeugt ({version_of(xml)}, {len(words_from_docx(xml))} Wörter, {structure(xml)})")


def check() -> None:
    if not MD.exists():
        sys.exit("Handbook nicht aktuell: Markdown-Datei fehlt. Bitte scripts/handbook.py convert ausführen.")
    first = MD.read_text(encoding="utf-8").splitlines()[0]
    match = HEADER_RE.match(first)
    if not match or match.group(2) != sha256(DOCX):
        sys.exit("Handbook nicht aktuell: Das Docx wurde seit der letzten Umwandlung geändert.")
    print(f"OK: Handbook {match.group(3)} ist aktuell.")


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    command = sys.argv[1] if len(sys.argv) > 1 else ""
    if command == "convert":
        convert()
    elif command == "check":
        check()
    else:
        sys.exit(__doc__)
