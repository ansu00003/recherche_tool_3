#!/usr/bin/env python3
"""
DOCX Generation Script for KALKU Tender Tool v3.
Reads a vorlage.docx template, fills in tender entries, handles variable
section counts (cut excess / extend with new), and outputs the result.

Usage:
    python3 generate_docx.py <vorlage_path> <output_path> < input.json

Input JSON (stdin):
    {
        "entries": [ { "titel", "dtad_id", "abgabetermin", "ausfuehrungsort",
                        "beginn", "ende", "leistung" }, ... ],
        "gewerk": "...",
        "region": "..."
    }
"""

import sys
import json
import re
import copy
import locale
from datetime import datetime
from lxml import etree
from docx import Document

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W}
DASH = "\u2014"  # em-dash default for missing fields

# German month names for date formatting
GERMAN_MONTHS = [
    "", "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]


def format_ausfuehrungsfrist(entry):
    """Format Ausführungsfrist based on available data.
    
    Case 1: Both dates exist → "Ausführungsfrist: Beginn: DD.MM.YYYY - Ende: DD.MM.YYYY"
    Case 2: No dates but duration → "Ausführungsfrist: Geschätzte Laufzeit - X Monate"
    Case 3: No dates, no duration → "Ausführungsfrist: Siehe Vergabeunterlagen"
    """
    beginn = entry.get("beginn", DASH)
    ende = entry.get("ende", DASH)
    duration = entry.get("duration", DASH)
    
    if beginn == DASH and ende == DASH:
        if duration != DASH:
            # Case 2: No dates but duration exists
            return f"Ausführungsfrist: Geschätzte Laufzeit - {duration}"
        else:
            # Case 3: No dates, no duration
            return "Ausführungsfrist: Siehe Vergabeunterlagen"
    else:
        # Case 1: At least one date exists
        return f"Ausführungsfrist: Beginn: {beginn} - Ende: {ende}"


# ============================================================
# XML HELPERS
# ============================================================

def clear_paragraph_text(p):
    """Remove all w:r runs from a paragraph, keeping w:pPr intact."""
    for r in p.findall(f"{{{W}}}r"):
        p.remove(r)
    # Also remove bookmarks, hyperlinks, etc. that contain text
    for tag in ["bookmarkStart", "bookmarkEnd", "hyperlink"]:
        for el in p.findall(f"{{{W}}}{tag}"):
            p.remove(el)


def make_run(text, font="Arial", sz=20, bold=False):
    """Create a new w:r element with specified formatting."""
    r = etree.SubElement(etree.Element("dummy"), f"{{{W}}}r")
    rpr = etree.SubElement(r, f"{{{W}}}rPr")
    fonts = etree.SubElement(rpr, f"{{{W}}}rFonts")
    fonts.set(f"{{{W}}}ascii", font)
    fonts.set(f"{{{W}}}hAnsi", font)
    fonts.set(f"{{{W}}}cs", font)
    sz_el = etree.SubElement(rpr, f"{{{W}}}sz")
    sz_el.set(f"{{{W}}}val", str(sz))
    sz_cs = etree.SubElement(rpr, f"{{{W}}}szCs")
    sz_cs.set(f"{{{W}}}val", str(sz))
    if bold:
        etree.SubElement(rpr, f"{{{W}}}b")
        etree.SubElement(rpr, f"{{{W}}}bCs")
    t = etree.SubElement(r, f"{{{W}}}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def make_run_with_rpr(source_run, text):
    """Create a new w:r cloning rPr from source_run, with new text."""
    r = etree.SubElement(etree.Element("dummy"), f"{{{W}}}r")
    source_rpr = source_run.find(f"{{{W}}}rPr")
    if source_rpr is not None:
        r.append(copy.deepcopy(source_rpr))
    t = etree.SubElement(r, f"{{{W}}}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def make_paragraph(ppr_source=None):
    """Create a new w:p element, optionally cloning pPr from source."""
    p = etree.Element(f"{{{W}}}p")
    if ppr_source is not None:
        p.append(copy.deepcopy(ppr_source))
    return p


def get_ppr(p):
    """Get or create w:pPr for a paragraph element."""
    ppr = p.find(f"{{{W}}}pPr")
    if ppr is None:
        ppr = etree.SubElement(p, f"{{{W}}}pPr")
        p.insert(0, ppr)
    return ppr


def set_keep_next(p):
    """Add w:keepNext to paragraph properties."""
    ppr = get_ppr(p)
    if ppr.find(f"{{{W}}}keepNext") is None:
        etree.SubElement(ppr, f"{{{W}}}keepNext")


def set_hanging_indent(p, left=144, hanging=144):
    """Set hanging indent on paragraph (for bullet points)."""
    ppr = get_ppr(p)
    ind = ppr.find(f"{{{W}}}ind")
    if ind is None:
        ind = etree.SubElement(ppr, f"{{{W}}}ind")
    ind.set(f"{{{W}}}left", str(left))
    ind.set(f"{{{W}}}hanging", str(hanging))


# ============================================================
# SECTION DETECTION
# ============================================================

def detect_sections(body):
    """
    Scan all <w:p> paragraphs looking for section headers like "1. ...", "2. ...".
    Returns list of {"num": int, "start_idx": int, "paras": [elements]}.
    Skips date-like patterns (e.g. "4. Juni 2025").
    """
    paragraphs = body.findall(f"{{{W}}}p")
    sections = []
    current_section = None

    # Pattern: starts with a number followed by "." and optional text
    section_re = re.compile(r"^(\d+)\.\s*(.*)")
    # Date pattern to skip: "4. Juni 2025" etc.
    date_skip_re = re.compile(
        r"^\d+\.\s*(?:Januar|Februar|März|April|Mai|Juni|Juli|August|"
        r"September|Oktober|November|Dezember)\s+\d{4}",
        re.IGNORECASE
    )

    for idx, p in enumerate(paragraphs):
        text = "".join(t.text or "" for t in p.findall(f".//{{{W}}}t")).strip()

        # Check if this paragraph starts a new section
        m = section_re.match(text)
        if m and not date_skip_re.match(text):
            num = int(m.group(1))
            # Save previous section
            if current_section is not None:
                sections.append(current_section)
            current_section = {"num": num, "start_idx": idx, "paras": [p]}
        elif current_section is not None:
            current_section["paras"].append(p)

    # Don't forget the last section
    if current_section is not None:
        sections.append(current_section)

    return sections


# ============================================================
# FILL A SECTION WITH ENTRY DATA
# ============================================================

def fill_section(section, entry, section_num):
    """Fill an existing template section with entry data.
    
    Adds keepNext to all content paragraphs (title through last bullet) to prevent
    page breaks within an entry. Empty spacers between entries do NOT get keepNext.
    """
    for p in section["paras"]:
        text = "".join(t.text or "" for t in p.findall(f".//{{{W}}}t")).strip()

        if text.startswith(f"{section['num']}."):
            # Title line — replace with new section number + titel
            titel = entry.get("titel", DASH)
            first_run = p.findall(f"{{{W}}}r")
            if first_run:
                ref_run = first_run[0]
                clear_paragraph_text(p)
                p.append(make_run_with_rpr(ref_run, f"{section_num}. {titel}"))
            set_keep_next(p)  # Keep title with following content

        elif text.startswith("ID:"):
            first_run = p.findall(f"{{{W}}}r")
            if first_run:
                ref_run = first_run[0]
                clear_paragraph_text(p)
                p.append(make_run_with_rpr(ref_run, f"ID: {entry.get('dtad_id', DASH)}"))
            set_keep_next(p)

        elif text.startswith("Abgabetermin:"):
            first_run = p.findall(f"{{{W}}}r")
            if first_run:
                ref_run = first_run[0]
                clear_paragraph_text(p)
                p.append(make_run_with_rpr(ref_run, f"Abgabetermin: {entry.get('abgabetermin', DASH)}"))
            set_keep_next(p)

        elif text.startswith("Ausführungsort:"):
            first_run = p.findall(f"{{{W}}}r")
            if first_run:
                ref_run = first_run[0]
                clear_paragraph_text(p)
                p.append(make_run_with_rpr(ref_run, f"Ausführungsort: {entry.get('ausfuehrungsort', DASH)}"))
            set_keep_next(p)

        elif text.startswith("Ausführungsfrist:"):
            first_run = p.findall(f"{{{W}}}r")
            if first_run:
                ref_run = first_run[0]
                clear_paragraph_text(p)
                p.append(make_run_with_rpr(ref_run, format_ausfuehrungsfrist(entry)))
            set_keep_next(p)

        elif text.startswith("Art und Umfang der Leistung:"):
            fill_leistung(p, entry, section)

        elif not text:
            # Empty spacer paragraphs within the entry — keep with next
            # so the entry is not split across pages
            set_keep_next(p)


def fill_leistung(heading_p, entry, section):
    """Fill the 'Art und Umfang der Leistung' field with bullet points.
    
    All paragraphs (heading + bullets except last) get keepNext to keep entry together.
    """
    leistung = entry.get("leistung", DASH)
    if not leistung or leistung == DASH:
        return

    # Get reference run from the heading
    runs = heading_p.findall(f"{{{W}}}r")
    ref_run = runs[0] if runs else None
    if ref_run is None:
        return

    # Add keepNext to heading so it stays with bullets
    set_keep_next(heading_p)

    # Split leistung into bullet points
    lines = split_leistung_to_bullets(leistung)
    if not lines:
        return

    # Find where to insert bullet paragraphs (after the heading)
    # Look for the next paragraph after the heading within this section
    heading_idx = section["paras"].index(heading_p)
    insert_after = heading_p

    # Remove empty paragraphs between heading and end of section (they'll be replaced by bullets)
    paras_to_remove = []
    for p in section["paras"][heading_idx + 1:]:
        text = "".join(t.text or "" for t in p.findall(f".//{{{W}}}t")).strip()
        if not text:
            paras_to_remove.append(p)

    for p in paras_to_remove:
        parent = p.getparent()
        if parent is not None:
            parent.remove(p)

    # Get paragraph properties from heading for cloning
    heading_ppr = heading_p.find(f"{{{W}}}pPr")

    # Insert each bullet as its own paragraph
    # All bullets except the last one get keepNext
    for idx, line in enumerate(lines):
        bullet_text = line if line.startswith("- ") else f"- {line}"
        new_p = make_paragraph(heading_ppr)

        # Clear any inherited keepNext first
        new_ppr = new_p.find(f"{{{W}}}pPr")
        if new_ppr is not None:
            kn = new_ppr.find(f"{{{W}}}keepNext")
            if kn is not None:
                new_ppr.remove(kn)

        # Add keepNext to all bullets EXCEPT the last one
        # (last bullet doesn't need to keep with anything - allows page break after entry)
        if idx < len(lines) - 1:
            set_keep_next(new_p)

        # Add hanging indent for bullets
        set_hanging_indent(new_p)

        # Add the text run
        new_p.append(make_run_with_rpr(ref_run, bullet_text))

        # Insert after the previous element
        insert_after.addnext(new_p)
        insert_after = new_p
    
    # ADD: Empty spacer paragraph AFTER the last bullet (gap before next entry)
    spacer_after_bullets = make_paragraph(heading_ppr)
    # Clear any inherited keepNext from spacer
    spacer_ppr = spacer_after_bullets.find(f"{{{W}}}pPr")
    if spacer_ppr is not None:
        kn = spacer_ppr.find(f"{{{W}}}keepNext")
        if kn is not None:
            spacer_ppr.remove(kn)
    insert_after.addnext(spacer_after_bullets)


def split_leistung_to_bullets(leistung):
    """Split leistung text into bullet point lines."""
    if not leistung or leistung == DASH:
        return []

    # Split on NEWLINES first (each line is a bullet)
    lines = [l.strip() for l in leistung.split("\n") if l.strip()]
    
    # Remove leading "- " from each line if present
    result = []
    for line in lines:
        if line.startswith("- "):
            result.append(line[2:].strip())
        else:
            result.append(line.strip())
    
    return [r for r in result if r]

    # Single block of text — split on sentences or semicolons
    parts = re.split(r"[;]\s*", leistung)
    if len(parts) > 1:
        return [p.strip() for p in parts if p.strip()]

    # Just return as single item
    return [leistung.strip()] if leistung.strip() else []


# ============================================================
# CUT EXCESS SECTIONS
# ============================================================

def cut_excess_sections(body, sections, num_entries):
    """Remove sections beyond num_entries (iterate in reverse to preserve indices)."""
    for sec in reversed(sections):
        if sec["num"] > num_entries:
            for p in sec["paras"]:
                parent = p.getparent()
                if parent is not None:
                    parent.remove(p)
    return [s for s in sections if s["num"] <= num_entries]


# ============================================================
# EXTEND WITH NEW SECTIONS
# ============================================================

def find_format_references(sections):
    """Find formatting references from section 1."""
    if not sections:
        return None, None, None

    sec1 = sections[0]
    bold_ref_run = None
    normal_ref_run = None
    normal_ppr_ref = None

    for p in sec1["paras"]:
        text = "".join(t.text or "" for t in p.findall(f".//{{{W}}}t")).strip()
        for r in p.findall(f"{{{W}}}r"):
            rpr = r.find(f"{{{W}}}rPr")
            if rpr is None:
                continue
            fonts = rpr.find(f"{{{W}}}rFonts")
            if fonts is None:
                continue
            font_name = fonts.get(f"{{{W}}}ascii", "")
            if "Arial" not in font_name:
                continue

            b_el = rpr.find(f"{{{W}}}b")
            if b_el is not None and bold_ref_run is None:
                bold_ref_run = r
            elif b_el is None and normal_ref_run is None:
                normal_ref_run = r
                # Also grab the paragraph properties
                ppr = p.find(f"{{{W}}}pPr")
                if ppr is not None:
                    normal_ppr_ref = ppr

        if bold_ref_run is not None and normal_ref_run is not None:
            break

    return bold_ref_run, normal_ref_run, normal_ppr_ref


def create_extra_sections(body, sections, entries, num_existing):
    """Create sections for entries beyond what the template has."""
    bold_ref, normal_ref, normal_ppr = find_format_references(sections)
    if normal_ref is None:
        print("Warning: Could not find formatting reference in section 1", file=sys.stderr)
        return

    # Find insertion point: the actual last w:p in the body
    # (can't use section["paras"][-1] because fill_leistung may have removed those elements)
    all_body_paras = body.findall(f"{{{W}}}p")
    insert_after = all_body_paras[-1]

    for i in range(num_existing, len(entries)):
        entry = entries[i]
        section_num = i + 1
        new_paras = []

        # 1. Empty paragraph BEFORE the title
        spacer_before_title = make_paragraph(normal_ppr)
        new_paras.append(spacer_before_title)

        # Title line (bold) - keepNext to stay with following content
        titel = entry.get("titel", DASH)
        title_p = make_paragraph(normal_ppr)
        if bold_ref is not None:
            title_p.append(make_run_with_rpr(bold_ref, f"{section_num}. {titel}"))
        else:
            title_p.append(make_run(f"{section_num}. {titel}", bold=True))
        set_keep_next(title_p)
        new_paras.append(title_p)

        # ID line - keepNext
        id_p = make_paragraph(normal_ppr)
        id_p.append(make_run_with_rpr(normal_ref, f"ID: {entry.get('dtad_id', DASH)}"))
        set_keep_next(id_p)
        new_paras.append(id_p)

        # Abgabetermin - keepNext
        abgabe_p = make_paragraph(normal_ppr)
        abgabe_p.append(make_run_with_rpr(normal_ref, f"Abgabetermin: {entry.get('abgabetermin', DASH)}"))
        set_keep_next(abgabe_p)
        new_paras.append(abgabe_p)

        # Ausführungsort - keepNext
        ort_p = make_paragraph(normal_ppr)
        ort_p.append(make_run_with_rpr(normal_ref, f"Ausführungsort: {entry.get('ausfuehrungsort', DASH)}"))
        set_keep_next(ort_p)
        new_paras.append(ort_p)

        # Ausführungsfrist - keepNext
        frist_p = make_paragraph(normal_ppr)
        frist_p.append(make_run_with_rpr(normal_ref, format_ausfuehrungsfrist(entry)))
        set_keep_next(frist_p)
        new_paras.append(frist_p)

        # 2. Empty paragraph BEFORE "Art und Umfang der Leistung" - keepNext (part of entry)
        spacer_before_leistung = make_paragraph(normal_ppr)
        set_keep_next(spacer_before_leistung)
        new_paras.append(spacer_before_leistung)

        # Art und Umfang der Leistung heading - keepNext
        leistung_heading = make_paragraph(normal_ppr)
        set_keep_next(leistung_heading)
        leistung_heading.append(make_run_with_rpr(normal_ref, "Art und Umfang der Leistung:"))
        new_paras.append(leistung_heading)

        # Leistung bullet points - all except last get keepNext
        leistung = entry.get("leistung", DASH)
        lines = split_leistung_to_bullets(leistung)
        if lines:
            for idx, line in enumerate(lines):
                bullet_text = line if line.startswith("- ") else f"- {line}"
                bullet_p = make_paragraph(normal_ppr)
                # Clear any inherited keepNext
                bullet_ppr = bullet_p.find(f"{{{W}}}pPr")
                if bullet_ppr is not None:
                    kn = bullet_ppr.find(f"{{{W}}}keepNext")
                    if kn is not None:
                        bullet_ppr.remove(kn)
                # Add keepNext to all bullets EXCEPT the last one
                if idx < len(lines) - 1:
                    set_keep_next(bullet_p)
                set_hanging_indent(bullet_p)
                bullet_p.append(make_run_with_rpr(normal_ref, bullet_text))
                new_paras.append(bullet_p)

        # 3. ALWAYS add spacer after bullets (gap before next entry)
        # This ensures proper spacing between entries
        spacer_after_bullets = make_paragraph(normal_ppr)
        # Clear any inherited keepNext from spacer (allows page break between entries)
        spacer_ppr = spacer_after_bullets.find(f"{{{W}}}pPr")
        if spacer_ppr is not None:
            kn = spacer_ppr.find(f"{{{W}}}keepNext")
            if kn is not None:
                spacer_ppr.remove(kn)
        new_paras.append(spacer_after_bullets)

        # Insert all new paragraphs into the document body
        for np in new_paras:
            insert_after.addnext(np)
            insert_after = np


# ============================================================
# SPACING NORMALIZATION
# ============================================================

def normalize_entry_spacing(body):
    """Ensure exactly 2 blank paragraphs before every section title (except the first).

    Runs repeatedly until no more adjustments are needed, so any number of
    existing blank lines (1, 3, 4, ...) will be normalized to exactly 2.
    """
    section_re = re.compile(r"^\d+\.\s+\S")      # "1. Something"
    date_skip_re = re.compile(
        r"^\d+\.\s*(?:Januar|Februar|März|April|Mai|Juni|Juli|August|"
        r"September|Oktober|November|Dezember)\s+\d{4}", re.IGNORECASE
    )

    while True:
        all_paras = list(body.findall(f"{{{W}}}p"))
        adjusted = False

        for i, p in enumerate(all_paras):
            text = "".join(t.text or "" for t in p.findall(f".//{{{W}}}t")).strip()

            # Only care about section titles that have something before them
            if not (section_re.match(text) and not date_skip_re.match(text) and i > 0):
                continue

            # Walk backwards to find last non-blank paragraph
            j = i - 1
            while j >= 0:
                t2 = "".join(tt.text or "" for tt in all_paras[j].findall(f".//{{{W}}}t")).strip()
                if t2:
                    break
                j -= 1

            num_blanks = (i - 1) - j   # blank paragraphs between content and this title

            # Skip if already correct, or if title is at the very top (j < 0)
            if j < 0 or num_blanks == 2:
                continue

            # Remove all existing blank paragraphs before this title
            for k in range(j + 1, i):
                bp = all_paras[k]
                parent = bp.getparent()
                if parent is not None:
                    parent.remove(bp)

            # Insert exactly 2 blank paragraphs after the last content para
            ref_p = all_paras[j]
            ref_ppr = ref_p.find(f"{{{W}}}pPr")
            insert_after = ref_p
            for _ in range(2):
                spacer = make_paragraph(ref_ppr)
                sp_ppr = spacer.find(f"{{{W}}}pPr")
                if sp_ppr is not None:
                    kn = sp_ppr.find(f"{{{W}}}keepNext")
                    if kn is not None:
                        sp_ppr.remove(kn)
                insert_after.addnext(spacer)
                insert_after = spacer

            adjusted = True
            break   # Restart scan — indices changed

        if not adjusted:
            break


# ============================================================
# HEADER REPLACEMENTS (gewerk, region, date)
# ============================================================

def replace_paragraph_text(p, new_text):
    """Replace all text in a paragraph with new_text, preserving first run's formatting."""
    runs = p.runs
    if not runs:
        return
    ref_run_el = runs[0]._element
    clear_paragraph_text(p._element)
    p._element.append(make_run_with_rpr(ref_run_el, new_text))


def replace_header_fields(doc, gewerk, region):
    """Replace gewerk and region placeholders in the header area, and update date."""
    today = datetime.now()
    date_str = f"{today.day}. {GERMAN_MONTHS[today.month]} {today.year}"

    for p in doc.paragraphs:
        text = p.text.strip()

        # Replace gewerk (text may span multiple runs, so work on joined text)
        if gewerk and ("Gebäudereinigungsarbeiten" in text or "Gebäudereinigung" in text):
            new_text = text.replace("Gebäudereinigungsarbeiten", gewerk)
            new_text = new_text.replace("Gebäudereinigung", gewerk)
            replace_paragraph_text(p, new_text)

        # Replace region
        elif region and ("59759 Arnsberg" in text or "47051 Duisburg" in text):
            new_text = text.replace("59759 Arnsberg + 50 km", region)
            new_text = new_text.replace("47051 Duisburg + 50km", region)
            replace_paragraph_text(p, new_text)

        # Update date
        elif re.match(
            r"\d{1,2}\.\s+(?:Januar|Februar|März|April|Mai|Juni|Juli|August|"
            r"September|Oktober|November|Dezember)\s+\d{4}", text
        ):
            replace_paragraph_text(p, date_str)


# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: generate_docx.py <vorlage_path> <output_path>"}))
        sys.exit(1)

    vorlage_path = sys.argv[1]
    output_path = sys.argv[2]

    # Read input JSON from stdin
    try:
        input_data = json.loads(sys.stdin.read())
    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"Invalid JSON input: {e}"}))
        sys.exit(1)

    entries = input_data.get("entries", [])
    gewerk = input_data.get("gewerk", "")
    region = input_data.get("region", "")

    if not entries:
        print(json.dumps({"error": "No entries provided"}))
        sys.exit(1)

    try:
        # Load template
        doc = Document(vorlage_path)
        body = doc.element.body

        # 1. Replace header fields (gewerk, region, date)
        replace_header_fields(doc, gewerk, region)

        # 2. Detect sections in template
        sections = detect_sections(body)
        num_sections = len(sections)
        num_entries = len(entries)

        print(f"Template has {num_sections} sections, {num_entries} entries to fill",
              file=sys.stderr)

        # 3. Cut excess sections if fewer entries than template slots
        if num_entries < num_sections:
            sections = cut_excess_sections(body, sections, num_entries)
            print(f"Cut to {len(sections)} sections", file=sys.stderr)

        # 4. Fill existing sections
        for i, sec in enumerate(sections):
            if i < num_entries:
                fill_section(sec, entries[i], i + 1)

        # 5. Extend with new sections if more entries than template slots
        if num_entries > num_sections:
            create_extra_sections(body, sections, entries, num_sections)
            print(f"Extended with {num_entries - num_sections} extra sections",
                  file=sys.stderr)

        # 6. Normalize spacing: exactly 2 blank lines between every entry
        normalize_entry_spacing(body)

        # 7. Save output
        doc.save(output_path)
        print(json.dumps({"ok": True, "output": output_path, "sections_filled": num_entries}))

    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
