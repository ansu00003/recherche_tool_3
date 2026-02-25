#!/usr/bin/env python3
"""PDF Extraction Script using pdfplumber + regex + Claude validation"""

import sys
import json
import re
import io
import base64
import os
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    print(json.dumps({"error": "pdfplumber not installed. Run: pip install pdfplumber"}))
    sys.exit(1)

# ========================
# TEXT EXTRACTION
# ========================
def extract_text_from_pdf(pdf_bytes):
    """Extract text from PDF using pdfplumber"""
    text_parts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(page_text)
    except Exception as e:
        return "", str(e)
    return "\n\n".join(text_parts), None

# ========================
# DATE NORMALIZATION
# ========================
def normalize_date(date_str):
    """Normalize date to DD.MM.YYYY format"""
    if not date_str or date_str == "—":
        return "—"
    
    date_str = date_str.strip()
    
    # Already in DD.MM.YYYY format
    if re.match(r'^\d{2}\.\d{2}\.\d{4}$', date_str):
        return date_str
    
    # Try various date formats
    formats = [
        '%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y',
        '%d. %B %Y', '%d %B %Y', '%B %d, %Y'
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime('%d.%m.%Y')
        except ValueError:
            continue
    
    # Extract date pattern from string
    match = re.search(r'(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})', date_str)
    if match:
        d, m, y = match.groups()
        return f"{int(d):02d}.{int(m):02d}.{y}"
    
    return date_str

# ========================
# REGEX FIELD EXTRACTION
# ========================
def extract_fields_from_text(text):
    """Extract fields using layered regex patterns"""
    
    # Normalize text: collapse excessive whitespace but preserve structure
    normalized = re.sub(r'[ \t]+', ' ', text)
    normalized = re.sub(r'\n{3,}', '\n\n', normalized)
    
    fields = {
        'titel': '—',
        'dtad_id': '—',
        'abgabetermin': '—',
        'ausfuehrungsort': '—',
        'beginn': '—',
        'ende': '—',
        'duration': '—',
        'leistung': '—'
    }
    
    # --- TITEL ---
    titel_patterns = [
        r'Titel[:\s]*\n?\s*([^\n]+)',
        r'Bezeichnung[:\s]*\n?\s*([^\n]+)',
        r'Auftragsbezeichnung[:\s]*\n?\s*([^\n]+)',
        r'Gegenstand[:\s]*\n?\s*([^\n]+)',
        r'Betreff[:\s]*\n?\s*([^\n]+)',
    ]
    for pattern in titel_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = match.group(1).strip()
            if len(val) > 5 and val != '—':
                fields['titel'] = val[:200]
                break
    
    # --- DTAD-ID ---
    dtad_patterns = [
        r'ID[::\s]*([\d]+)',  # Simple ID: 24030532
        r'DTAD[- ]?ID[::\s]*([A-Z0-9\-]+)',
        r'Vergabe[- ]?Nr\.?[::\s]*([A-Z0-9\-]+)',
        r'Ausschreibungs[- ]?ID[::\s]*([A-Z0-9\-]+)',
        r'Referenz[- ]?Nr\.?[::\s]*([A-Z0-9\-]+)',
        r'Aktenzeichen[::\s]*([A-Z0-9\-/]+)',
    ]
    for pattern in dtad_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = match.group(1).strip()
            if len(val) >= 4:
                fields['dtad_id'] = val
                break
    
    # --- ABGABETERMIN ---
    abgabe_patterns = [
        r'Abgabetermin[::\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})',  # Primary
        r'Schlusstermin[::\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})',
        r'Angebotsfrist[::\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})',
        r'Frist\s+(?:zur\s+)?Angebotsabgabe[::\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Abgabefrist[::\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Einreichungsfrist[::\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Ablauf\s+der\s+Frist[::\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
    ]
    for pattern in abgabe_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            fields['abgabetermin'] = normalize_date(match.group(1))
            break
    
    # --- AUSFUEHRUNGSORT ---
    # Helper function to clean extracted location
    def clean_location(val):
        if not val or val == '—':
            return '—'
        # Remove NUTS codes and metadata
        val = re.sub(r'\s*\([A-Z]{2,3}\d+\).*', '', val)  # (DE300) and everything after
        val = re.sub(r'\s*Land:.*', '', val, flags=re.IGNORECASE)  # "Land: Deutschland" and after
        val = re.sub(r'\s*NUTS.*', '', val, flags=re.IGNORECASE)  # NUTS codes
        val = re.sub(r'\s*Gliederung.*', '', val, flags=re.IGNORECASE)
        val = re.sub(r'\s*Postleitzahl:.*', '', val, flags=re.IGNORECASE)
        val = re.sub(r'\s*\d+\.\d+\.\d+.*', '', val)  # Section numbers like 2.1.2
        val = re.sub(r'\s*Allgemeine.*', '', val, flags=re.IGNORECASE)
        val = val.strip()
        # Check if result is meaningful (not just metadata)
        if len(val) < 3 or val.upper().startswith('NUTS') or val.startswith('Land'):
            return '—'
        
        # CAP LENGTH for multi-address cases (safety net)
        if len(val) > 150:
            # Keep first ~120 chars and cut at last clean separator
            first_part = val[:120]
            last_sep = max(first_part.rfind(';'), first_part.rfind(','))
            if last_sep > 50:
                return first_part[:last_sep].strip() + ' u.a.'
            return first_part.strip() + ' u.a.'
        return val
    
    # Step 1: Try Erfüllungsort with multi-line capture
    erfuellungsort_match = re.search(
        r'Erfüllungsort[^:\n]*:\s*([^\n]+(?:\n(?!\s*(?:NUTS|Geschätzte|Zusätzliche|Hauptklassifizierung|CPV|Region|Postleitzahl|Verfahrensart|Allgemeine|II\.|III\.|IV\.|Land:|Gliederung))[^\n]*)*)',
        text, re.IGNORECASE
    )
    if erfuellungsort_match:
        raw_val = erfuellungsort_match.group(1).strip()
        # Join lines with commas and clean
        lines = [l.strip() for l in raw_val.split('\n') if l.strip()]
        joined = ', '.join(lines)
        fields['ausfuehrungsort'] = clean_location(joined)
    
    # Step 2: Fallback to Ausführungsort patterns
    if fields['ausfuehrungsort'] == '—':
        ort_patterns = [
            r'Ausführungsort[^:\n]*:\s*([^\n(]+)',  # Stop at ( for NUTS codes
            r'Hauptort\s+der\s+Ausführung[^:\n]*:\s*([^\n(]+)',
            r'Leistungsort[:\s]*([^\n(]+)',
            r'Ort\s+der\s+(?:Leistungs)?ausführung[:\s]*([^\n(]+)',
        ]
        for pattern in ort_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                cleaned = clean_location(match.group(1))
                if cleaned != '—':
                    fields['ausfuehrungsort'] = cleaned
                    break
    
    # Step 3: Final fallback to Region
    if fields['ausfuehrungsort'] == '—':
        region_match = re.search(r'Region[:\s]*([^\n(]+)', text, re.IGNORECASE)
        if region_match:
            cleaned = clean_location(region_match.group(1))
            if cleaned != '—':
                fields['ausfuehrungsort'] = cleaned
    
    # --- BEGINN / ENDE ---
    # FIRST: Try date range pattern (most reliable for paired dates)
    range_patterns = [
        r'Beginn[:\s]*([\d./]+)[\s]*[-–]+[\s]*Ende[:\s]*([\d./]+)',  # Beginn: X - Ende: Y
        r'Ausführungsfrist[^\n]*Beginn[:\s]*([\d./]+)[\s]*[-–]+[\s]*Ende[:\s]*([\d./]+)',  # In context
        r'vom[:\s]*([\d./]+)[\s]*[-–bis]+[\s]*([\d./]+)',  # vom X bis Y
        r'(\d{1,2}[./]\d{1,2}[./]\d{4})[\s]*[-–bis]+[\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',  # X - Y
    ]
    for pattern in range_patterns:
        range_match = re.search(pattern, text, re.IGNORECASE)
        if range_match:
            beginn_val = range_match.group(1).strip()
            ende_val = range_match.group(2).strip()
            if re.match(r'\d{1,2}[./]\d{1,2}[./]\d{4}', beginn_val):
                fields['beginn'] = normalize_date(beginn_val)
            if re.match(r'\d{1,2}[./]\d{1,2}[./]\d{4}', ende_val):
                fields['ende'] = normalize_date(ende_val)
            if fields['beginn'] != '—' and fields['ende'] != '—':
                break
    
    # SECOND: Try individual patterns if still missing
    if fields['beginn'] == '—' or fields['ende'] == '—':
        date_pattern_pairs = [
            # Standard patterns with . or / separator
            (r'Datum\s+des\s+Beginns[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})', 
             r'(?:Enddatum\s+der\s+Laufzeit|Datum\s+des\s+Endes)[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
            (r'Beginn[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
             r'Ende[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
            (r'Leistungsbeginn[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
             r'Leistungsende[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
            (r'ab[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
             r'bis[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
        ]
        
        for beginn_pattern, ende_pattern in date_pattern_pairs:
            if fields['beginn'] == '—':
                beginn_match = re.search(beginn_pattern, text, re.IGNORECASE)
                if beginn_match:
                    fields['beginn'] = normalize_date(beginn_match.group(1))
            if fields['ende'] == '—':
                ende_match = re.search(ende_pattern, text, re.IGNORECASE)
                if ende_match:
                    fields['ende'] = normalize_date(ende_match.group(1))
            if fields['beginn'] != '—' and fields['ende'] != '—':
                break
    
    # --- DURATION ---
    # Extract contract duration (e.g., "48 Monate", "4 Jahre", "Laufzeit: 24 Monate")
    duration_patterns = [
        r'Laufzeit[:\s]*(\d+\s*(?:Monate?|Jahre?|Wochen?))',
        r'Vertragslaufzeit[:\s]*(\d+\s*(?:Monate?|Jahre?|Wochen?))',
        r'Dauer[:\s]*(\d+\s*(?:Monate?|Jahre?|Wochen?))',
        r'Dauerschuldverhältnis[^\d]*(\d+\s*(?:Monate?|Jahre?|Wochen?))',
        r'(\d+\s*(?:Monate?|Jahre?))\s*(?:Vertragslaufzeit|Laufzeit)',
        r'für\s+(?:eine\s+)?(?:Dauer\s+von\s+)?(\d+\s*(?:Monate?|Jahre?))',
    ]
    for pattern in duration_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            duration_val = match.group(1).strip()
            # Normalize: "48 Monate" or "4 Jahre"
            duration_val = re.sub(r'\s+', ' ', duration_val)
            fields['duration'] = duration_val
            break
    
    # Also try to extract from KW (Kalenderwoche) patterns
    if fields['duration'] == '—':
        kw_match = re.search(r'KW\s*(\d+)(?:/\d+)?\s*[-–]\s*KW\s*(\d+)', text, re.IGNORECASE)
        if kw_match:
            start_kw = int(kw_match.group(1))
            end_kw = int(kw_match.group(2))
            weeks = end_kw - start_kw if end_kw > start_kw else (52 - start_kw + end_kw)
            if weeks > 0:
                if weeks >= 48:
                    fields['duration'] = f"{weeks // 4} Monate"
                else:
                    fields['duration'] = f"{weeks} Wochen"
    
    # --- LEISTUNG ---
    # Priority-based extraction - EXACT spec table
    leistung_source = 'none'
    leistung_val = None
    
    # ========== PRIORITY 1: "Art und Umfang der Leistung" ==========
    # Min: 10 chars | Always tried first
    art_umfang_patterns = [
        r'Art\s+und\s+Umfang\s+der\s+Leistung[::\s]*([\s\S]{10,5000}?)(?=\n\s*(?:Ausführungsort|Ausführungsfrist|Abgabetermin|ID:|DTAD|Vergabestelle)|\n\n\n)',
        r'Art\s+und\s+Umfang\s+der\s+Leistung[::\s]*([\s\S]{10,5000}?)(?=\n\d+\.\s)',
        r'Art\s+und\s+Umfang\s+der\s+Leistung[::\s]*([\s\S]{10,5000}?)(?=\n\s*Art\s+und\s+Umfang)',
        r'Art\s+und\s+Umfang\s+der\s+Leistung[::\s]*([\s\S]{10,5000}?)(?=\n\n[A-ZÄÖÜ])',
    ]
    
    for i, pattern in enumerate(art_umfang_patterns):
        art_umfang_match = re.search(pattern, text, re.IGNORECASE)
        if art_umfang_match:
            art_umfang_section = art_umfang_match.group(1).strip()
            if len(art_umfang_section) >= 10:
                leistung_val = art_umfang_section
                leistung_source = 'art_und_umfang'
                print(f"PRIORITY 1 (pattern {i+1}): Art und Umfang: {len(art_umfang_section)} chars", file=sys.stderr)
                break
    
    # ========== PRIORITY 1.5: Only if Kurzbeschreibung < 100 chars ==========
    if leistung_source == 'none':
        kurz_check = re.search(r'Kurzbeschreibung[::\s]*([^\n]{5,200})', text, re.IGNORECASE)
        kurz_val = kurz_check.group(1).strip() if kurz_check else ''
        
        if len(kurz_val) < 100:
            # --- PRIORITY 1.5a: Vollständige Bekanntmachung → Beschreibung ---
            # Min: 50 chars
            voll_match = re.search(r'Vollständige\s+Bekanntmachung([\s\S]{50,50000}?)(?=\n\n\d+\.\s|\Z)', text, re.IGNORECASE)
            if voll_match:
                voll_section = voll_match.group(1)
                desc_match = re.search(r'Beschreibung[::\s]*\n([\s\S]{50,3000}?)(?=\n\s*(?:Interne\s+Kennung|Hauptklassifizierung|CPV|Erfüllungsort|Geschätzte|Dauer)|\Z)', voll_section, re.IGNORECASE)
                if desc_match:
                    leistung_val = desc_match.group(1).strip()
                    leistung_source = 'vollstaendige_beschreibung'
                    print(f"PRIORITY 1.5a: Vollständige → Beschreibung: {len(leistung_val)} chars", file=sys.stderr)
            
            # --- PRIORITY 1.5b: "5. Los" → "5.1 Los: LOT-" → Beschreibung ---
            # Min: 50 chars | Only if 1.5a failed
            if leistung_source == 'none':
                los_match = re.search(r'5\.\s*Los[\s\S]{0,100}?5\.1\s+Los:\s*LOT-[\s\S]{0,200}?Beschreibung[::\s]*\n([\s\S]{50,3000}?)(?=\n\s*(?:Interne\s+Kennung|Hauptklassifizierung|CPV|Erfüllungsort|Geschätzte|Dauer)|\Z)', text, re.IGNORECASE)
                if los_match:
                    leistung_val = los_match.group(1).strip()
                    leistung_source = 'los_beschreibung'
                    print(f"PRIORITY 1.5b: 5. Los → 5.1 Los: LOT- → Beschreibung: {len(leistung_val)} chars", file=sys.stderr)
            
            # --- PRIORITY 1.5c: Titel as heading → text below ---
            # Min: 20 chars | Only if 1.5a and 1.5b failed
            if leistung_source == 'none' and kurz_val:
                # Use Kurzbeschreibung as title, find text after it
                escaped_title = re.escape(kurz_val[:50])
                titel_match = re.search(escaped_title + r'[\s\S]{0,50}?\n([\s\S]{20,2000}?)(?=\n\n|\Z)', text, re.IGNORECASE)
                if titel_match:
                    leistung_val = titel_match.group(1).strip()
                    leistung_source = 'titel_heading'
                    print(f"PRIORITY 1.5c: Titel as heading → text below: {len(leistung_val)} chars", file=sys.stderr)
    
    # ========== PRIORITY 2: "Kurzbeschreibung" ==========
    # Min: 10 chars | Only if all 1.x failed
    if leistung_source == 'none':
        kurz_match = re.search(r'Kurzbeschreibung[::\s]*([^\n]{10,500})', text, re.IGNORECASE)
        if kurz_match:
            kurz_val = kurz_match.group(1).strip()
            if len(kurz_val) >= 10:
                leistung_val = kurz_val
                leistung_source = 'kurzbeschreibung'
                print(f"PRIORITY 2: Kurzbeschreibung: {len(kurz_val)} chars", file=sys.stderr)
    
    # ========== PRIORITY 3: "Beschreibung" ==========
    # Pattern: Beschreibung[::]\n(capture 10-3000 chars)
    # Stops at: Interne Kennung | Hauptklassifizierung | CPV | Erfüllungsort | Geschätzte | Dauer
    # Min: 10 chars | Only if P2 failed
    if leistung_source == 'none':
        besch_match = re.search(r'Beschreibung[::\s]*\n([\s\S]{10,3000}?)(?=\n\s*(?:Interne\s+Kennung|Hauptklassifizierung|CPV|Erfüllungsort|Geschätzte|Dauer)|\Z)', text, re.IGNORECASE)
        if besch_match:
            leistung_val = besch_match.group(1).strip()
            leistung_source = 'beschreibung'
            print(f"PRIORITY 3: Beschreibung: {len(leistung_val)} chars", file=sys.stderr)
    
    # ========== PRIORITY 4: Broad fallback ==========
    # Tries: Kurzbeschreibung:? → Beschreibung:? → Leistungen:?
    # Takes first paragraph block (double newline), caps at 2000 chars
    # Min: 10 chars | Last resort
    if leistung_source == 'none':
        fallback_patterns = [
            r'Kurzbeschreibung[::\s]*([\s\S]{10,2000}?)(?=\n\n|\Z)',
            r'Beschreibung[::\s]*([\s\S]{10,2000}?)(?=\n\n|\Z)',
            r'Leistungen[::\s]*([\s\S]{10,2000}?)(?=\n\n|\Z)',
        ]
        for fb_pattern in fallback_patterns:
            fb_match = re.search(fb_pattern, text, re.IGNORECASE)
            if fb_match:
                fb_val = fb_match.group(1).strip()
                if len(fb_val) >= 10:
                    leistung_val = fb_val[:2000]
                    leistung_source = 'fallback'
                    print(f"PRIORITY 4: Broad fallback: {len(leistung_val)} chars", file=sys.stderr)
                    break
    
    # Clean up the extracted value
    if leistung_val:
        leistung_val = re.sub(r'[ \t]+', ' ', leistung_val)  # Collapse spaces but keep newlines
        leistung_val = re.sub(r'Kategorien?:.*$', '', leistung_val, flags=re.IGNORECASE | re.MULTILINE)
        leistung_val = re.sub(r'CPV[- ]?Code.*$', '', leistung_val, flags=re.IGNORECASE | re.MULTILINE)
        leistung_val = re.sub(r'Klassifizierung.*$', '', leistung_val, flags=re.IGNORECASE | re.MULTILINE)
        leistung_val = re.sub(r'Hauptklassifizierung.*$', '', leistung_val, flags=re.IGNORECASE | re.MULTILINE)
        if len(leistung_val) >= 10:
            fields['leistung'] = leistung_val[:1500]
    
    # Store source for Claude
    fields['_leistung_source'] = leistung_source
    
    return fields

# ========================
# CLAUDE VALIDATION
# ========================
def validate_with_claude(fields, text, api_key):
    """Format ONLY leistung field using Claude API - other fields stay as regex extracted"""
    try:
        import anthropic
    except ImportError:
        return fields  # Fall back to regex-only
    
    if not api_key:
        return fields
    
    # Get leistung source for Claude prompt
    leistung_source = fields.pop('_leistung_source', 'none')
    
    # Get raw leistung text from regex extraction
    raw_leistung = fields.get('leistung', '—')
    
    # If leistung is empty/dash, nothing to format
    if not raw_leistung or raw_leistung == '—':
        return fields
    
    # Claude ONLY formats leistung - doesn't touch other fields
    prompt = f"""Formatiere den folgenden Leistungstext als Stichpunkte.

PDF-TEXT (erste 4000 Zeichen zur Prüfung):
{text[:4000]}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ROHER LEISTUNGSTEXT:
{raw_leistung}

Quelle: {leistung_source}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FORMATIERUNGSREGELN:
1. Jede Zeile beginnt mit "- " (Bindestrich + Leerzeichen)
2. Trenne die Zeilen mit Zeilenumbrüchen (\n)
3. Kein Fließtext-Block! Alles als einzelne, klare Stichpunkte

AUSSCHLÜSSE - ENTFERNE IMMER:
- Zeilen mit "Kategorien:"
- CPV-Codes, Klassifizierungen, Metadaten
- "Dauerschuldverhältnis" → NIEMALS im Ergebnis!
- Vertragslaufzeit (→ gehört in Ausführungsfrist)
- Standorte/Adressen (→ gehört in Ausführungsort)
- Datumsangaben in Klammern
- Qualitäts-/Umweltstandards (ISO 9001, ISO 14001)
- Vertragsbedingungen
- E-Mail / E-Mail-Adressen (z.B. vergabestelle@...)
- Auftraggeber (z.B. "Auftraggeber: Stadt...")
- Rechtsform (z.B. "Rechtsform: Lokale Gebietskörperschaft")
- Tätigkeit (z.B. "Tätigkeit: Allgemeine öffentliche Verwaltung")
- Vergabeordnung (z.B. "Vergabeordnung: Bauauftrag (VOB)")
- ID-Nummern (z.B. "ID: 24030532")
- Aufteilung in Lose
- Dauer der Leistungen
- Vorgesehener Ausführungszeitraum / Ausführungszeitraum
- Nebenangebote
- Mehrere Hauptangebote
- Vergabeunterlagen
- Verschwiegenheitserklärung
- Fehlende Unterlagen
- Bindefrist
- Zuschlagskriterien
- Sicherheitsleistung
- Bietergemeinschaft

ZEILENUMBRUCH-REGELN (KRITISCH):

a) ZUSAMMENFÜHREN BEI "und/oder/sowie":
   Wenn eine Zeile mit "und ", "oder ", "sowie " beginnt,
   MUSS sie mit der VORHERIGEN Zeile zusammengeführt werden!
   
   FALSCH: "- Unterhalts-\n- und Grundreinigung"
   RICHTIG: "- Unterhalts- und Grundreinigung"

b) ZUSAMMENFÜHREN BEI BINDESTRICH AM ENDE:
   Wenn eine Zeile mit Bindestrich endet (z.B. "Unterhalts-", "Glas-"),
   gehört die nächste Zeile IMMER dazu!

c) ALLE LOSE BEIBEHALTEN (KRITISCH!):
   "Los X:" NIEMALS entfernen oder ändern!
   
   Wenn mehrere Lose vorhanden (Los 1, Los 2, Los 3, Los 4, etc.),
   MUSS jedes "Los X:" als eigene Zeile erhalten bleiben!
   
   RICHTIG:
   - Los 1: Garten- und Landschaftsbauarbeiten
   - Los 2: Erd- und Pflasterarbeiten
   
   FALSCH:
   - Garten- und Landschaftsbauarbeiten (Los fehlt!)
   - "Die Leistung wird in Losen ausgeschrieben" (zusammenfassend statt einzeln)

d) Mengenangaben und Details beibehalten (z.B. "ca. 2.262.700 m²/Jahr")

Antworte NUR mit dem formatierten Text (Stichpunkte). Kein JSON, keine Erklärungen."""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response_text = response.content[0].text.strip()
        
        # Claude returns plain text bullets - use directly as leistung
        if response_text:
            # Clean up: ensure each line starts with "- "
            formatted_lines = []
            for line in response_text.split('\n'):
                line = line.strip()
                if line:
                    if not line.startswith('- '):
                        line = f'- {line}'
                    formatted_lines.append(line)
            
            if formatted_lines:
                fields['leistung'] = '\n'.join(formatted_lines)
        
        # POST-PROCESS: Force merge bullet points starting with "und ", "oder ", "sowie "
        if 'leistung' in fields and fields['leistung'] != '—':
            fields['leistung'] = merge_bullet_points(fields['leistung'])
            print(f"Claude formatted leistung: {len(fields['leistung'])} chars", file=sys.stderr)
    except Exception as e:
        print(f"Claude formatting error: {e}", file=sys.stderr)
    
    return fields


def merge_bullet_points(text):
    """Force merge bullet points that start with 'und ', 'oder ', 'sowie ' with previous line."""
    if not text:
        return text
    
    lines = text.split('\n')
    merged = []
    
    for line in lines:
        stripped = line.strip()
        # Remove leading "- " for processing
        content = stripped[2:] if stripped.startswith('- ') else stripped
        
        # Check if this line should merge with previous
        if merged and content and content.lower().startswith(('und ', 'oder ', 'sowie ')):
            # Merge with previous line
            prev = merged[-1]
            # Remove trailing "- " prefix if present for clean merge
            if prev.endswith('-'):
                # Previous ends with hyphen, just append
                merged[-1] = prev + ' ' + content
            else:
                # Add hyphen before "und/oder/sowie"
                merged[-1] = prev + '- ' + content
        elif stripped:
            merged.append(stripped if stripped.startswith('- ') else f'- {stripped}')
    
    return '\n'.join(merged)

# ========================
# MAIN
# ========================
def main():
    # Read base64 PDF from stdin or argument
    if len(sys.argv) > 1:
        # File path provided
        with open(sys.argv[1], 'rb') as f:
            pdf_bytes = f.read()
    else:
        # Read base64 from stdin
        input_data = sys.stdin.read().strip()
        if input_data.startswith('data:'):
            # Remove data URL prefix
            input_data = input_data.split(',', 1)[1]
        try:
            pdf_bytes = base64.b64decode(input_data)
        except Exception as e:
            print(json.dumps({"error": f"Invalid base64: {e}"}))
            sys.exit(1)
    
    # Extract text
    text, error = extract_text_from_pdf(pdf_bytes)
    if error:
        print(json.dumps({"error": f"PDF extraction failed: {error}"}))
        sys.exit(1)
    
    if not text.strip():
        print(json.dumps({"error": "No text extracted from PDF (might be scanned/image-only)"}))
        sys.exit(1)
    
    # Extract fields with regex
    fields = extract_fields_from_text(text)
    
    # Validate with Claude if API key available
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if api_key:
        fields = validate_with_claude(fields, text, api_key)
    
    # ALWAYS apply bullet point merging to leistung (force merge "und"/"oder"/"sowie" lines)
    if fields.get('leistung') and fields['leistung'] != '—':
        fields['leistung'] = merge_bullet_points(fields['leistung'])
    
    # Output JSON
    print(json.dumps(fields, ensure_ascii=False))

if __name__ == '__main__':
    main()
