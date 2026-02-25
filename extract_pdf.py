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
        r'DTAD[- ]?ID[:\s]*([A-Z0-9\-]+)',
        r'Vergabe[- ]?Nr\.?[:\s]*([A-Z0-9\-]+)',
        r'Ausschreibungs[- ]?ID[:\s]*([A-Z0-9\-]+)',
        r'Referenz[- ]?Nr\.?[:\s]*([A-Z0-9\-]+)',
        r'Aktenzeichen[:\s]*([A-Z0-9\-/]+)',
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
        r'Frist\s+(?:zur\s+)?Angebotsabgabe[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Angebotsfrist[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Abgabefrist[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Einreichungsfrist[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Schlusstermin[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'Ablauf\s+der\s+Frist[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
    ]
    for pattern in abgabe_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            fields['abgabetermin'] = normalize_date(match.group(1))
            break
    
    # --- AUSFUEHRUNGSORT ---
    # Step 1: Try multi-line Erfüllungsort pattern (captures address until stop-words)
    erfuellungsort_pattern = r'Erfüllungsort[^:\n]*:\s*([^\n]+(?:\n(?!\s*(?:NUTS|Geschätzte|Zusätzliche|Hauptklassifizierung|CPV|Region|Postleitzahl|Verfahrensart|II\.|III\.|IV\.))[^\n]*)*?)(?=\n\s*(?:NUTS|Geschätzte|Zusätzliche|Hauptklassifizierung|CPV|Region|Postleitzahl|Verfahrensart|II\.|III\.|IV\.|$))'
    match = re.search(erfuellungsort_pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        val = match.group(1).strip()
        # Step 2: Join multi-line result with commas
        lines = [l.strip() for l in val.split('\n') if l.strip()]
        if lines:
            fields['ausfuehrungsort'] = ', '.join(lines)[:300]
    
    # Step 3: Fallback patterns if Erfüllungsort not found
    if fields['ausfuehrungsort'] == '—':
        ort_patterns = [
            r'Ausführungsort[^:\n]*:\s*([^\n]+)',
            r'Leistungsort[:\s]*\n?\s*([^\n]+)',
            r'Ort\s+der\s+(?:Leistungs)?ausführung[:\s]*\n?\s*([^\n]+)',
            r'Hauptort\s+der\s+Ausführung[:\s]*\n?\s*([^\n]+)',
        ]
        for pattern in ort_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                val = match.group(1).strip()
                if len(val) > 3 and val != '—' and not val.upper().startswith('NUTS'):
                    fields['ausfuehrungsort'] = val[:150]
                    break
    
    # Step 4: Final fallback to Region
    if fields['ausfuehrungsort'] == '—':
        region_match = re.search(r'Region[:\s]*([^\n]+)', text, re.IGNORECASE)
        if region_match:
            val = region_match.group(1).strip()
            if len(val) > 2 and not val.upper().startswith('NUTS'):
                fields['ausfuehrungsort'] = val[:100]
    
    # --- BEGINN / ENDE ---
    date_pattern_pairs = [
        (r'Datum\s+des\s+Beginns[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})', 
         r'Datum\s+des\s+Endes[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
        (r'Beginn[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
         r'Ende[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
        (r'Leistungsbeginn[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
         r'Leistungsende[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
        (r'ab[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
         r'bis[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
        (r'vom[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})',
         r'bis[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4})'),
    ]
    
    for beginn_pattern, ende_pattern in date_pattern_pairs:
        beginn_match = re.search(beginn_pattern, text, re.IGNORECASE)
        ende_match = re.search(ende_pattern, text, re.IGNORECASE)
        if beginn_match:
            fields['beginn'] = normalize_date(beginn_match.group(1))
        if ende_match:
            fields['ende'] = normalize_date(ende_match.group(1))
        if fields['beginn'] != '—' or fields['ende'] != '—':
            break
    
    # Try date range pattern: "01.01.2024 - 31.12.2024"
    if fields['beginn'] == '—' and fields['ende'] == '—':
        range_match = re.search(
            r'(\d{1,2}[./]\d{1,2}[./]\d{4})\s*[-–bis]+\s*(\d{1,2}[./]\d{1,2}[./]\d{4})',
            text, re.IGNORECASE
        )
        if range_match:
            fields['beginn'] = normalize_date(range_match.group(1))
            fields['ende'] = normalize_date(range_match.group(2))
    
    # --- LEISTUNG ---
    leistung_patterns = [
        r'Art\s+und\s+Umfang\s+der\s+Leistung[:\s]*\n?([\s\S]{20,1000}?)(?=\n\n|\nII\.|\nIII\.|\n[A-Z]{2,}:)',
        r'Kurzbeschreibung[:\s]*\n?([\s\S]{20,500}?)(?=\n\n|\n[A-Z])',
        r'Leistungsbeschreibung[:\s]*\n?([\s\S]{20,500}?)(?=\n\n)',
        r'Beschreibung[:\s]*\n?([\s\S]{20,500}?)(?=\n\n)',
    ]
    for pattern in leistung_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = match.group(1).strip()
            # Clean up the text
            val = re.sub(r'\s+', ' ', val)
            val = re.sub(r'Kategorien?:.*$', '', val, flags=re.IGNORECASE)
            val = re.sub(r'CPV[- ]?Code.*$', '', val, flags=re.IGNORECASE)
            if len(val) > 20:
                fields['leistung'] = val[:800]
                break
    
    return fields

# ========================
# CLAUDE VALIDATION
# ========================
def validate_with_claude(fields, text, api_key):
    """Validate and correct fields using Claude API"""
    try:
        import anthropic
    except ImportError:
        return fields  # Fall back to regex-only
    
    if not api_key:
        return fields
    
    # Truncate text for API
    text_preview = text[:15000] if len(text) > 15000 else text
    
    prompt = f"""Du bist ein Experte für deutsche Ausschreibungsdokumente. Analysiere den folgenden PDF-Text und korrigiere/vervollständige die extrahierten Felder.

EXTRAHIERTE FELDER (per Regex):
{json.dumps(fields, indent=2, ensure_ascii=False)}

PDF TEXT (Auszug):
{text_preview}

AUFGABEN:
1. Korrigiere falsche Feldwerte
2. Fülle fehlende Daten (beginn/ende) durch Suche im Text
3. WICHTIG FÜR "ausfuehrungsort":
   - Extrahiere die VOLLSTÄNDIGE Adresse inklusive Straße, Hausnummer, PLZ und Ort
   - Beispiel RICHTIG: "Gaußstr. 20, 42119 Wuppertal"
   - Beispiel FALSCH: nur "42119 Wuppertal" (fehlt Straße)
   - Wenn mehrere Zeilen: kombiniere sie zu einer vollständigen Adresse
   - Ignoriere NUTS-Codes, CPV-Codes und andere Metadaten
4. Formatiere leistung als saubere Stichpunkte (Zeilen mit - am Anfang)
5. Entferne Metadaten wie "Kategorien:", CPV-Codes, etc.

Antworte NUR mit einem validen JSON-Objekt mit denselben Feldnamen. Keine Erklärungen."""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response_text = response.content[0].text.strip()
        
        # Extract JSON from response
        json_match = re.search(r'\{[\s\S]*\}', response_text)
        if json_match:
            claude_fields = json.loads(json_match.group())
            # Merge: Claude values override regex if not empty
            for key in fields:
                if key in claude_fields:
                    val = claude_fields[key]
                    if val and str(val).strip() and str(val).strip() != '—':
                        fields[key] = str(val).strip()
    except Exception as e:
        print(f"Claude validation error: {e}", file=sys.stderr)
    
    return fields

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
    
    # Output JSON
    print(json.dumps(fields, ensure_ascii=False))

if __name__ == '__main__':
    main()
