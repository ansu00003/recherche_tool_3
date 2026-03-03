// server.js — KALKU Tender Tool v2 Backend
import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import multer from 'multer';
import nodemailer from 'nodemailer';
import { ConfidentialClientApplication } from '@azure/msal-node';
import fetch from 'node-fetch';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { execSync, exec, spawn } from 'child_process';
import { promisify } from 'util';
const execAsync = promisify(exec);
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const pdfParse = require('pdf-parse');

const app = express();
app.use(cors());
app.use((req, res, next) => { console.log(`[${new Date().toISOString().slice(11,19)}] ${req.method} ${req.url}`); next(); });
app.use(express.json({ limit: '500mb' }));

const DATA_DIR = process.env.DATA_DIR || './data';
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const ATTACHMENTS_DIR = path.join(DATA_DIR, 'attachments');
if (!fs.existsSync(ATTACHMENTS_DIR)) fs.mkdirSync(ATTACHMENTS_DIR, { recursive: true });

// ========================
// FILE STORES
// ========================
const EMAIL_CACHE_FILE = path.join(DATA_DIR, 'customer-emails.json');
const JOBS_FILE = path.join(DATA_DIR, 'jobs.json');

function loadJSON(file, fallback) {
  try { if (fs.existsSync(file)) return JSON.parse(fs.readFileSync(file, 'utf8')); } catch (e) { console.error(`Load error ${file}:`, e.message); }
  return fallback;
}
function saveJSON(file, data) { fs.writeFileSync(file, JSON.stringify(data, null, 2)); }

// ========================
// MICROSOFT GRAPH
// ========================
const msalClient = new ConfidentialClientApplication({
  auth: { clientId: process.env.MS_CLIENT_ID, authority: `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}`, clientSecret: process.env.MS_CLIENT_SECRET },
});
let graphToken = null, tokenExpiry = 0;

async function getGraphToken() {
  if (graphToken && Date.now() < tokenExpiry) return graphToken;
  const result = await msalClient.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
  graphToken = result.accessToken;
  tokenExpiry = Date.now() + (result.expiresOn - Date.now()) - 60000;
  return graphToken;
}

async function graphGet(url) {
  const token = await getGraphToken();
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.json();
}

async function listSharePointFolders(subfolder) {
  const driveId = process.env.MS_DRIVE_ID;
  const bp = process.env.MS_BASE_FOLDER;
  const fp = bp ? `${bp}/${subfolder}` : subfolder;
  let url = fp
    ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(fp)}:/children?$top=999&$filter=folder ne null`
    : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=999&$filter=folder ne null`;
  console.log(`  📂 Listing SharePoint: ${subfolder}...`);
  const allFolders = [];
  try {
    while (url) {
      const result = await graphGet(url);
      if (result.value) allFolders.push(...result.value.map((i) => i.name));
      url = result['@odata.nextLink'] || null;
    }
    console.log(`  ✅ Found ${allFolders.length} folders in ${subfolder}`);
    return allFolders;
  } catch (err) {
    console.error(`  ❌ SharePoint listing failed for ${subfolder}:`, err.message);
    return [];
  }
}

function parseFolderName(raw) {
  const parts = raw.split('_');
  if (parts.length >= 2) {
    const id = parts[0], nameParts = parts.slice(1);
    return { id, name: nameParts.join(' '), raw, searchText: `${id} ${nameParts.join(' ')} ${raw}`.toLowerCase() };
  }
  return { id: raw, name: raw, raw, searchText: raw.toLowerCase() };
}

// ========================
// SMTP
// ========================
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST, port: parseInt(process.env.SMTP_PORT || '587'), secure: false,
  auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
  tls: { rejectUnauthorized: false },
});
transporter.verify().then(() => console.log('✅ SMTP verified')).catch((e) => console.error('❌ SMTP:', e.message));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 100 * 1024 * 1024 } }); // 100MB per file

// ========================
// VORLAGE FILE
// ========================
const VORLAGE_FILE = path.join(DATA_DIR, 'vorlage.docx');

// ========================
// PDF EXTRACTION: pdf-parse → normalize → regex → Claude validation
// ========================

// Normalize text: collapse excessive whitespace but preserve paragraph breaks
function normalizeText(text) {
  return text
    .replace(/\r\n/g, '\n')           // Normalize line endings
    .replace(/[ \t]+/g, ' ')          // Collapse horizontal whitespace
    .replace(/ ?\n ?/g, '\n')         // Trim spaces around newlines
    .replace(/\n{3,}/g, '\n\n')       // Max 2 consecutive newlines
    .trim();
}

// Create a "flattened" version where soft line breaks become spaces (for multi-line field matching)
function flattenText(text) {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/([^\n])\n([^\n])/g, '$1 $2')  // Single newlines → spaces
    .replace(/[ \t]+/g, ' ')
    .trim();
}

// Extract field with multiple pattern attempts
function extractField(text, flatText, patterns, defaultVal = '—') {
  for (const pattern of patterns) {
    // Try on flattened text first (handles wrapped lines)
    let m = flatText.match(pattern);
    if (m?.[1]) return m[1].trim().substring(0, 500);
    // Then try original text
    m = text.match(pattern);
    if (m?.[1]) return m[1].trim().substring(0, 500);
  }
  return defaultVal;
}

// Extract date with common German formats
function extractDate(text, flatText, patterns) {
  const dateRegex = /(\d{1,2})[./](\d{1,2})[./](\d{2,4})/;
  for (const pattern of patterns) {
    let m = flatText.match(pattern) || text.match(pattern);
    if (m?.[1]) {
      const dateMatch = m[1].match(dateRegex);
      if (dateMatch) return dateMatch[0];
      return m[1].trim();
    }
  }
  return '—';
}

function extractFieldsFromText(rawText) {
  const text = normalizeText(rawText);
  const flatText = flattenText(rawText);

  // === TITEL ===
  const titel = extractField(text, flatText, [
    /Titel:\s*(.+?)(?=\n(?:Beschreibung|Kennung|Verfahrensart|DTAD|Kategorien|$))/is,
    /Titel:\s*([^\n]+)/i,
    /Bekanntmachung[:\s]+(.+?)(?=\n\n)/is,
    /Ausschreibung[:\s]+(.+?)(?=\n)/i,
  ]);

  // === DTAD-ID ===
  const dtad_id = extractField(text, flatText, [
    /DTAD[- ]?ID[:\s]*(\d{6,10})/i,
    /Referenznummer[:\s]*(\d{6,10})/i,
    /Aktenzeichen[:\s]*([A-Z0-9\-\/]+)/i,
    /Interne Kennung[:\s]*([^\n]+)/i,
  ]);

  // === ABGABETERMIN ===
  const abgabetermin = extractDate(text, flatText, [
    /Frist\s*(?:für\s*(?:die\s*)?)?Angebotsabgabe[:\s]*([^\n]+)/i,
    /Angebotsfrist[:\s]*([^\n]+)/i,
    /Schlusstermin[:\s]*([^\n]+)/i,
    /Abgabefrist[:\s]*([^\n]+)/i,
    /Einreichungsfrist[:\s]*([^\n]+)/i,
    /Angebote\s*(?:sind\s*)?(?:bis|einzureichen\s*bis)[:\s]*([^\n]+)/i,
    /Ablauf\s*der\s*Frist[:\s]*([^\n]+)/i,
  ]);

  // === AUSFÜHRUNGSORT ===
  let ausfuehrungsort = extractField(text, flatText, [
    /Erfüllungsort[^:\n]*[:\s]*([^\n]+)/i,
    /Ort\s*der\s*(?:Leistungs)?[Aa]usführung[:\s]*([^\n]+)/i,
    /Ausführungsort[:\s]*([^\n]+)/i,
    /Leistungsort[:\s]*([^\n]+)/i,
    /NUTS[- ]?Code[:\s]*([A-Z0-9]+)[^\n]*([^\n]*)/i,
    /Region[:\s]*([^\n]+)/i,
    /Postleitzahl[:\s]*(\d{5})[^\n]*([^\n]*)/i,
  ]);
  // Clean up NUTS codes - extract city if present
  if (ausfuehrungsort.match(/^DE[A-Z0-9]+$/)) {
    const cityMatch = flatText.match(/(?:Ort|Stadt)[:\s]*([A-Za-zäöüÄÖÜß\s\-]+)/i);
    if (cityMatch) ausfuehrungsort = cityMatch[1].trim();
  }

  // === BEGINN / ENDE ===
  let beginn = '—', ende = '—';
  
  // Pattern pairs for start/end dates
  const datePairs = [
    [/Datum\s*des\s*Beginns[:\s]*([^\n]+)/i, /Enddatum(?:\s*der\s*Laufzeit)?[:\s]*([^\n]+)/i],
    [/Beginn\s*der\s*Ausführung[:\s]*([^\n]+)/i, /Ende\s*der\s*Ausführung[:\s]*([^\n]+)/i],
    [/Ausführungsbeginn[:\s]*([^\n]+)/i, /Ausführungsende[:\s]*([^\n]+)/i],
    [/Leistungsbeginn[:\s]*([^\n]+)/i, /Leistungsende[:\s]*([^\n]+)/i],
    [/Vertragsbeginn[:\s]*([^\n]+)/i, /Vertragsende[:\s]*([^\n]+)/i],
    [/Laufzeit\s*(?:ab|von)[:\s]*([^\n]+)/i, /Laufzeit\s*bis[:\s]*([^\n]+)/i],
    [/\bab[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i, /\bbis[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i],
    [/\bvom[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i, /\bbis[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i],
  ];

  for (const [startP, endP] of datePairs) {
    if (beginn === '—') beginn = extractDate(text, flatText, [startP]);
    if (ende === '—') ende = extractDate(text, flatText, [endP]);
    if (beginn !== '—' && ende !== '—') break;
  }

  // Try date range patterns
  if (beginn === '—' || ende === '—') {
    const rangePatterns = [
      /Ausführungsfrist[^:]*[:\s]*(?:Beginn[:\s]*)?(\d{1,2}[./]\d{1,2}[./]\d{2,4})\s*[-–bis]+\s*(?:Ende[:\s]*)?(\d{1,2}[./]\d{1,2}[./]\d{2,4})/is,
      /Laufzeit[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})\s*[-–bis]+\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/is,
      /Zeitraum[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{2,4})\s*[-–bis]+\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/is,
      /(\d{1,2}[./]\d{1,2}[./]\d{2,4})\s*[-–]\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/,
    ];
    for (const p of rangePatterns) {
      const m = flatText.match(p) || text.match(p);
      if (m) {
        if (beginn === '—') beginn = m[1].trim();
        if (ende === '—') ende = m[2].trim();
        break;
      }
    }
  }

  // Try duration in months
  if (beginn === '—' && ende === '—') {
    const durationMatch = flatText.match(/Laufzeit[:\s]*(\d+)\s*Monat/i) || text.match(/Laufzeit[:\s]*(\d+)\s*Monat/i);
    if (durationMatch) {
      beginn = `${durationMatch[1]} Monate`;
      ende = '—';
    }
  }

  // === LEISTUNG ===
  let leistung = '—';
  const leistungPatterns = [
    /Art\s*und\s*Umfang\s*der\s*Leistung[:\s]*([\s\S]{20,}?)(?=\n(?:Kategorien|CPV|Vergabe|Erfüllungsort|Ausführ|Laufzeit|\d+\.\s|$))/i,
    /Kurzbeschreibung[:\s]*([\s\S]{20,}?)(?=\n(?:Kategorien|CPV|Kennung|Verfahrensart|\d+\.\s|$))/i,
    /Beschreibung[:\s]*([\s\S]{20,}?)(?=\n(?:Kategorien|CPV|Kennung|\d+\.\s|$))/i,
    /Leistungsgegenstand[:\s]*([\s\S]{20,}?)(?=\n\n)/i,
    /Gegenstand\s*des\s*Auftrags[:\s]*([\s\S]{20,}?)(?=\n\n)/i,
  ];
  for (const p of leistungPatterns) {
    const m = text.match(p);
    if (m?.[1]) {
      leistung = m[1].trim()
        .split(/\n\s*\n/)[0]           // Take first paragraph
        .replace(/\s+/g, ' ')          // Normalize whitespace
        .substring(0, 2000);
      if (leistung.length > 30) break;
    }
  }

  return { titel, dtad_id, abgabetermin, ausfuehrungsort, beginn, ende, leistung };
}

async function validateWithAnthropic(fields, text) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.log('    ⚠️  No Anthropic API key, skipping AI validation');
    return fields;
  }

  // Only call Claude if we have missing fields
  const missingFields = Object.entries(fields).filter(([k, v]) => v === '—').map(([k]) => k);
  if (missingFields.length === 0) {
    console.log('    ✅ All fields extracted, skipping Claude');
    return fields;
  }

  try {
    console.log(`    🤖 Calling Claude for ${missingFields.length} missing fields: ${missingFields.join(', ')}`);
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 15000);
    
    const prompt = `Du bist ein Datenextraktor für deutsche DTAD-Ausschreibungs-PDFs.

Hier ist der PDF-Text (gekürzt):
${text.slice(0, 10000)}

Bereits extrahierte Felder: ${JSON.stringify(fields, null, 2)}

Felder mit "—" konnten nicht extrahiert werden. Versuche diese aus dem Text zu finden.
Behalte bereits extrahierte Werte bei, es sei denn sie sind offensichtlich falsch.

Gib NUR ein valides JSON-Objekt zurück mit diesen Feldern:
- titel: Titel der Ausschreibung
- dtad_id: DTAD-ID oder Referenznummer (nur Zahlen/Buchstaben)
- abgabetermin: Frist für Angebotsabgabe (Format: TT.MM.JJJJ)
- ausfuehrungsort: Ort der Leistung (Stadt, PLZ oder Region)
- beginn: Startdatum der Ausführung (Format: TT.MM.JJJJ)
- ende: Enddatum der Ausführung (Format: TT.MM.JJJJ)
- leistung: Nur die konkreten Bauleistungen/Arbeiten als Stichpunkte (max 500 Zeichen).
  NUR aufnehmen: Was wird gebaut/geliefert/ausgeführt? (z.B. "Dachabdichtungsarbeiten", "Los 1: Erdarbeiten")
  NICHT aufnehmen: Erfüllungsort, Straßenadressen (z.B. "Böblinger Straße 36-38"), Frist Angebotsabgabe/Abgabefrist, Vergabeplattform-Info, elektronische Vergabe, Bieterfragen, Download-Anweisungen, Registrierung, Support-Telefon, Kostenhinweise, Verfahrensart, Dokumententyp, Datumsangaben, URLs, E-Mail-Adressen, Auftraggeber, Eignungskriterien, Zuschlagskriterien, Vertragsbedingungen`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2024-10-22' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 800, messages: [{ role: 'user', content: prompt }] }),
      signal: controller.signal,
    });
    clearTimeout(timeout);

    const data = await response.json();
    if (!response.ok || data.error) {
      console.log('    ⚠️  Claude API error:', data.error?.message || response.status);
      return fields;
    }
    
    const textOut = (data.content || []).map((c) => (c.type === 'text' ? c.text : '')).join('');
    const jsonMatch = textOut.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      // Only update fields that were missing or are clearly better
      const merged = { ...fields };
      for (const [key, val] of Object.entries(parsed)) {
        if (val && val !== '—' && (fields[key] === '—' || !fields[key])) {
          merged[key] = String(val).trim();
        }
      }
      console.log('    ✅ Claude validation complete');
      return merged;
    }
    return fields;
  } catch (err) {
    console.log('    ⚠️  Claude validation skipped:', err.name === 'AbortError' ? 'timeout' : err.message);
    return fields;
  }
}

async function extractFromPdfBuffer(pdfBuffer) {
  console.log('    📄 Extracting with Python/pdfplumber...');
  
  // Write buffer to temp file
  const tempFile = path.join(DATA_DIR, '_tmp', `pdf_${Date.now()}.pdf`);
  fs.writeFileSync(tempFile, pdfBuffer);
  
  // Auto-detect Python: venv or system python3
  let pythonCmd = 'python3';
  if (fs.existsSync('./venv/bin/python3')) {
    pythonCmd = './venv/bin/python3';
  }
  
  try {
    // Call Python extractor
    const env = { ...process.env };
    const { stdout, stderr } = await execAsync(
      `${pythonCmd} extract_pdf.py "${tempFile}"`,
      { cwd: process.cwd(), env, timeout: 60000 }
    );
    
    if (stderr) console.log('    ⚠️  Python stderr:', stderr);
    
    const fields = JSON.parse(stdout.trim());
    if (fields.error) {
      throw new Error(fields.error);
    }
    
    const missing = Object.values(fields).filter(v => v === '—').length;
    console.log(`    ✅ Extracted ${7 - missing}/7 fields`);
    return fields;
  } catch (err) {
    console.error('    ❌ Python extraction failed:', err.message);
    // Fallback to Node.js pdf-parse
    console.log('    🔄 Falling back to pdf-parse...');
    const result = await pdfParse(pdfBuffer);
    const text = result.text || '';
    const fieldsRaw = extractFieldsFromText(text);
    const missing = Object.values(fieldsRaw).filter(v => v === '—').length;
    if (missing >= 3 && process.env.ANTHROPIC_API_KEY) {
      return await validateWithAnthropic(fieldsRaw, text);
    }
    return fieldsRaw;
  } finally {
    try { fs.unlinkSync(tempFile); } catch(e) {}
  }
}

// ========================
// DOCX GENERATION (Python-based with python-docx/lxml)
// ========================
async function generateDocxBuffer(vorlagePath, entries, gewerk, region) {
  const ts = Date.now();
  const inputPath = path.join(LOCAL_TMP_DIR, `input_${ts}.json`);
  const outputPath = path.join(LOCAL_TMP_DIR, `docx_${ts}.docx`);

  try {
    // Write input JSON to temp file (exec doesn't support stdin pipe)
    fs.writeFileSync(inputPath, JSON.stringify({ entries, gewerk: gewerk || '', region: region || '' }));

    console.log('    📝 Calling Python DOCX generator...');
    
    // Auto-detect Python: venv or system python3
    let pythonCmd = 'python3';
    if (fs.existsSync('./venv/bin/python3')) {
      pythonCmd = './venv/bin/python3';
    }
    
    const { stdout, stderr } = await execAsync(
      `${pythonCmd} generate_docx.py "${vorlagePath}" "${outputPath}" < "${inputPath}"`,
      { cwd: process.cwd(), timeout: 30000, shell: true }
    );
    if (stderr) console.log('    ℹ️  generate_docx:', stderr.trim());

    const result = JSON.parse(stdout.trim());
    if (result.error) throw new Error(result.error);

    const buffer = fs.readFileSync(outputPath);
    console.log(`    ✅ DOCX generated: ${(buffer.length / 1024).toFixed(1)} KB, ${entries.length} sections`);
    return buffer;
  } finally {
    try { fs.unlinkSync(inputPath); } catch (e) { /* ignore */ }
    try { fs.unlinkSync(outputPath); } catch (e) { /* ignore */ }
  }
}

// ========================
// DOCX → PDF CONVERSION
// ========================
const LOCAL_TMP_DIR = path.join(DATA_DIR, '_tmp');
if (!fs.existsSync(LOCAL_TMP_DIR)) fs.mkdirSync(LOCAL_TMP_DIR, { recursive: true });

async function convertDocxToPdf(docxBuffer, outputFilename) {
  const tmpDir = fs.mkdtempSync(path.join(LOCAL_TMP_DIR, 'tender-'));
  const docxPath = path.join(tmpDir, outputFilename);

  try {
    // Ensure docxBuffer is resolved (not a Promise) before writing
    const resolvedBuffer = await Promise.resolve(docxBuffer);
    fs.writeFileSync(docxPath, resolvedBuffer);
    console.log(`  🔄 Converting with LibreOffice: ${docxPath}`);
    // Kill any hanging soffice processes first
    try { execSync('pkill -9 soffice', { timeout: 2000 }); } catch (e) { /* ignore */ }

    // Detect LibreOffice path (macOS vs Linux)
    let sofficePath = 'soffice'; // Default to PATH
    if (fs.existsSync('/opt/homebrew/bin/soffice')) {
      sofficePath = '/opt/homebrew/bin/soffice'; // macOS Homebrew
    } else if (fs.existsSync('/usr/bin/soffice')) {
      sofficePath = '/usr/bin/soffice'; // Linux
    }

    // Run conversion async so we don't block the event loop
    await execAsync(`${sofficePath} --headless --nofirststartwizard --nologo --norestore --convert-to pdf --outdir "${tmpDir}" "${docxPath}"`, {
      timeout: 60000,
      env: { ...process.env, HOME: os.homedir() }
    });

    const pdfPath = path.join(tmpDir, outputFilename.replace(/\.docx$/i, '.pdf'));
    if (fs.existsSync(pdfPath)) {
      const pdfBuffer = fs.readFileSync(pdfPath);
      console.log(`  ✅ PDF conversion successful: ${(pdfBuffer.length / 1024).toFixed(1)} KB`);
      return pdfBuffer;
    }
    console.error('  ⚠️  PDF conversion: output file not found');
    return null;
  } catch (err) {
    console.error('  ⚠️  PDF conversion failed:', err.message);
    if (err.killed) console.error('  ⚠️  Process was killed (timeout or manual)');
    return null;
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

// ========================
// SHAREPOINT UPLOAD
// ========================
async function graphPut(url, buffer, contentType) {
  const token = await getGraphToken();
  const res = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType || 'application/octet-stream' },
    body: buffer,
  });
  if (!res.ok) throw new Error(`Graph PUT ${res.status}: ${await res.text()}`);
  return res.json();
}

async function uploadToSharePoint(folderPath, fileName, buffer) {
  const driveId = process.env.MS_DRIVE_ID;
  // Encode each path segment separately, not the whole path (slashes must stay as-is)
  const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
  const filePath = `${encodedPath}/${encodeURIComponent(fileName)}`;
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`;
  console.log(`  📤 Uploading to SharePoint: ${folderPath}/${fileName}`);
  return graphPut(url, buffer);
}

// ========================
// ROUTES
// ========================

app.get('/api/health', (req, res) => res.json({ status: 'ok', dryRun: process.env.DRY_RUN === 'true', hasVorlage: fs.existsSync(VORLAGE_FILE) }));

// --- PDF Extraction Endpoint ---
app.post('/api/extract', express.json({ limit: '500mb' }), async (req, res) => {
  try {
    const { pdfBase64 } = req.body;
    if (!pdfBase64) return res.status(400).json({ error: 'pdfBase64 required' });
    
    console.log('📄 /api/extract - Processing PDF...');
    
    // Decode base64
    let base64Data = pdfBase64;
    if (base64Data.includes(',')) base64Data = base64Data.split(',')[1];
    const pdfBuffer = Buffer.from(base64Data, 'base64');
    
    // Extract fields using Python/pdfplumber
    const fields = await extractFromPdfBuffer(pdfBuffer);
    
    console.log('✅ /api/extract - Extraction complete');
    res.json(fields);
  } catch (err) {
    console.error('❌ /api/extract error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// --- Vorlage Upload/Status ---
app.get('/api/vorlage', (req, res) => {
  const exists = fs.existsSync(VORLAGE_FILE);
  let size = 0, uploadedAt = null;
  if (exists) { const stat = fs.statSync(VORLAGE_FILE); size = stat.size; uploadedAt = stat.mtime.toISOString(); }
  res.json({ exists, size, uploadedAt });
});

app.post('/api/vorlage', upload.single('vorlage'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  fs.writeFileSync(VORLAGE_FILE, req.file.buffer);
  console.log(`📄 Vorlage uploaded: ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB)`);
  res.json({ ok: true, size: req.file.size, name: req.file.originalname });
});

// --- Pipeline Step 1 (NEW): Generate DOCX from pre-extracted entries ---
app.post('/api/pipeline/generate-docx', express.json({ limit: '500mb' }), async (req, res) => {
  console.log('📥 Pipeline generate-docx - received pre-extracted entries');
  try {
    const { kundeName, kundeId, gewerk, region, entries, sourcePdfs } = req.body;
    if (!entries?.length) return res.status(400).json({ error: 'No entries provided' });
    if (!fs.existsSync(VORLAGE_FILE)) return res.status(400).json({ error: 'Keine Vorlage hochgeladen.' });

    const pdfNames = entries.map((e) => e._filename || e.titel || 'PDF');
    console.log(`\n🔧 Generate DOCX for "${kundeName}" (${entries.length} entries)`);

    // 1. Generate DOCX
    console.log('  📝 Generating DOCX...');
    const _now = new Date();
    const _dateStr = `${String(_now.getFullYear()).slice(2)}${String(_now.getMonth()+1).padStart(2,'0')}${String(_now.getDate()).padStart(2,'0')}`;
    const _idPart = kundeId ? `${kundeId}_` : '';
    const baseName = `Recherche_${_idPart}${_dateStr}_${(kundeName || 'Kunde').replace(/\s+/g, '_')}`;
    const docxFilename = `${baseName}.docx`;
    const pdfFilename = `${baseName}.pdf`;
    const docxBuffer = await generateDocxBuffer(VORLAGE_FILE, entries, gewerk, region);

    // 2. Convert DOCX → PDF
    console.log('  📄 Converting DOCX → PDF...');
    const pdfOutputBuffer = await convertDocxToPdf(docxBuffer, docxFilename);
    if (pdfOutputBuffer) console.log(`  ✅ PDF generated: ${pdfFilename}`);
    else console.log('  ⚠️  PDF conversion failed, will use DOCX');

    // 3. Save to temp preview directory
    const previewId = Date.now().toString(36) + Math.random().toString(36).substr(2, 6);
    const previewDir = path.join(ATTACHMENTS_DIR, '_preview', previewId);
    fs.mkdirSync(previewDir, { recursive: true });

    const attachmentBuffer = pdfOutputBuffer || docxBuffer;
    const attachmentName = pdfOutputBuffer ? pdfFilename : docxFilename;
    fs.writeFileSync(path.join(previewDir, attachmentName), attachmentBuffer);
    fs.writeFileSync(path.join(previewDir, docxFilename), docxBuffer);

    // Save source PDFs temporarily for SharePoint upload
    const savedSourcePdfs = [];
    if (sourcePdfs && Array.isArray(sourcePdfs)) {
      const srcDir = path.join(previewDir, '_source_pdfs');
      fs.mkdirSync(srcDir, { recursive: true });
      for (const sp of sourcePdfs) {
        if (sp.name && sp.data) {
          let b64 = sp.data;
          if (b64.includes(',')) b64 = b64.split(',')[1];
          fs.writeFileSync(path.join(srcDir, sp.name), Buffer.from(b64, 'base64'));
          savedSourcePdfs.push(sp.name);
        }
      }
      console.log(`  📎 Saved ${savedSourcePdfs.length} source PDFs for SharePoint upload`);
    }

    // Save entries as metadata for the send step
    fs.writeFileSync(path.join(previewDir, '_meta.json'), JSON.stringify({ entries, pdfNames, docxFilename, pdfFilename, attachmentName, savedSourcePdfs }));

    console.log(`  ✅ Preview ready: ${previewId}/${attachmentName}`);
    res.json({ ok: true, previewId, attachmentName, entries, pdfNames });
  } catch (e) {
    console.error('❌ Generate-docx error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// --- Pipeline Step 1: Generate PDF only (for preview before sending) ---
app.post('/api/pipeline/generate', upload.array('pdfs', 50), async (req, res) => {
  console.log('📥 Pipeline generate - received multipart request');
  try {
    const { kundeName, kundeId, gewerk, region } = req.body;
    if (!req.files?.length) return res.status(400).json({ error: 'No PDFs provided' });
    if (!fs.existsSync(VORLAGE_FILE)) return res.status(400).json({ error: 'Keine Vorlage hochgeladen.' });

    const pdfNames = req.files.map((f) => f.originalname);
    console.log(`\n🔧 Generate started for "${kundeName}" (${req.files.length} PDFs)`);

    // 1. Extract data from PDFs (direct buffer from multer)
    console.log('  📄 Extracting PDF data...');
    const entries = [];
    for (let i = 0; i < req.files.length; i++) {
      console.log(`    → ${pdfNames[i]}...`);
      try {
        const data = await extractFromPdfBuffer(req.files[i].buffer);
        entries.push(data);
        console.log(`    ✅ ID: ${data.dtad_id} | ${data.titel?.substring(0, 50)}`);
      } catch (err) {
        console.error(`    ❌ Failed: ${pdfNames[i]}:`, err.message);
        entries.push({ titel: pdfNames[i], dtad_id: '—', abgabetermin: '—', ausfuehrungsort: '—', beginn: '—', ende: '—', leistung: '—' });
      }
    }

    // 2. Generate DOCX
    console.log('  📝 Generating DOCX...');
    const _now2 = new Date();
    const _dateStr2 = `${String(_now2.getFullYear()).slice(2)}${String(_now2.getMonth()+1).padStart(2,'0')}${String(_now2.getDate()).padStart(2,'0')}`;
    const baseName = `Recherche_${kundeId || ''}_${_dateStr2}_${(kundeName || 'Kunde').replace(/\s+/g, '_')}`;
    const docxFilename = `${baseName}.docx`;
    const pdfFilename = `${baseName}.pdf`;
    const docxBuffer = await generateDocxBuffer(VORLAGE_FILE, entries, gewerk, region);

    // 3. Convert DOCX → PDF
    console.log('  📄 Converting DOCX → PDF...');
    const pdfOutputBuffer = await convertDocxToPdf(docxBuffer, docxFilename);
    if (pdfOutputBuffer) console.log(`  ✅ PDF generated: ${pdfFilename}`);
    else console.log('  ⚠️  PDF conversion failed, will use DOCX');

    // 4. Save to temp preview directory
    const previewId = Date.now().toString(36) + Math.random().toString(36).substr(2, 6);
    const previewDir = path.join(ATTACHMENTS_DIR, '_preview', previewId);
    fs.mkdirSync(previewDir, { recursive: true });

    const attachmentBuffer = pdfOutputBuffer || docxBuffer;
    const attachmentName = pdfOutputBuffer ? pdfFilename : docxFilename;
    fs.writeFileSync(path.join(previewDir, attachmentName), attachmentBuffer);
    fs.writeFileSync(path.join(previewDir, docxFilename), docxBuffer);

    // Save entries as metadata for the send step
    fs.writeFileSync(path.join(previewDir, '_meta.json'), JSON.stringify({ entries, pdfNames, docxFilename, pdfFilename, attachmentName }));

    console.log(`  ✅ Preview ready: ${previewId}/${attachmentName}`);
    res.json({ ok: true, previewId, attachmentName, entries, pdfNames });
  } catch (e) {
    console.error('❌ Generate error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// --- Serve preview PDF ---
app.get('/api/pipeline/preview/:previewId/:filename', (req, res) => {
  const { previewId, filename } = req.params;
  const filePath = path.join(ATTACHMENTS_DIR, '_preview', previewId, filename);
  
  console.log(`📄 Preview request: ${previewId}/${filename}`);
  console.log(`   Full path: ${filePath}`);
  
  if (!fs.existsSync(filePath)) {
    console.error(`   ❌ File not found: ${filePath}`);
    // List what's in the preview directory
    const previewDir = path.join(ATTACHMENTS_DIR, '_preview', previewId);
    if (fs.existsSync(previewDir)) {
      console.log(`   📁 Contents of ${previewDir}:`, fs.readdirSync(previewDir));
    } else {
      console.log(`   📁 Preview directory doesn't exist: ${previewDir}`);
    }
    return res.status(404).json({ error: 'Preview not found', path: filePath });
  }
  
  // Get file size for Content-Length header
  const stat = fs.statSync(filePath);
  console.log(`   ✅ File found: ${(stat.size / 1024).toFixed(1)} KB`);
  
  const ext = path.extname(filename).toLowerCase();
  const ct = ext === '.pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  
  // Set headers for inline PDF viewing
  res.setHeader('Content-Type', ct);
  res.setHeader('Content-Length', stat.size);
  res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');
  res.setHeader('Cache-Control', 'no-cache');
  // Remove headers that might block PDF rendering
  res.removeHeader('X-Content-Type-Options');
  res.removeHeader('Content-Security-Policy');
  
  const stream = fs.createReadStream(filePath);
  stream.on('error', (err) => {
    console.error(`   ❌ Read stream error: ${err.message}`);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Failed to read file', message: err.message });
    }
  });
  stream.pipe(res);
});

// --- Pipeline Step 2: Send the pre-generated email ---
app.post('/api/pipeline/send', express.json(), async (req, res) => {
  try {
    const { previewId, to, subject, body, kundeRaw, kundeName, kundeId, kundeTyp, anrede, empfaengerName, signaturPerson, emailTyp, gewerk, region } = req.body;
    if (!previewId || !to || !subject || !body) return res.status(400).json({ error: 'previewId, to, subject, body required' });

    const previewDir = path.join(ATTACHMENTS_DIR, '_preview', previewId);
    if (!fs.existsSync(previewDir)) return res.status(404).json({ error: 'Preview expired or not found. Please regenerate.' });

    const meta = JSON.parse(fs.readFileSync(path.join(previewDir, '_meta.json'), 'utf8'));
    const attachmentBuffer = fs.readFileSync(path.join(previewDir, meta.attachmentName));
    const docxBuffer = fs.readFileSync(path.join(previewDir, meta.docxFilename));
    const attachmentType = meta.attachmentName.endsWith('.pdf') ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    const bcc = process.env.PIPEDRIVE_SMART_BCC;

    console.log(`\n✉️  Sending pre-generated email for "${kundeName}" — preview ${previewId}`);

    const job = {
      id: Date.now().toString(36) + Math.random().toString(36).substr(2, 4),
      to, subject, kundeRaw: kundeRaw || '', kundeName: kundeName || '', kundeId: kundeId || '',
      anrede: anrede || 'Herr', empfaengerName: empfaengerName || '',
      signaturPerson: signaturPerson || '', emailTyp: emailTyp || 'regular',
      gewerk: gewerk || '', region: region || '',
      pdfNames: meta.pdfNames, entries: meta.entries,
      generatedFile: meta.attachmentName,
      sentAt: new Date().toISOString(),
      reminders: [],
    };

    if (process.env.DRY_RUN === 'true') {
      console.log(`  🔸 DRY RUN — Would send: ${to} | ${subject} | Attachment: ${meta.attachmentName}`);
      job.dryRun = true;
    } else {
      const info = await transporter.sendMail({
        from: process.env.EMAIL_FROM, to, bcc, subject, text: body,
        attachments: [{ filename: meta.attachmentName, content: attachmentBuffer, contentType: attachmentType }],
      });
      console.log(`  ✅ Sent: ${info.messageId} → ${to}`);
      job.messageId = info.messageId;
    }

    // Save attachment for reminders
    const jobAttDir = path.join(ATTACHMENTS_DIR, job.id);
    fs.mkdirSync(jobAttDir, { recursive: true });
    fs.writeFileSync(path.join(jobAttDir, meta.attachmentName), attachmentBuffer);
    job.savedAttachments = [meta.attachmentName];

    // Save email+name to cache
    if (kundeRaw && to) {
      const c = loadJSON(EMAIL_CACHE_FILE, {});
      const prev = typeof c[kundeRaw] === 'string' ? { email: c[kundeRaw], name: '' } : (c[kundeRaw] || { email: '', name: '' });
      prev.email = to;
      if (empfaengerName) prev.name = empfaengerName;
      c[kundeRaw] = prev;
      saveJSON(EMAIL_CACHE_FILE, c);
    }

    // Upload to SharePoint
    const today = new Date();
    const yy = String(today.getFullYear()).slice(-2);
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const dateFolder = `${yy}${mm}${dd}`;
    const spDir = kundeTyp === 'Interessenten' ? 'Interessenten' : 'Aktive_Kunden';
    const spPath = `${spDir}/${kundeRaw}/${dateFolder}`;

    if (process.env.DRY_RUN === 'true') {
      console.log(`  🔸 DRY RUN — Would upload to SharePoint: ${spPath}/`);
      if (meta.savedSourcePdfs?.length) console.log(`  🔸 DRY RUN — Would upload ${meta.savedSourcePdfs.length} source PDFs to ${spPath}/PDFs/`);
      job.sharePointPath = spPath;
    } else {
      try {
        console.log(`  📁 Uploading to SharePoint: ${spPath}/`);
        await uploadToSharePoint(spPath, meta.docxFilename, docxBuffer);
        console.log(`    ✅ ${meta.docxFilename}`);
        if (meta.attachmentName.endsWith('.pdf')) {
          await uploadToSharePoint(spPath, meta.attachmentName, attachmentBuffer);
          console.log(`    ✅ ${meta.attachmentName}`);
        }
        // Upload source PDFs to PDFs/ subfolder
        if (meta.savedSourcePdfs?.length) {
          const srcDir = path.join(previewDir, '_source_pdfs');
          const spPdfsPath = `${spPath}/PDFs`;
          console.log(`  📎 Uploading ${meta.savedSourcePdfs.length} source PDFs to ${spPdfsPath}/`);
          for (const srcName of meta.savedSourcePdfs) {
            const srcFile = path.join(srcDir, srcName);
            if (fs.existsSync(srcFile)) {
              await uploadToSharePoint(spPdfsPath, srcName, fs.readFileSync(srcFile));
              console.log(`    ✅ PDFs/${srcName}`);
            }
          }
        }
        job.sharePointPath = spPath;
      } catch (err) {
        console.error(`  ❌ SharePoint upload failed:`, err.message);
        job.sharePointError = err.message;
      }
    }

    // Save job
    const jobs = loadJSON(JOBS_FILE, []); jobs.unshift(job); saveJSON(JOBS_FILE, jobs);

    // Clean up preview directory
    fs.rmSync(previewDir, { recursive: true, force: true });

    console.log(`✅ Pipeline send complete for "${kundeName}" — Job ${job.id}`);
    res.json({ ok: true, dryRun: !!job.dryRun, job, entries: meta.entries, generatedFile: meta.attachmentName, sharePointPath: spPath });
  } catch (e) {
    console.error('❌ Pipeline send error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// --- Pipeline (legacy): Extract PDFs → Generate DOCX → Convert PDF → Send Email → Upload SharePoint ---
app.post('/api/pipeline/run', upload.array('pdfs', 20), async (req, res) => {
  try {
    const { to, subject, body, kundeRaw, kundeName, kundeId, kundeTyp, anrede, empfaengerName, signaturPerson, emailTyp, gewerk, region } = req.body;
    if (!to || !subject || !body) return res.status(400).json({ error: 'to, subject, body required' });
    if (!req.files?.length) return res.status(400).json({ error: 'No PDFs uploaded' });
    if (!fs.existsSync(VORLAGE_FILE)) return res.status(400).json({ error: 'Keine Vorlage hochgeladen. Bitte zuerst eine DOCX-Vorlage hochladen.' });

    const pdfNames = req.files.map((f) => f.originalname);
    console.log(`\n🚀 Pipeline started for "${kundeName}" (${req.files.length} PDFs)`);

    // 1. Extract data from PDFs
    console.log('  📄 Extracting PDF data...');
    const entries = [];
    for (let i = 0; i < req.files.length; i++) {
      console.log(`    → ${pdfNames[i]}...`);
      try {
        const data = await extractFromPdfBuffer(req.files[i].buffer);
        entries.push(data);
        console.log(`    ✅ ID: ${data.dtad_id} | ${data.titel?.substring(0, 50)}`);
      } catch (err) {
        console.error(`    ❌ Failed: ${pdfNames[i]}:`, err.message);
        entries.push({ titel: pdfNames[i], dtad_id: '—', abgabetermin: '—', ausfuehrungsort: '—', beginn: '—', ende: '—', leistung: '—' });
      }
    }

    // 2. Generate DOCX from Vorlage
    console.log('  📝 Generating DOCX...');
    const baseName = `Recherche_${(kundeName || 'Kunde').replace(/\s+/g, '_')}`;
    const docxFilename = `${baseName}.docx`;
    const pdfFilename = `${baseName}.pdf`;
    const docxBuffer = await generateDocxBuffer(VORLAGE_FILE, entries, gewerk, region);

    // 3. Convert DOCX → PDF
    console.log('  📄 Converting DOCX → PDF...');
    const pdfOutputBuffer = await convertDocxToPdf(docxBuffer, docxFilename);
    if (pdfOutputBuffer) console.log(`  ✅ PDF generated: ${pdfFilename}`);
    else console.log('  ⚠️  PDF conversion failed, will send DOCX instead');

    // 4. Send email with generated document
    const attachmentBuffer = pdfOutputBuffer || docxBuffer;
    const attachmentName = pdfOutputBuffer ? pdfFilename : docxFilename;
    const attachmentType = pdfOutputBuffer ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    const bcc = process.env.PIPEDRIVE_SMART_BCC;

    const job = {
      id: Date.now().toString(36) + Math.random().toString(36).substr(2, 4),
      to, subject, kundeRaw: kundeRaw || '', kundeName: kundeName || '', kundeId: kundeId || '',
      anrede: anrede || 'Herr', empfaengerName: empfaengerName || '',
      signaturPerson: signaturPerson || '', emailTyp: emailTyp || 'regular',
      gewerk: gewerk || '', region: region || '',
      pdfNames, entries,
      generatedFile: attachmentName,
      sentAt: new Date().toISOString(),
      reminders: [],
    };

    if (process.env.DRY_RUN === 'true') {
      console.log(`  🔸 DRY RUN — Would send: ${to} | ${subject} | Attachment: ${attachmentName}`);
      job.dryRun = true;
    } else {
      console.log(`  ✉️  Sending email to ${to}...`);
      const info = await transporter.sendMail({
        from: process.env.EMAIL_FROM, to, bcc, subject, text: body,
        attachments: [{ filename: attachmentName, content: attachmentBuffer, contentType: attachmentType }],
      });
      console.log(`  ✅ Sent: ${info.messageId} → ${to}`);
      job.messageId = info.messageId;
    }

    // 5. Save attachment locally for preview & reminder re-attach
    const jobAttDir = path.join(ATTACHMENTS_DIR, job.id);
    fs.mkdirSync(jobAttDir, { recursive: true });
    fs.writeFileSync(path.join(jobAttDir, attachmentName), attachmentBuffer);
    job.savedAttachments = [attachmentName];

    // 6. Save email+name to cache
    if (kundeRaw && to) {
      const c = loadJSON(EMAIL_CACHE_FILE, {});
      const prev = typeof c[kundeRaw] === 'string' ? { email: c[kundeRaw], name: '' } : (c[kundeRaw] || { email: '', name: '' });
      prev.email = to;
      if (empfaengerName) prev.name = empfaengerName;
      c[kundeRaw] = prev;
      saveJSON(EMAIL_CACHE_FILE, c);
    }

    // 7. Upload DOCX + PDF to SharePoint date folder
    const today = new Date();
    const yy = String(today.getFullYear()).slice(-2);
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const dateFolder = `${yy}${mm}${dd}`;
    const spDir = kundeTyp === 'Interessenten' ? 'Interessenten' : 'Aktive_Kunden';
    const spPath = `${spDir}/${kundeRaw}/${dateFolder}`;

    if (process.env.DRY_RUN === 'true') {
      console.log(`  🔸 DRY RUN — Would upload to SharePoint: ${spPath}/${docxFilename} + ${pdfFilename}`);
      job.sharePointPath = spPath;
    } else {
      try {
        console.log(`  📁 Uploading to SharePoint: ${spPath}/`);
        await uploadToSharePoint(spPath, docxFilename, docxBuffer);
        console.log(`    ✅ ${docxFilename}`);
        if (pdfOutputBuffer) {
          await uploadToSharePoint(spPath, pdfFilename, pdfOutputBuffer);
          console.log(`    ✅ ${pdfFilename}`);
        }
        job.sharePointPath = spPath;
      } catch (err) {
        console.error(`  ❌ SharePoint upload failed:`, err.message);
        job.sharePointError = err.message;
      }
    }

    // 7. Save job
    const jobs = loadJSON(JOBS_FILE, []); jobs.unshift(job); saveJSON(JOBS_FILE, jobs);
    console.log(`✅ Pipeline complete for "${kundeName}" — Job ${job.id}`);

    res.json({ ok: true, dryRun: !!job.dryRun, job, entries, generatedFile: attachmentName, sharePointPath: spPath });
  } catch (e) {
    console.error('❌ Pipeline error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// --- Kunden from SharePoint ---
// SharePoint folder cache (60s TTL)
let _kundenCache = null;
let _kundenCacheTime = 0;
const KUNDEN_CACHE_TTL = 60 * 1000; // 60 seconds

app.get('/api/kunden', async (req, res) => {
  try {
    const forceRefresh = req.query.refresh === 'true';
    const now = Date.now();

    // Use cache if fresh and not forced
    if (!forceRefresh && _kundenCache && (now - _kundenCacheTime) < KUNDEN_CACHE_TTL) {
      // Re-merge with latest email cache
      const ec = loadJSON(EMAIL_CACHE_FILE, {});
      const resolve = (r) => {
        const entry = ec[r];
        if (!entry) return { email: '', contactName: '' };
        if (typeof entry === 'string') return { email: entry, contactName: '' };
        return { email: entry.email || '', contactName: entry.name || '' };
      };
      return res.json({
        aktive: _kundenCache.aktive.map((r) => ({ ...parseFolderName(r), typ: 'Aktive Kunden', ...resolve(r) })),
        interessenten: _kundenCache.interessenten.map((r) => ({ ...parseFolderName(r), typ: 'Interessenten', ...resolve(r) })),
      });
    }

    const [aRaw, iRaw] = await Promise.all([
      listSharePointFolders('Aktive_Kunden').catch(() => []),
      listSharePointFolders('Interessenten').catch(() => []),
    ]);

    // Update cache
    _kundenCache = { aktive: aRaw, interessenten: iRaw };
    _kundenCacheTime = now;

    const ec = loadJSON(EMAIL_CACHE_FILE, {});
    const resolve = (r) => {
      const entry = ec[r];
      if (!entry) return { email: '', contactName: '' };
      if (typeof entry === 'string') return { email: entry, contactName: '' };
      return { email: entry.email || '', contactName: entry.name || '' };
    };
    res.json({
      aktive: aRaw.map((r) => ({ ...parseFolderName(r), typ: 'Aktive Kunden', ...resolve(r) })),
      interessenten: iRaw.map((r) => ({ ...parseFolderName(r), typ: 'Interessenten', ...resolve(r) })),
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/kunden/email', (req, res) => {
  const { kundeRaw, email, name } = req.body;
  if (!kundeRaw) return res.status(400).json({ error: 'missing fields' });
  const c = loadJSON(EMAIL_CACHE_FILE, {});
  const prev = typeof c[kundeRaw] === 'string' ? { email: c[kundeRaw], name: '' } : (c[kundeRaw] || { email: '', name: '' });
  if (email !== undefined) prev.email = email;
  if (name !== undefined) prev.name = name;
  c[kundeRaw] = prev;
  saveJSON(EMAIL_CACHE_FILE, c);
  res.json({ ok: true });
});

// --- Pipedrive Contact Lookup ---
app.get('/api/pipedrive/lookup', async (req, res) => {
  const { name } = req.query;
  if (!name) return res.json({ email: '', contactName: '', source: 'none' });

  const apiToken = process.env.PIPEDRIVE_API_TOKEN;
  const domain = process.env.PIPEDRIVE_COMPANY_DOMAIN;
  if (!apiToken || !domain) {
    return res.json({ email: '', contactName: '', source: 'none' });
  }

  try {
    console.log(`  🔍 Pipedrive lookup: "${name}"`);

    // Search organizations by company name
    const searchUrl = `https://${domain}.pipedrive.com/api/v1/organizations/search?term=${encodeURIComponent(name)}&limit=5&api_token=${apiToken}`;
    const searchRes = await fetch(searchUrl);
    if (!searchRes.ok) throw new Error(`Pipedrive search ${searchRes.status}`);
    const searchData = await searchRes.json();

    const items = searchData?.data?.items;
    if (!items || items.length === 0) {
      console.log(`  ℹ️ Pipedrive: no org found for "${name}"`);
      return res.json({ email: '', contactName: '', source: 'pipedrive_no_match' });
    }

    const orgId = items[0].item.id;
    const orgName = items[0].item.name;
    console.log(`  ✅ Pipedrive org: "${orgName}" (ID: ${orgId})`);

    // Get persons linked to this organization
    const personsUrl = `https://${domain}.pipedrive.com/api/v1/organizations/${orgId}/persons?limit=10&api_token=${apiToken}`;
    const personsRes = await fetch(personsUrl);
    if (!personsRes.ok) throw new Error(`Pipedrive persons ${personsRes.status}`);
    const personsData = await personsRes.json();

    let persons = personsData?.data;

    // Fallback: if no persons linked to org, search persons by org name
    if (!persons || persons.length === 0) {
      console.log(`  ℹ️ Pipedrive: no linked persons, searching by org name...`);
      const psUrl = `https://${domain}.pipedrive.com/api/v1/persons/search?term=${encodeURIComponent(orgName)}&limit=5&api_token=${apiToken}`;
      const psRes = await fetch(psUrl);
      if (psRes.ok) {
        const psData = await psRes.json();
        const psItems = psData?.data?.items;
        if (psItems?.length) {
          persons = [];
          for (const item of psItems) {
            const pRes = await fetch(`https://${domain}.pipedrive.com/api/v1/persons/${item.item.id}?api_token=${apiToken}`);
            if (pRes.ok) { const pd = await pRes.json(); if (pd.data) persons.push(pd.data); }
          }
          console.log(`  ✅ Pipedrive: found ${persons.length} person(s) via search`);
        }
      }
    }

    if (!persons || persons.length === 0) {
      console.log(`  ℹ️ Pipedrive: no persons found for "${orgName}"`);
      return res.json({ email: '', extraEmails: '', contactName: '', source: 'pipedrive_no_persons' });
    }

    // Collect all emails from all persons, use first person's last name
    const allEmails = [];
    let contactLastName = '';
    for (const person of persons) {
      if (!contactLastName) contactLastName = person.last_name || person.name?.split(' ').pop() || '';
      const emails = person.email || [];
      // Primary email first
      const primary = emails.find(e => e.primary);
      if (primary?.value && !allEmails.includes(primary.value)) allEmails.push(primary.value);
      for (const em of emails) {
        if (em.value && !allEmails.includes(em.value)) allEmails.push(em.value);
      }
    }

    if (allEmails.length === 0) {
      return res.json({ email: '', extraEmails: '', contactName: contactLastName, source: 'pipedrive_no_email' });
    }

    const primaryEmail = allEmails[0];
    const extraEmails = allEmails.slice(1).join(', ');
    console.log(`  ✅ Pipedrive contact: ${contactLastName} — ${allEmails.length} email(s): ${allEmails.join(', ')}`);
    return res.json({ email: primaryEmail, extraEmails, contactName: contactLastName, source: 'pipedrive' });

  } catch (err) {
    console.error(`  ❌ Pipedrive lookup failed:`, err.message);
    return res.json({ email: '', contactName: '', source: 'error' });
  }
});

// ========================
// PIPEDRIVE REPLY TRACKING
// ========================
async function checkPipedriveReply(job) {
  const apiToken = process.env.PIPEDRIVE_API_TOKEN;
  const domain = process.env.PIPEDRIVE_COMPANY_DOMAIN;
  if (!apiToken || !domain) return { replied: false, error: 'no_config' };

  try {
    // Fast path: if we already have the thread ID, just check message_count
    if (job.pipedriveThreadId) {
      const threadUrl = `https://${domain}.pipedrive.com/api/v1/mailbox/mailThreads/${job.pipedriveThreadId}?api_token=${apiToken}`;
      const res = await fetch(threadUrl);
      if (res.ok) {
        const data = await res.json();
        if (data.data && data.data.message_count > 1) {
          return { replied: true, threadId: job.pipedriveThreadId };
        }
        return { replied: false, threadId: job.pipedriveThreadId };
      }
      // Thread ID might be stale, fall through to full lookup
    }

    // Step 1: Find person by recipient email
    let personId = job.pipedrivePersonId || null;
    if (!personId) {
      const searchUrl = `https://${domain}.pipedrive.com/api/v1/persons/search?term=${encodeURIComponent(job.to)}&limit=5&api_token=${apiToken}`;
      const searchRes = await fetch(searchUrl);
      if (!searchRes.ok) throw new Error(`Person search failed: ${searchRes.status}`);
      const searchData = await searchRes.json();
      const items = searchData?.data?.items;
      if (items?.length) {
        personId = items[0].item.id;
      }
    }

    if (!personId) return { replied: false, error: 'person_not_found' };

    // Step 2: Get person's mail messages
    const mailUrl = `https://${domain}.pipedrive.com/api/v1/persons/${personId}/mailMessages?limit=50&api_token=${apiToken}`;
    const mailRes = await fetch(mailUrl);
    if (!mailRes.ok) throw new Error(`Mail fetch failed: ${mailRes.status}`);
    const mailData = await mailRes.json();

    // Step 3: Find matching thread by subject
    const jobSubject = (job.subject || '').toLowerCase().replace(/^(re:|aw:|fwd?:)\s*/gi, '').trim();
    for (const item of (mailData.data || [])) {
      const threadSubject = (item.subject || '').toLowerCase().replace(/^(re:|aw:|fwd?:)\s*/gi, '').trim();
      if (threadSubject.includes(jobSubject) || jobSubject.includes(threadSubject)) {
        const threadId = item.mail_thread_id || item.id;
        if (item.message_count > 1) {
          return { replied: true, threadId, personId };
        }
        return { replied: false, threadId, personId };
      }
    }

    return { replied: false, personId };
  } catch (err) {
    console.error(`  ❌ Reply check failed for job ${job.id}:`, err.message);
    return { replied: false, error: err.message };
  }
}

// Background reply checker — runs every 10 minutes
let replyCheckRunning = false;
async function backgroundReplyCheck() {
  if (replyCheckRunning) return;
  const apiToken = process.env.PIPEDRIVE_API_TOKEN;
  const domain = process.env.PIPEDRIVE_COMPANY_DOMAIN;
  if (!apiToken || !domain) return;

  replyCheckRunning = true;
  try {
    const jobs = loadJSON(JOBS_FILE, []);
    const now = Date.now();
    const sevenDaysMs = 7 * 24 * 60 * 60 * 1000;
    const tenMinMs = 10 * 60 * 1000;

    const toCheck = jobs.filter(j =>
      j.replyStatus !== 'replied' &&
      !j.dryRun &&
      j.messageId &&
      (now - new Date(j.sentAt).getTime()) < sevenDaysMs &&
      (!j.lastReplyCheck || (now - new Date(j.lastReplyCheck).getTime()) > tenMinMs)
    );

    if (toCheck.length === 0) { replyCheckRunning = false; return; }
    console.log(`\n🔍 [ReplyCheck] Checking ${toCheck.length} job(s)...`);

    let changed = false;
    for (const job of toCheck) {
      const result = await checkPipedriveReply(job);
      job.lastReplyCheck = new Date().toISOString();

      if (result.replied && job.replyStatus !== 'replied') {
        job.replyStatus = 'replied';
        job.replyDetectedAt = new Date().toISOString();
        console.log(`  ✅ [ReplyCheck] Reply detected: ${job.kundeName} (${job.id})`);
        changed = true;
      }
      if (result.threadId) job.pipedriveThreadId = result.threadId;
      if (result.personId) job.pipedrivePersonId = result.personId;
      if (result.error) job.replyCheckError = result.error;

      // Throttle: 200ms between checks
      await new Promise(r => setTimeout(r, 200));
    }

    if (changed || toCheck.length > 0) saveJSON(JOBS_FILE, jobs);
    console.log(`  🔍 [ReplyCheck] Done. ${changed ? 'Updates saved.' : 'No new replies.'}`);
  } catch (err) {
    console.error('❌ [ReplyCheck] Error:', err.message);
  } finally {
    replyCheckRunning = false;
  }
}

// Run every 10 minutes, first check 30s after startup
setInterval(backgroundReplyCheck, 10 * 60 * 1000);
setTimeout(backgroundReplyCheck, 30000);

// --- Jobs ---
app.get('/api/jobs', (req, res) => res.json(loadJSON(JOBS_FILE, [])));

// --- Reply Status (lightweight read from jobs.json) ---
app.get('/api/jobs/reply-status', (req, res) => {
  const jobs = loadJSON(JOBS_FILE, []);
  const statuses = jobs.map(j => ({
    id: j.id,
    replyStatus: j.replyStatus || 'none',
    replyDetectedAt: j.replyDetectedAt || null,
    lastReplyCheck: j.lastReplyCheck || null,
  }));
  res.json(statuses);
});

// --- Manual reply check for a single job ---
app.post('/api/jobs/:jobId/check-reply', async (req, res) => {
  try {
    const jobs = loadJSON(JOBS_FILE, []);
    const job = jobs.find(j => j.id === req.params.jobId);
    if (!job) return res.status(404).json({ error: 'Job not found' });

    const result = await checkPipedriveReply(job);
    job.lastReplyCheck = new Date().toISOString();

    if (result.replied) {
      job.replyStatus = 'replied';
      job.replyDetectedAt = job.replyDetectedAt || new Date().toISOString();
    } else {
      job.replyStatus = 'none';
    }
    if (result.threadId) job.pipedriveThreadId = result.threadId;
    if (result.personId) job.pipedrivePersonId = result.personId;
    if (result.error) job.replyCheckError = result.error;

    saveJSON(JOBS_FILE, jobs);
    res.json({ replyStatus: job.replyStatus, replyDetectedAt: job.replyDetectedAt });
  } catch (e) {
    console.error('❌ Check reply:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// --- Manually mark a job as replied ---
app.post('/api/jobs/:jobId/mark-replied', (req, res) => {
  const jobs = loadJSON(JOBS_FILE, []);
  const job = jobs.find(j => j.id === req.params.jobId);
  if (!job) return res.status(404).json({ error: 'Job not found' });

  job.replyStatus = 'replied';
  job.replyDetectedAt = job.replyDetectedAt || new Date().toISOString();
  saveJSON(JOBS_FILE, jobs);
  res.json({ ok: true, replyStatus: 'replied' });
});

// --- Send Email (saves as job) ---
app.post('/api/send', upload.array('pdfs', 20), async (req, res) => {
  try {
    const { to, subject, body, kundeRaw, kundeName, kundeId, anrede, empfaengerName, signaturPerson, emailTyp, gewerk, region } = req.body;
    if (!to || !subject || !body) return res.status(400).json({ error: 'to, subject, body required' });

    const attachments = (req.files || []).map((f) => ({ filename: f.originalname, content: f.buffer, contentType: 'application/pdf' }));
    const bcc = process.env.PIPEDRIVE_SMART_BCC;

    const job = {
      id: Date.now().toString(36) + Math.random().toString(36).substr(2, 4),
      to, subject, kundeRaw: kundeRaw || '', kundeName: kundeName || '', kundeId: kundeId || '',
      anrede: anrede || 'Herr', empfaengerName: empfaengerName || '',
      signaturPerson: signaturPerson || '', emailTyp: emailTyp || 'regular',
      gewerk: gewerk || '', region: region || '',
      pdfNames: attachments.map((a) => a.filename),
      sentAt: new Date().toISOString(),
      reminders: [],
    };

    if (process.env.DRY_RUN === 'true') {
      console.log(`\n🔸 DRY RUN — Send: ${to} | ${subject} | ${attachments.length} PDFs`);
      job.dryRun = true;
    } else {
      const info = await transporter.sendMail({ from: process.env.EMAIL_FROM, to, bcc, subject, text: body, attachments });
      console.log(`✅ Sent: ${info.messageId} → ${to}`);
      job.messageId = info.messageId;
    }

    // Save attachments locally for preview & reminder re-attach
    if (req.files?.length) {
      const jobAttDir = path.join(ATTACHMENTS_DIR, job.id);
      fs.mkdirSync(jobAttDir, { recursive: true });
      req.files.forEach((f) => fs.writeFileSync(path.join(jobAttDir, f.originalname), f.buffer));
      job.savedAttachments = req.files.map((f) => f.originalname);
    }

    if (kundeRaw && to) {
      const c = loadJSON(EMAIL_CACHE_FILE, {});
      const prev = typeof c[kundeRaw] === 'string' ? { email: c[kundeRaw], name: '' } : (c[kundeRaw] || { email: '', name: '' });
      prev.email = to;
      if (empfaengerName) prev.name = empfaengerName;
      c[kundeRaw] = prev;
      saveJSON(EMAIL_CACHE_FILE, c);
    }

    const jobs = loadJSON(JOBS_FILE, []); jobs.unshift(job); saveJSON(JOBS_FILE, jobs);
    res.json({ ok: true, dryRun: !!job.dryRun, job });
  } catch (e) { console.error('❌ Send:', e.message); res.status(500).json({ error: e.message }); }
});

// --- Serve saved PDF attachment for preview ---
app.get('/api/jobs/:jobId/attachments/:filename', (req, res) => {
  const { jobId, filename } = req.params;
  const filePath = path.join(ATTACHMENTS_DIR, jobId, filename);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Attachment not found' });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
  fs.createReadStream(filePath).pipe(res);
});

// --- Send Reminder (HTML email with clickable link) ---
app.post('/api/send-reminder', upload.array('pdfs', 5), async (req, res) => {
  try {
    const { jobId, bodyText, bodyHtml, subject } = req.body;
    if (!jobId || !bodyText) return res.status(400).json({ error: 'jobId and bodyText required' });

    const jobs = loadJSON(JOBS_FILE, []);
    const job = jobs.find((j) => j.id === jobId);
    if (!job) return res.status(404).json({ error: 'Job not found' });

    // Start with any newly uploaded PDFs
    const attachments = (req.files || []).map((f) => ({ filename: f.originalname, content: f.buffer, contentType: 'application/pdf' }));

    // Re-attach original PDFs from the first email
    if (job.savedAttachments?.length) {
      const jobAttDir = path.join(ATTACHMENTS_DIR, job.id);
      for (const fname of job.savedAttachments) {
        const fpath = path.join(jobAttDir, fname);
        if (fs.existsSync(fpath)) {
          attachments.push({ filename: fname, content: fs.readFileSync(fpath), contentType: 'application/pdf' });
        }
      }
    }

    const bcc = process.env.PIPEDRIVE_SMART_BCC;
    const subj = subject || `Erinnerung – KALKU`;

    const reminder = { sentAt: new Date().toISOString(), subject: subj };

    if (process.env.DRY_RUN === 'true') {
      console.log(`\n🔸 DRY RUN — Reminder: ${job.to} | ${subj} | ${attachments.length} attachment(s)`);
      reminder.dryRun = true;
    } else {
      const info = await transporter.sendMail({
        from: process.env.EMAIL_FROM, to: job.to, bcc, subject: subj,
        text: bodyText, html: bodyHtml || undefined, attachments,
      });
      console.log(`✅ Reminder: ${info.messageId} → ${job.to}`);
      reminder.messageId = info.messageId;
    }

    job.reminders.push(reminder);
    saveJSON(JOBS_FILE, jobs);
    res.json({ ok: true, dryRun: !!reminder.dryRun, reminder });
  } catch (e) { console.error('❌ Reminder:', e.message); res.status(500).json({ error: e.message }); }
});

// Static
if (fs.existsSync('./dist')) { app.use(express.static('./dist')); app.get('*', (req, res) => res.sendFile(path.resolve('./dist/index.html'))); }

// ========================
// SSH TUNNEL FOR BACKEND API
// ========================
function setupSSHTunnel() {
  // Check if sshpass is available
  try {
    execSync('which sshpass', { stdio: 'ignore' });
  } catch (e) {
    console.log('⚠️  sshpass not installed, skipping SSH tunnel setup');
    return null;
  }

  const sshHost = '91.98.185.113';
  const sshUser = 'anjali';
  const sshPassword = 'U&!uwZ#FYv0gi8';
  const remotePort = 9090;
  const localPort = 9090;

  console.log(`🔐 Setting up SSH tunnel to ${sshUser}@${sshHost}:${remotePort}...`);

  // Use sshpass to provide password non-interactively
  const sshCmd = spawn('sshpass', [
    '-p', sshPassword,
    'ssh',
    '-o', 'StrictHostKeyChecking=no',
    '-o', 'ServerAliveInterval=60',
    '-L', `${localPort}:localhost:${remotePort}`,
    '-N',
    `${sshUser}@${sshHost}`
  ]);

  sshCmd.stdout.on('data', (data) => {
    console.log(`SSH: ${data.toString().trim()}`);
  });

  sshCmd.stderr.on('data', (data) => {
    const msg = data.toString().trim();
    if (msg.includes('Allocated port') || msg.includes('Entering interactive session')) {
      console.log(`✅ SSH tunnel established: localhost:${localPort} → ${sshHost}:${remotePort}`);
    } else if (!msg.includes('Warning')) {
      console.log(`SSH: ${msg}`);
    }
  });

  sshCmd.on('error', (err) => {
    console.error('❌ SSH tunnel error:', err.message);
    console.error('   Make sure sshpass is installed: brew install hudochenkov/sshpass/sshpass');
  });

  sshCmd.on('close', (code) => {
    console.log(`⚠️  SSH tunnel closed with code ${code}`);
    if (code !== 0) {
      console.log('   Retrying in 5 seconds...');
      setTimeout(setupSSHTunnel, 5000);
    }
  });

  return sshCmd;
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n🚀 KALKU Tender Tool v2 — http://localhost:${PORT}`);
  console.log(`   DRY_RUN: ${process.env.DRY_RUN === 'true' ? '✅ ON' : '❌ OFF'} | Pipedrive BCC: ${process.env.PIPEDRIVE_SMART_BCC}`);
  
  // Setup SSH tunnel for backend API
  setupSSHTunnel();
  
  console.log(`   Backend API: https://91.98.185.113:9090/users (via SSH tunnel)\n`);
});
