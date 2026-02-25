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
import { execSync } from 'child_process';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const pdfParse = require('pdf-parse');
import PizZip from 'pizzip';

const app = express();
app.use(cors());
app.use((req, res, next) => { console.log(`[${new Date().toISOString().slice(11,19)}] ${req.method} ${req.url}`); next(); });
app.use(express.json({ limit: '50mb' }));

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
  const url = fp
    ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(fp)}:/children?$top=999&$filter=folder ne null`
    : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=999&$filter=folder ne null`;
  return (await graphGet(url)).value?.map((i) => i.name) || [];
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

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// ========================
// VORLAGE FILE
// ========================
const VORLAGE_FILE = path.join(DATA_DIR, 'vorlage.docx');

// ========================
// PDF EXTRACTION: pdf-parse → regex → Claude validation
// ========================
function matchField(text, regex) {
  const m = text.match(regex);
  return m ? m[1].trim() : '—';
}

function extractFieldsFromText(text) {
  const titel = matchField(text, /Titel:\s*([^\n]+)/i);
  const dtad_id = matchField(text, /DTAD[- ]ID:\s*([^\n]+)/i);
  const abgabetermin = matchField(text, /Frist Angebotsabgabe:\s*([^\n]+)/i);

  let ausfuehrungsort = matchField(text, /Erfüllungsort[^:\n]*:\s*([^\n]+)/i);
  if (ausfuehrungsort === '—') ausfuehrungsort = matchField(text, /Region:\s*([^\n]+)/i);

  let beginn = '—', ende = '—';
  const datePatterns = [
    [/Datum des Beginns:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i, /Enddatum der Laufzeit:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i],
    [/\bBeginn:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i, /\bEnde:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i],
    [/Ausführungsbeginn:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i, /Ausführungsende:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i],
    [/Leistungsbeginn:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i, /Leistungsende:\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i],
    [/\bab\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i, /\bbis\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/i],
  ];

  for (const [bp, ep] of datePatterns) {
    if (beginn === '—') { const m = text.match(bp); if (m) beginn = m[1].trim(); }
    if (ende === '—') { const m = text.match(ep); if (m) ende = m[1].trim(); }
    if (beginn !== '—' && ende !== '—') break;
  }

  if (beginn === '—' || ende === '—') {
    const rangePatterns = [
      /Ausführungsfrist[^:]*:\s*Beginn:?\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})\s*[-–]?\s*Ende:?\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/is,
      /(\d{1,2}[./]\d{1,2}[./]\d{2,4})\s*[-–]\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4})/,
    ];
    for (const p of rangePatterns) {
      const m = text.match(p);
      if (m) { if (beginn === '—') beginn = m[1].trim(); if (ende === '—') ende = m[2].trim(); break; }
    }
  }

  let leistung = '—';
  for (const p of [/Art und Umfang der Leistung:?\s*([\s\S]{10,3000})/i, /Kurzbeschreibung:?\s*([\s\S]{10,2000})/i, /Beschreibung:?\s*([\s\S]{10,2000})/i]) {
    const m = text.match(p);
    if (m?.[1]) { leistung = m[1].trim().split(/\n\s*\n/)[0].substring(0, 2000); if (leistung.length > 10) break; }
  }

  return { titel, dtad_id, abgabetermin, ausfuehrungsort, beginn, ende, leistung };
}

async function validateWithAnthropic(fields, text) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.log('    ⚠️  No Anthropic API key, skipping AI validation');
    return fields;
  }

  try {
    console.log('    🤖 Calling Claude for field validation...');
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000); // 10 second timeout
    
    const prompt = `Du bist ein Datenextraktor für deutsche DTAD-Ausschreibungs-PDFs.

Hier ist der PDF-Text:
${text.slice(0, 8000)}

Extrahierte Felder: ${JSON.stringify(fields)}

Prüfe und korrigiere alle Felder. Wenn ein Feld "—" ist, versuche es aus dem Text zu extrahieren.
Gib NUR ein JSON-Objekt zurück: titel, dtad_id, abgabetermin, ausfuehrungsort, beginn, ende, leistung`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2024-10-22' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 500, messages: [{ role: 'user', content: [{ type: 'text', text: prompt }] }] }),
      signal: controller.signal,
    });
    clearTimeout(timeout);

    const data = await response.json();
    if (!response.ok || data.error) {
      console.log('    ⚠️  Claude API error:', data.error?.message || response.status);
      return fields;
    }
    const textOut = (data.content || []).map((c) => (c.type === 'text' ? c.text : '')).join('');
    console.log('    ✅ Claude validation complete');
    return { ...fields, ...JSON.parse(textOut.replace(/```json|```/g, '').trim()) };
  } catch (err) {
    console.log('    ⚠️  Claude validation skipped:', err.name === 'AbortError' ? 'timeout' : err.message);
    return fields;
  }
}

async function extractFromPdfBuffer(pdfBuffer) {
  console.log('    📄 Parsing PDF...');
  // Step 1: pdf-parse extracts text
  const result = await pdfParse(pdfBuffer);
  const text = result.text || '';
  console.log(`    ✅ PDF parsed (${text.length} chars)`);

  // Step 2: regex first pass
  const fieldsRaw = extractFieldsFromText(text);

  // Step 3: Skip Claude validation for speed - just use regex results
  // return validateWithAnthropic(fieldsRaw, text);
  return fieldsRaw;
}

// ========================
// DOCX GENERATION (ported from v2.1)
// ========================
function generateDocxBuffer(vorlageBuffer, entries, gewerk, region) {
  const zip = new PizZip(vorlageBuffer);
  let docXml = zip.file('word/document.xml').asText();

  const escapeXml = (t) => String(t || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const replaceText = (xml, s, r) => xml.split(s).join(escapeXml(r));

  if (gewerk) { docXml = replaceText(docXml, 'Gebäudereinigungsarbeiten', gewerk); docXml = replaceText(docXml, 'Gebäudereinigung', gewerk); }
  if (region) { docXml = replaceText(docXml, '59759 Arnsberg + 50 km', region); docXml = replaceText(docXml, '47051 Duisburg + 50km', region); }

  let idx;
  idx = 0; docXml = docXml.replace(/ID:[^<]*/g, () => `ID: ${escapeXml(entries[idx]?.dtad_id || '—')}` + (idx++, ''));
  idx = 0; docXml = docXml.replace(/Abgabetermin:[^<]*/g, () => `Abgabetermin: ${escapeXml(entries[idx]?.abgabetermin || '—')}` + (idx++, ''));
  idx = 0; docXml = docXml.replace(/Ausführungsort:[^<]*/g, () => `Ausführungsort: ${escapeXml(entries[idx]?.ausfuehrungsort || '—')}` + (idx++, ''));
  idx = 0; docXml = docXml.replace(/Ausführungsfrist:[^<]*Beginn:[^<]*Ende:[^<]*/g, () => {
    const e = entries[idx] || {}; idx++;
    return `Ausführungsfrist: Beginn: ${escapeXml(e.beginn || '—')} - Ende: ${escapeXml(e.ende || '—')}`;
  });
  idx = 0; docXml = docXml.replace(/Art und Umfang der Leistung:[^<]*/g, () => {
    const e = entries[idx] || {}; idx++;
    const leistung = (e.leistung || '—').replace(/\n/g, ' ');
    return `Art und Umfang der Leistung: ${escapeXml(leistung)}`;
  });

  zip.file('word/document.xml', docXml);
  return zip.generate({ type: 'nodebuffer', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

// ========================
// DOCX → PDF CONVERSION
// ========================
const LOCAL_TMP_DIR = path.join(DATA_DIR, '_tmp');
if (!fs.existsSync(LOCAL_TMP_DIR)) fs.mkdirSync(LOCAL_TMP_DIR, { recursive: true });

function convertDocxToPdf(docxBuffer, outputFilename) {
  const tmpDir = fs.mkdtempSync(path.join(LOCAL_TMP_DIR, 'tender-'));
  const docxPath = path.join(tmpDir, outputFilename);
  fs.writeFileSync(docxPath, docxBuffer);

  try {
    console.log(`  🔄 Converting with LibreOffice: ${docxPath}`);
    // Kill any hanging soffice processes first
    try { execSync('pkill -9 soffice', { timeout: 2000 }); } catch (e) { /* ignore */ }
    
    // Run conversion with increased timeout and explicit environment
    execSync(`/opt/homebrew/bin/soffice --headless --nofirststartwizard --nologo --norestore --convert-to pdf --outdir "${tmpDir}" "${docxPath}"`, { 
      timeout: 60000,
      stdio: 'pipe',
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
  const filePath = `${folderPath}/${fileName}`;
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(filePath)}:/content`;
  return graphPut(url, buffer);
}

// ========================
// ROUTES
// ========================

app.get('/api/health', (req, res) => res.json({ status: 'ok', dryRun: process.env.DRY_RUN === 'true', hasVorlage: fs.existsSync(VORLAGE_FILE) }));

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

// --- Pipeline Step 1: Generate PDF only (for preview before sending) ---
app.post('/api/pipeline/generate', upload.array('pdfs', 20), async (req, res) => {
  process.stdout.write('\n📥 /api/pipeline/generate called\n');
  try {
    const { kundeName, gewerk, region } = req.body;
    if (!req.files?.length) return res.status(400).json({ error: 'No PDFs uploaded' });
    if (!fs.existsSync(VORLAGE_FILE)) return res.status(400).json({ error: 'Keine Vorlage hochgeladen.' });

    const vorlageBuffer = fs.readFileSync(VORLAGE_FILE);
    const pdfNames = req.files.map((f) => f.originalname);
    console.log(`\n🔧 Generate started for "${kundeName}" (${req.files.length} PDFs)`);

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

    // 2. Generate DOCX
    console.log('  📝 Generating DOCX...');
    const baseName = `Recherche_${(kundeName || 'Kunde').replace(/\s+/g, '_')}`;
    const docxFilename = `${baseName}.docx`;
    const pdfFilename = `${baseName}.pdf`;
    const docxBuffer = generateDocxBuffer(vorlageBuffer, entries, gewerk, region);

    // 3. Convert DOCX → PDF
    console.log('  📄 Converting DOCX → PDF...');
    const pdfOutputBuffer = convertDocxToPdf(docxBuffer, docxFilename);
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
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Preview not found' });
  const ext = path.extname(filename).toLowerCase();
  const ct = ext === '.pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  res.setHeader('Content-Type', ct);
  res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
  fs.createReadStream(filePath).pipe(res);
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

    const vorlageBuffer = fs.readFileSync(VORLAGE_FILE);
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
    const docxBuffer = generateDocxBuffer(vorlageBuffer, entries, gewerk, region);

    // 3. Convert DOCX → PDF
    console.log('  📄 Converting DOCX → PDF...');
    const pdfOutputBuffer = convertDocxToPdf(docxBuffer, docxFilename);
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
app.get('/api/kunden', async (req, res) => {
  try {
    const [aRaw, iRaw] = await Promise.all([
      listSharePointFolders('Aktive_Kunden').catch(() => []),
      listSharePointFolders('Interessenten').catch(() => []),
    ]);
    const ec = loadJSON(EMAIL_CACHE_FILE, {});
    const resolve = (r) => {
      const entry = ec[r];
      if (!entry) return { email: '', contactName: '' };
      if (typeof entry === 'string') return { email: entry, contactName: '' }; // backward compat
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

// --- Jobs ---
app.get('/api/jobs', (req, res) => res.json(loadJSON(JOBS_FILE, [])));

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

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n🚀 KALKU Tender Tool v2 — http://localhost:${PORT}`);
  console.log(`   DRY_RUN: ${process.env.DRY_RUN === 'true' ? '✅ ON' : '❌ OFF'} | Pipedrive BCC: ${process.env.PIPEDRIVE_SMART_BCC}\n`);
});
