import { useState, useRef, useEffect, useMemo, useCallback } from "react";

// ========================
// API
// ========================
const api = {
  async getKunden() { const r = await fetch("/api/kunden"); if (!r.ok) throw new Error(await r.text()); return r.json(); },
  async saveEmail(kundeRaw, email, name) { return (await fetch("/api/kunden/email", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ kundeRaw, email, name }) })).json(); },
  async getJobs() { const r = await fetch("/api/jobs"); return r.json(); },
  async sendEmail(data) {
    const form = new FormData();
    Object.entries(data).forEach(([k, v]) => { if (k === "pdfs") v.forEach((f) => form.append("pdfs", f)); else if (v) form.append(k, v); });
    const r = await fetch("/api/send", { method: "POST", body: form });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async sendReminder(jobId, bodyText, bodyHtml, subject) {
    const r = await fetch("/api/send-reminder", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jobId, bodyText, bodyHtml, subject }),
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async health() { return (await fetch("/api/health")).json(); },
  async getVorlage() { return (await fetch("/api/vorlage")).json(); },
  async uploadVorlage(file) {
    const form = new FormData();
    form.append("vorlage", file);
    const r = await fetch("/api/vorlage", { method: "POST", body: form });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async runPipeline(data) {
    const form = new FormData();
    Object.entries(data).forEach(([k, v]) => { if (k === "pdfs") v.forEach((f) => form.append("pdfs", f)); else if (v) form.append(k, v); });
    const r = await fetch("/api/pipeline/run", { method: "POST", body: form });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async extractPdf(file) {
    // Extract single PDF - convert to base64 and send to /api/extract
    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
      reader.readAsDataURL(file);
    });
    
    const r = await fetch("/api/extract", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pdfBase64: base64 })
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async generatePipeline(data, onProgress) {
    // Process PDFs one at a time (avoids payload size issues)
    const entries = [];
    for (let i = 0; i < data.pdfs.length; i++) {
      const file = data.pdfs[i];
      if (onProgress) onProgress(i + 1, data.pdfs.length, file.name);
      try {
        const extracted = await this.extractPdf(file);
        entries.push({ ...extracted, _filename: file.name });
      } catch (err) {
        console.error(`Failed to extract ${file.name}:`, err);
        entries.push({ titel: file.name, dtad_id: '—', abgabetermin: '—', ausfuehrungsort: '—', beginn: '—', ende: '—', leistung: '—', _filename: file.name });
      }
    }
    
    // Now generate DOCX with extracted data
    const r = await fetch("/api/pipeline/generate-docx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ kundeName: data.kundeName, gewerk: data.gewerk, region: data.region, entries })
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async sendPipeline(data) {
    const r = await fetch("/api/pipeline/send", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
};

// ========================
// SIGNATURES
// ========================
const SIGNATUREN = {
  "Dennis Engel": { rolle: "Rechercheteam", full: "Dennis Engel\n- Rechercheteam -\nTel.: (+49) 01579-2600775\n______________________\nKALKU - Baukalkulationen\nInh. Alaatdin Coksari\nBerliner Promenade 15\nD-66111 Saarbrücken\nTel.: (+49) 0681-41096430\nInternet: https://kalku.de\nSteuer-Nr.: 040/221/35122\nUSt-IdNr.: DE334890692\n______________________" },
  "Julian Kallenborn": { rolle: "Kundenservice", full: "Julian Kallenborn\n- Kundenservice -\nTel.: (+49) 0157-92600772\n______________________\nKALKU - Baukalkulationen\nInh. Alaatdin Coksari\nBerliner Promenade 15\nD-66111 Saarbrücken\nTel.: (+49) 0681-41096430\nInternet: https://kalku.de\nSteuer-Nr.: 040/221/35122\nUSt-IdNr.: DE334890692\n______________________" },
  "Anna Buxbaum": { rolle: "Kundenservice", full: "Anna Buxbaum\n- Kundenservice -\nTel.: (+49) 01579-2600703\n______________________\nKALKU - Baukalkulationen\nInh. Alaatdin Coksari\nBerliner Promenade 15\nD-66111 Saarbrücken\nTel.: (+49) 0681-41096430\nInternet: https://kalku.de\nSteuer-Nr.: 040/221/35122\nUSt-IdNr.: DE334890692\n______________________" },
};

// ========================
// EMAIL TEMPLATES
// ========================
const EMAIL_TEMPLATES = {
  regular: (anrede, name, sig) => `Hallo ${anrede} ${name},\n\nim Anhang finden Sie unsere aktuelle für Sie erstellte Recherche.\n\nWenn eine Ausschreibung für Sie interessant klingt, teilen Sie uns bitte die zugehörige ID mit.\n\nBitte geben Sie uns bis morgen Rückmeldung, für welche Ausschreibungen Sie sich interessieren.\n\nVielen Dank.\n\nBei Rückfragen stehen wir Ihnen zur Verfügung.\n\nMit freundlichen Grüßen\n\n${sig}`,
  erste: (anrede, name, sig) => `Hallo ${anrede} ${name},\n\nvielen Dank für Ihr Interesse an unserer Kalkulationsdienstleistung.\n\nIm Anhang finden Sie unsere erste für Sie erstellte Recherche.\n\nUnsere Recherche beinhaltet immer die wichtigsten Angaben und eine Zusammenfassung der Leistungen.\n\nWenn eine Ausschreibung für Sie interessant klingt, teilen Sie uns bitte die zugehörige ID mit.\n\nAls nächstes laden wir die jeweiligen Vergabeunterlagen runter und senden Ihnen die Dokumente zur Durchsicht zu.\n\nDanach können weitere Details besprochen werden und Sie können sich abstimmen, welche Ausschreibungen wir für Sie kalkulieren sollen.\n\nSind Arbeiten dabei, die Sie gar nicht ausführen, weisen Sie uns bitte darauf hin, sodass wir unsere nächsten Recherchen besser anpassen können.\n\nBitte geben Sie uns bis morgen Rückmeldung für welche Ausschreibungen Sie sich interessieren.\n\nVielen Dank.\n\nBei Rückfragen sind wir jederzeit gerne für Sie da.\n\nMit freundlichen Grüßen\n\n${sig}`,
};

// Reminder template — plain text version
function buildReminderText(anrede, nachname, signaturPerson) {
  const sig = SIGNATUREN[signaturPerson]?.full || signaturPerson;
  return `Sehr geehrte${anrede === "Frau" ? "" : "r"} ${anrede} ${nachname},\n\nSie hatten unsere Werbung in den Sozialen Medien gesehen und Ihre Kontaktdaten hinterlassen. Wir unterstützen Unternehmen dabei an öffentlichen Ausschreibungen teilzunehmen.\n\nLeider war es mir nicht möglich, ein Qualifizierungstelefonat mit Ihnen durchzuführen. Anbei finden Sie nochmal unsere Info PDF mit unseren Dienstleistungen.\n\nSollten Sie weiterhin an dieser Dienstleistung interessiert sein, buchen Sie bitte einen Abstimmungstermin über unseren Link:\nhttps://kalku.de/abstimmungstermin\n\nWir freuen uns von Ihnen zu hören.\n\nFreundliche Grüße\n\n${sig}`;
}

// Reminder template — HTML version with clickable link
function buildReminderHtml(anrede, nachname, signaturPerson) {
  const sig = SIGNATUREN[signaturPerson]?.full || signaturPerson;
  const sigHtml = sig.replace(/\n/g, "<br>").replace(/https:\/\/kalku\.de/g, '<a href="https://kalku.de" style="color:#2563eb">https://kalku.de</a>');
  return `<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.6">
<p>Sehr geehrte${anrede === "Frau" ? "" : "r"} ${anrede} ${nachname},</p>
<p>Sie hatten unsere Werbung in den Sozialen Medien gesehen und Ihre Kontaktdaten hinterlassen. Wir unterstützen Unternehmen dabei an öffentlichen Ausschreibungen teilzunehmen.</p>
<p>Leider war es mir nicht möglich, ein Qualifizierungstelefonat mit Ihnen durchzuführen. Anbei finden Sie nochmal unsere Info PDF mit unseren Dienstleistungen.</p>
<p>Sollten Sie weiterhin an dieser Dienstleistung interessiert sein, buchen Sie bitte einen <a href="https://kalku.de/abstimmungstermin" style="color:#2563eb;text-decoration:underline;font-weight:600">Abstimmungstermin</a> über unseren Link.</p>
<p>Wir freuen uns von Ihnen zu hören.</p>
<p>Freundliche Grüße</p>
<p style="margin-top:16px;color:#555">${sigHtml}</p>
</div>`;
}

// ========================
// ICONS
// ========================
const I = {
  Send: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>,
  File: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>,
  Users: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>,
  Settings: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>,
  Briefcase: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>,
  Check: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  X: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  Mail: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>,
  Search: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>,
  Upload: () => <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
  ChevronDown: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 12 15 18 9"/></svg>,
  Eye: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>,
  Edit: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>,
  Trash: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>,
  Pipedrive: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg>,
  Attachment: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/></svg>,
  Refresh: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>,
  Bell: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>,
  Clock: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>,
};

// ========================
// KUNDEN DROPDOWN
// ========================
function KundenDropdown({ kunden, value, onChange, loading, onRefresh }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("alle");
  const ref = useRef(null);
  const inputRef = useRef(null);
  useEffect(() => { const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }; document.addEventListener("mousedown", h); return () => document.removeEventListener("mousedown", h); }, []);
  useEffect(() => { if (open && inputRef.current) inputRef.current.focus(); }, [open]);

  const filtered = useMemo(() => {
    let list = [];
    if (filter === "alle" || filter === "aktive") list = list.concat(kunden.filter((k) => k.typ === "Aktive Kunden"));
    if (filter === "alle" || filter === "interessenten") list = list.concat(kunden.filter((k) => k.typ === "Interessenten"));
    if (search.trim()) { const s = search.toLowerCase(); list = list.filter((k) => k.searchText.includes(s) || k.raw.toLowerCase().includes(s)); }
    return list;
  }, [kunden, filter, search]);

  const grouped = useMemo(() => { const g = {}; filtered.forEach((k) => { if (!g[k.typ]) g[k.typ] = []; g[k.typ].push(k); }); return g; }, [filtered]);
  const sel = value ? kunden.find((k) => k.raw === value) : null;
  const isAktiv = sel?.typ === "Aktive Kunden";

  return (
    <div ref={ref} style={{ position: "relative", width: "100%" }}>
      <div onClick={() => setOpen(!open)} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 14px", background: "#fff", border: open ? "2px solid #2563eb" : "1.5px solid #d1d5db", borderRadius: 10, cursor: "pointer", fontSize: 14, color: sel ? "#1e293b" : "#94a3b8", minHeight: 44 }}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "flex", alignItems: "center", gap: 8 }}>
          {loading ? "SharePoint laden..." : sel ? (<><span style={{ width: 8, height: 8, borderRadius: "50%", background: isAktiv ? "#16a34a" : "#f59e0b", flexShrink: 0 }} /><span style={{ fontWeight: 600 }}>{sel.id}</span><span style={{ color: "#475569" }}>{sel.name}</span></>) : "— Kunde wählen —"}
        </span>
        <I.ChevronDown />
      </div>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, background: "#fff", border: "1.5px solid #d1d5db", borderRadius: 12, boxShadow: "0 12px 40px rgba(0,0,0,0.12)", zIndex: 100, overflow: "hidden", animation: "dropIn .15s ease-out" }}>
          <div style={{ padding: "10px 12px 6px", borderBottom: "1px solid #f1f5f9" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#f8fafc", borderRadius: 8, padding: "8px 12px", border: "1px solid #e2e8f0" }}>
              <I.Search /><input ref={inputRef} value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Nr. oder Name..." style={{ border: "none", outline: "none", background: "transparent", fontSize: 14, width: "100%", color: "#1e293b" }} />
              {search && <span style={{ cursor: "pointer", color: "#94a3b8" }} onClick={() => setSearch("")}><I.X /></span>}
            </div>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 12px", borderBottom: "1px solid #f1f5f9", alignItems: "center" }}>
            <div style={{ display: "flex" }}>{[{ key: "alle", label: "Alle" }, { key: "aktive", label: "Aktive" }, { key: "interessenten", label: "Interessenten" }].map((t) => (<button key={t.key} onClick={(e) => { e.stopPropagation(); setFilter(t.key); }} style={{ padding: "6px 12px", fontSize: 12, fontWeight: filter === t.key ? 700 : 500, color: filter === t.key ? "#2563eb" : "#64748b", background: filter === t.key ? "#eff6ff" : "transparent", border: "none", borderRadius: 6, cursor: "pointer" }}>{t.label}</button>))}</div>
            <button onClick={(e) => { e.stopPropagation(); onRefresh(); }} style={{ border: "none", background: "none", cursor: "pointer", color: "#94a3b8", padding: 4 }}><I.Refresh /></button>
          </div>
          <div style={{ maxHeight: 300, overflowY: "auto" }}>
            {!loading && Object.keys(grouped).length === 0 && <div style={{ padding: 20, textAlign: "center", color: "#94a3b8", fontSize: 13 }}>Keine Kunden</div>}
            {Object.entries(grouped).map(([typ, items]) => (<div key={typ}>
              <div style={{ padding: "8px 16px 4px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: typ === "Aktive Kunden" ? "#16a34a" : "#d97706", background: typ === "Aktive Kunden" ? "#f0fdf4" : "#fffbeb" }}><span style={{ width: 7, height: 7, borderRadius: "50%", background: typ === "Aktive Kunden" ? "#16a34a" : "#f59e0b", display: "inline-block", marginRight: 6 }} />{typ} ({items.length})</div>
              {items.map((k) => (<div key={k.raw} onClick={() => { onChange(k.raw); setOpen(false); setSearch(""); }} style={{ padding: "10px 16px", cursor: "pointer", background: value === k.raw ? "#eff6ff" : "transparent", borderLeft: value === k.raw ? "3px solid #2563eb" : "3px solid transparent" }} onMouseEnter={(e) => { if (value !== k.raw) e.currentTarget.style.background = "#f8fafc"; }} onMouseLeave={(e) => { if (value !== k.raw) e.currentTarget.style.background = "transparent"; }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontWeight: 700, color: "#2563eb", minWidth: 30 }}>{k.id}</span><span style={{ fontWeight: 600, fontSize: 14 }}>{k.name}</span></div>
                <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 2, paddingLeft: 38 }}>{k.email ? <span style={{ color: "#16a34a" }}>📧 {k.email}</span> : "Keine E-Mail"}</div>
              </div>))}
            </div>))}
          </div>
        </div>
      )}
    </div>
  );
}

// ========================
// CONFIRM MODAL
// ========================
function ConfirmModal({ open, data, onConfirm, onCancel, sending, title, pdfPreviewUrl, pdfPreviewName, sendLabel }) {
  const [showPdf, setShowPdf] = useState(false);
  if (!open) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", backdropFilter: "blur(4px)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", animation: "fadeIn .2s" }}>
      <div style={{ background: "#fff", borderRadius: 16, maxWidth: showPdf ? 1000 : 600, width: "92%", maxHeight: "90vh", overflow: "hidden", boxShadow: "0 24px 80px rgba(0,0,0,0.2)", animation: "modalIn .25s ease-out", display: "flex", flexDirection: "column", transition: "max-width .3s ease" }}>
        <div style={{ padding: "18px 24px", borderBottom: "1px solid #f1f5f9", display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 40, height: 40, borderRadius: 10, background: "linear-gradient(135deg, #fef3c7, #fde68a)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>✉️</div>
          <div><h3 style={{ margin: 0, fontSize: 17, fontWeight: 700 }}>{title || "E-Mail senden bestätigen"}</h3><p style={{ margin: 0, fontSize: 13, color: "#64748b" }}>Bitte überprüfen</p></div>
          <button onClick={onCancel} style={{ marginLeft: "auto", background: "none", border: "none", cursor: "pointer", color: "#94a3b8" }}><I.X /></button>
        </div>
        <div style={{ padding: "16px 24px", overflow: "auto", flex: 1 }}>
          <div style={{ background: "#f8fafc", borderRadius: 10, padding: 14, marginBottom: 14, display: "grid", gap: 8, fontSize: 14 }}>
            <div><span style={{ color: "#64748b", fontWeight: 500 }}>An: </span><strong>{data?.to}</strong></div>
            <div><span style={{ color: "#64748b", fontWeight: 500 }}>Betreff: </span><strong>{data?.subject}</strong></div>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ color: "#64748b", fontWeight: 500 }}>BCC: </span><span style={{ color: "#16a34a", fontWeight: 600 }}><I.Pipedrive /> Pipedrive</span><span style={{ background: "#dcfce7", color: "#16a34a", fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 10 }}>AUTO</span></div>
          </div>
          {/* PDF Attachment Preview */}
          {pdfPreviewUrl && (
            <div style={{ border: "1px solid #e2e8f0", borderRadius: 10, overflow: "hidden", marginBottom: 14 }}>
              <div style={{ padding: "8px 14px", background: "#eff6ff", borderBottom: "1px solid #e2e8f0", fontSize: 12, fontWeight: 600, color: "#2563eb", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span style={{ display: "flex", alignItems: "center", gap: 6 }}><I.Attachment /> Anhang: {pdfPreviewName}</span>
                <button onClick={() => setShowPdf(!showPdf)} style={{ border: "none", background: showPdf ? "#2563eb" : "#dbeafe", color: showPdf ? "#fff" : "#2563eb", padding: "4px 12px", borderRadius: 6, fontSize: 11, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}>
                  <I.Eye /> {showPdf ? "PDF ausblenden" : "PDF anzeigen"}
                </button>
              </div>
              {showPdf && (
                <iframe src={pdfPreviewUrl} style={{ width: "100%", height: 400, border: "none", background: "#f1f5f9" }} title="PDF Vorschau" />
              )}
            </div>
          )}
          <div style={{ border: "1px solid #e2e8f0", borderRadius: 10, overflow: "hidden" }}>
            <div style={{ padding: "8px 14px", background: "#f8fafc", borderBottom: "1px solid #e2e8f0", fontSize: 12, fontWeight: 600, color: "#64748b" }}>Inhalt</div>
            <div style={{ padding: 14, maxHeight: 220, overflow: "auto", fontSize: 13, color: "#475569", whiteSpace: "pre-wrap", lineHeight: 1.6 }}>{data?.body}</div>
          </div>
        </div>
        <div style={{ padding: "14px 24px", borderTop: "1px solid #f1f5f9", display: "flex", gap: 10, justifyContent: "flex-end", background: "#fafbff" }}>
          <button onClick={onCancel} disabled={sending} style={{ padding: "10px 20px", borderRadius: 10, border: "1.5px solid #d1d5db", background: "#fff", fontSize: 14, fontWeight: 600, color: "#64748b", cursor: "pointer" }}><I.X /> Abbrechen</button>
          <button onClick={onConfirm} disabled={sending} style={{ padding: "10px 24px", borderRadius: 10, border: "none", background: sending ? "#86efac" : "linear-gradient(135deg, #16a34a, #15803d)", fontSize: 14, fontWeight: 700, color: "#fff", cursor: sending ? "wait" : "pointer", display: "flex", alignItems: "center", gap: 8, boxShadow: "0 4px 16px rgba(22,163,74,0.3)", minWidth: 160, justifyContent: "center" }}>
            {sending ? "⏳ Wird gesendet..." : <><I.Send /> {sendLabel || "Senden"}</>}
          </button>
        </div>
      </div>
    </div>
  );
}

// ========================
// TOAST
// ========================
function Toast({ show, msg, sub, err }) {
  if (!show) return null;
  return (<div style={{ position: "fixed", top: 24, right: 24, background: "#fff", padding: "14px 20px", borderRadius: 14, boxShadow: "0 12px 40px rgba(0,0,0,0.12)", zIndex: 2000, display: "flex", alignItems: "center", gap: 12, animation: "slideIn .3s ease-out", border: `1px solid ${err ? "#fecaca" : "#dcfce7"}`, maxWidth: 420 }}>
    <div style={{ width: 36, height: 36, borderRadius: 8, background: err ? "#ef4444" : "#16a34a", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><span style={{ color: "#fff" }}>{err ? <I.X /> : <I.Check />}</span></div>
    <div><div style={{ fontWeight: 700, fontSize: 14 }}>{msg}</div>{sub && <div style={{ fontSize: 12, color: "#64748b", marginTop: 1 }}>{sub}</div>}</div>
  </div>);
}

// ========================
// JOBS TAB
// ========================
function JobsTab({ jobs, onRefresh, onSendReminder }) {
  const fmtDate = (iso) => { const d = new Date(iso); return d.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" }) + " " + d.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }); };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ fontSize: 14, color: "#64748b" }}>{jobs.length} Job(s) gesamt</div>
        <button onClick={onRefresh} style={{ border: "none", background: "none", cursor: "pointer", color: "#2563eb", fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}><I.Refresh /> Aktualisieren</button>
      </div>
      {jobs.length === 0 && (
        <div style={{ background: "#fff", borderRadius: 16, padding: 48, textAlign: "center", border: "1px solid #e8edf5" }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>📋</div>
          <h3 style={{ margin: "0 0 8px", fontSize: 18, fontWeight: 700 }}>Noch keine Jobs</h3>
          <p style={{ color: "#64748b", fontSize: 14 }}>Gesendete E-Mails erscheinen hier</p>
        </div>
      )}
      {jobs.map((job) => (
        <div key={job.id} style={{ background: "#fff", borderRadius: 14, border: "1px solid #e8edf5", boxShadow: "0 2px 8px rgba(0,0,0,0.03)", overflow: "hidden" }}>
          {/* Header */}
          <div style={{ padding: "14px 20px", display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: job.emailTyp === "erste" ? "#f5f3ff" : "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <I.Mail />
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontWeight: 700, fontSize: 15, color: "#0f172a" }}>{job.kundeName || job.kundeRaw}</span>
                {job.kundeId && <span style={{ color: "#2563eb", fontSize: 12, fontWeight: 600 }}>#{job.kundeId}</span>}
                <span style={{ background: job.emailTyp === "erste" ? "#f5f3ff" : "#eff6ff", color: job.emailTyp === "erste" ? "#7c3aed" : "#2563eb", fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 6 }}>
                  {job.emailTyp === "erste" ? "Erste" : "Regulär"}
                </span>
                {job.dryRun && <span style={{ background: "#fef3c7", color: "#b45309", fontSize: 10, fontWeight: 700, padding: "2px 6px", borderRadius: 4 }}>DRY</span>}
              </div>
              <div style={{ fontSize: 12, color: "#64748b", marginTop: 2, display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ display: "flex", alignItems: "center", gap: 4 }}><I.Clock /> {fmtDate(job.sentAt)}</span>
                <span>→ {job.to}</span>
                <span>· {job.pdfNames?.length || 0} PDF(s)</span>
                <span>· von {job.signaturPerson}</span>
              </div>
            </div>
            {/* Reminder button */}
            <button onClick={() => onSendReminder(job)} style={{
              padding: "8px 16px", borderRadius: 8,
              border: "1.5px solid #f59e0b", background: "#fffbeb",
              fontSize: 13, fontWeight: 600, color: "#b45309",
              cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
              transition: "all .15s",
            }} onMouseEnter={(e) => { e.currentTarget.style.background = "#fef3c7"; }} onMouseLeave={(e) => { e.currentTarget.style.background = "#fffbeb"; }}>
              <I.Bell /> Erinnerung
            </button>
          </div>
          {/* Reminders sent */}
          {job.reminders?.length > 0 && (
            <div style={{ padding: "0 20px 12px", display: "flex", flexDirection: "column", gap: 4 }}>
              {job.reminders.map((r, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "#16a34a", padding: "4px 10px", background: "#f0fdf4", borderRadius: 6 }}>
                  <I.Check /> Erinnerung gesendet: {fmtDate(r.sentAt)} {r.dryRun && <span style={{ color: "#b45309" }}>(DRY RUN)</span>}
                </div>
              ))}
            </div>
          )}
          {/* PDF names */}
          {job.pdfNames?.length > 0 && (
            <div style={{ padding: "0 20px 14px", display: "flex", flexWrap: "wrap", gap: 6 }}>
              {job.pdfNames.map((name, i) => (
                <span key={i} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#64748b", background: "#f1f5f9", padding: "3px 8px", borderRadius: 6 }}><I.File /> {name}</span>
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// ========================
// KUNDEN TAB
// ========================
function KundenTab({ kunden, jobs, onSendReminder }) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("alle");
  const [expandedKunde, setExpandedKunde] = useState(null);
  const [sortBy, setSortBy] = useState("neueste");

  const fmtDate = (iso) => { const d = new Date(iso); return d.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" }) + " " + d.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }); };
  const fmtShort = (iso) => { const d = new Date(iso); return d.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "2-digit" }); };

  // Group jobs by kundeRaw
  const jobsByKunde = useMemo(() => {
    const map = {};
    jobs.forEach((j) => { const k = j.kundeRaw; if (!map[k]) map[k] = []; map[k].push(j); });
    // Sort each group newest first
    Object.values(map).forEach((arr) => arr.sort((a, b) => new Date(b.sentAt) - new Date(a.sentAt)));
    return map;
  }, [jobs]);

  // Build customer list with stats
  const kundenList = useMemo(() => {
    const seen = new Set();
    const list = [];

    // Customers from kunden array
    kunden.forEach((k) => {
      seen.add(k.raw);
      const kundeJobs = jobsByKunde[k.raw] || [];
      const totalReminders = kundeJobs.reduce((s, j) => s + (j.reminders?.length || 0), 0);
      const lastDate = kundeJobs.length ? kundeJobs[0].sentAt : null;
      list.push({ ...k, jobs: kundeJobs, emailCount: kundeJobs.length, reminderCount: totalReminders, lastDate });
    });

    // Customers that appear only in jobs (not in SharePoint)
    Object.keys(jobsByKunde).forEach((raw) => {
      if (!seen.has(raw)) {
        const kundeJobs = jobsByKunde[raw];
        const first = kundeJobs[0];
        const totalReminders = kundeJobs.reduce((s, j) => s + (j.reminders?.length || 0), 0);
        list.push({
          raw, id: first.kundeId || raw, name: first.kundeName || raw, typ: "Unbekannt",
          email: first.to, contactName: first.empfaengerName || "",
          searchText: `${first.kundeId || ""} ${first.kundeName || ""} ${raw} ${first.to}`.toLowerCase(),
          jobs: kundeJobs, emailCount: kundeJobs.length, reminderCount: totalReminders, lastDate: kundeJobs[0].sentAt,
        });
      }
    });

    return list;
  }, [kunden, jobsByKunde]);

  // Filter
  const filtered = useMemo(() => {
    let list = kundenList;
    if (filter === "aktive") list = list.filter((k) => k.typ === "Aktive Kunden");
    else if (filter === "interessenten") list = list.filter((k) => k.typ === "Interessenten");
    if (search.trim()) {
      const s = search.toLowerCase();
      list = list.filter((k) => (k.searchText || "").includes(s) || (k.email || "").toLowerCase().includes(s) || (k.raw || "").toLowerCase().includes(s));
    }
    return list;
  }, [kundenList, filter, search]);

  // Sort
  const sorted = useMemo(() => {
    const withEmails = filtered.filter((k) => k.emailCount > 0);
    const withoutEmails = filtered.filter((k) => k.emailCount === 0);

    if (sortBy === "neueste") withEmails.sort((a, b) => new Date(b.lastDate) - new Date(a.lastDate));
    else if (sortBy === "aelteste") withEmails.sort((a, b) => new Date(a.lastDate) - new Date(b.lastDate));
    else if (sortBy === "name-az") withEmails.sort((a, b) => a.name.localeCompare(b.name));
    else if (sortBy === "name-za") withEmails.sort((a, b) => b.name.localeCompare(a.name));

    withoutEmails.sort((a, b) => a.name.localeCompare(b.name));
    return [...withEmails, ...withoutEmails];
  }, [filtered, sortBy]);

  const withEmailCount = filtered.filter((k) => k.emailCount > 0).length;
  const isAktiv = (k) => k.typ === "Aktive Kunden";

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {/* Search + Filter row */}
      <div style={{ background: "#fff", borderRadius: 14, padding: "14px 20px", border: "1px solid #e8edf5", display: "flex", flexDirection: "column", gap: 10 }}>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <div style={{ flex: 1, display: "flex", alignItems: "center", gap: 8, background: "#f8fafc", borderRadius: 8, padding: "8px 12px", border: "1px solid #e2e8f0" }}>
            <I.Search />
            <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Kunde suchen (Name, Nr., E-Mail)..." style={{ border: "none", outline: "none", background: "transparent", fontSize: 14, width: "100%", color: "#1e293b" }} />
            {search && <span style={{ cursor: "pointer", color: "#94a3b8" }} onClick={() => setSearch("")}><I.X /></span>}
          </div>
          <div style={{ display: "flex", gap: 2 }}>
            {[{ key: "alle", label: "Alle" }, { key: "aktive", label: "Aktive" }, { key: "interessenten", label: "Interessenten" }].map((t) => (
              <button key={t.key} onClick={() => setFilter(t.key)} style={{ padding: "8px 14px", fontSize: 12, fontWeight: filter === t.key ? 700 : 500, color: filter === t.key ? "#2563eb" : "#64748b", background: filter === t.key ? "#eff6ff" : "transparent", border: "none", borderRadius: 6, cursor: "pointer" }}>{t.label}</button>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 13, color: "#64748b" }}>
            <strong>{withEmailCount}</strong> mit E-Mails · {sorted.length} gesamt
          </div>
          <div style={{ display: "flex", gap: 2 }}>
            {[{ key: "neueste", label: "Neueste" }, { key: "aelteste", label: "Älteste" }, { key: "name-az", label: "A–Z" }, { key: "name-za", label: "Z–A" }].map((s) => (
              <button key={s.key} onClick={() => setSortBy(s.key)} style={{ padding: "4px 10px", fontSize: 11, fontWeight: sortBy === s.key ? 700 : 500, color: sortBy === s.key ? "#fff" : "#64748b", background: sortBy === s.key ? "#2563eb" : "#f1f5f9", border: "none", borderRadius: 5, cursor: "pointer" }}>{s.label}</button>
            ))}
          </div>
        </div>
      </div>

      {/* Empty state */}
      {sorted.length === 0 && (
        <div style={{ background: "#fff", borderRadius: 16, padding: 48, textAlign: "center", border: "1px solid #e8edf5" }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>👥</div>
          <h3 style={{ margin: "0 0 8px", fontSize: 18, fontWeight: 700 }}>Keine Kunden gefunden</h3>
          <p style={{ color: "#64748b", fontSize: 14 }}>{search ? "Suchbegriff ändern" : "Noch keine Kunden vorhanden"}</p>
        </div>
      )}

      {/* Customer cards */}
      {sorted.map((k) => {
        const expanded = expandedKunde === k.raw;
        const hasEmails = k.emailCount > 0;
        const dotColor = isAktiv(k) ? "#16a34a" : k.typ === "Interessenten" ? "#f59e0b" : "#94a3b8";
        const typLabel = isAktiv(k) ? "Aktive Kunden" : k.typ === "Interessenten" ? "Interessenten" : "";

        return (
          <div key={k.raw} style={{ background: "#fff", borderRadius: 14, border: expanded ? "2px solid #2563eb" : "1px solid #e8edf5", boxShadow: "0 2px 8px rgba(0,0,0,0.03)", overflow: "hidden", transition: "border .15s" }}>
            {/* Customer header — click to expand */}
            <div onClick={() => hasEmails && setExpandedKunde(expanded ? null : k.raw)} style={{ padding: "14px 20px", display: "flex", alignItems: "center", gap: 14, cursor: hasEmails ? "pointer" : "default", opacity: hasEmails ? 1 : 0.6 }}
              onMouseEnter={(e) => { if (hasEmails) e.currentTarget.style.background = "#f8fafc"; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}>
              {/* Type dot */}
              <div style={{ width: 40, height: 40, borderRadius: 10, background: isAktiv(k) ? "#f0fdf4" : k.typ === "Interessenten" ? "#fffbeb" : "#f8fafc", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <span style={{ width: 12, height: 12, borderRadius: "50%", background: dotColor }} />
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontWeight: 700, color: "#2563eb", fontSize: 13 }}>{k.id}</span>
                  <span style={{ fontWeight: 700, fontSize: 15, color: "#0f172a" }}>{k.name}</span>
                  {typLabel && <span style={{ background: isAktiv(k) ? "#f0fdf4" : "#fffbeb", color: isAktiv(k) ? "#16a34a" : "#d97706", fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 6 }}>{typLabel}</span>}
                </div>
                <div style={{ fontSize: 12, color: "#64748b", marginTop: 3, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                  {k.email && <span style={{ color: "#16a34a" }}>📧 {k.email}</span>}
                  {k.contactName && <span>· {k.contactName}</span>}
                  {hasEmails ? (
                    <>
                      <span style={{ color: "#2563eb", fontWeight: 600 }}>· {k.emailCount} E-Mail{k.emailCount !== 1 ? "s" : ""}</span>
                      {k.reminderCount > 0 && <span style={{ color: "#f59e0b" }}>· {k.reminderCount} Erinnerung{k.reminderCount !== 1 ? "en" : ""}</span>}
                      <span style={{ display: "flex", alignItems: "center", gap: 3 }}><I.Clock /> {fmtShort(k.lastDate)}</span>
                    </>
                  ) : <span style={{ fontStyle: "italic" }}>Keine E-Mails gesendet</span>}
                </div>
              </div>
              {hasEmails && (
                <div style={{ transform: expanded ? "rotate(180deg)" : "rotate(0)", transition: "transform .2s", color: "#94a3b8" }}>
                  <I.ChevronDown />
                </div>
              )}
            </div>

            {/* Expanded email history */}
            {expanded && k.jobs.length > 0 && (
              <div style={{ padding: "0 20px 16px", display: "flex", flexDirection: "column", gap: 8 }}>
                <div style={{ borderTop: "1px solid #e2e8f0", paddingTop: 12, marginBottom: 4, fontSize: 12, fontWeight: 700, color: "#475569" }}>E-Mail Verlauf</div>
                {k.jobs.map((job) => (
                  <div key={job.id} style={{ background: "#f8fafc", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "hidden" }}>
                    <div style={{ padding: "10px 14px", display: "flex", alignItems: "center", gap: 10 }}>
                      <div style={{ width: 32, height: 32, borderRadius: 8, background: job.emailTyp === "erste" ? "#f5f3ff" : "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                        <I.Mail />
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                          <span style={{ fontSize: 13, fontWeight: 600, color: "#0f172a", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{job.subject}</span>
                          <span style={{ background: job.emailTyp === "erste" ? "#f5f3ff" : "#eff6ff", color: job.emailTyp === "erste" ? "#7c3aed" : "#2563eb", fontSize: 10, fontWeight: 700, padding: "1px 6px", borderRadius: 4 }}>
                            {job.emailTyp === "erste" ? "Erste" : "Regulär"}
                          </span>
                          {job.dryRun && <span style={{ background: "#fef3c7", color: "#b45309", fontSize: 10, fontWeight: 700, padding: "1px 5px", borderRadius: 4 }}>DRY</span>}
                        </div>
                        <div style={{ fontSize: 11, color: "#64748b", marginTop: 2, display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{ display: "flex", alignItems: "center", gap: 3 }}><I.Clock /> {fmtDate(job.sentAt)}</span>
                          <span>· {job.pdfNames?.length || 0} PDF(s)</span>
                          <span>· von {job.signaturPerson}</span>
                        </div>
                      </div>
                      <button onClick={(e) => { e.stopPropagation(); onSendReminder(job); }} style={{
                        padding: "6px 12px", borderRadius: 6, border: "1px solid #f59e0b", background: "#fffbeb",
                        fontSize: 11, fontWeight: 600, color: "#b45309", cursor: "pointer", display: "flex", alignItems: "center", gap: 4, flexShrink: 0,
                      }} onMouseEnter={(e) => { e.currentTarget.style.background = "#fef3c7"; }} onMouseLeave={(e) => { e.currentTarget.style.background = "#fffbeb"; }}>
                        <I.Bell /> Erinnerung
                      </button>
                    </div>
                    {/* Reminders for this email */}
                    {job.reminders?.length > 0 && (
                      <div style={{ padding: "0 14px 10px", display: "flex", flexDirection: "column", gap: 3 }}>
                        {job.reminders.map((r, i) => (
                          <div key={i} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11, color: "#16a34a", padding: "3px 8px", background: "#f0fdf4", borderRadius: 4 }}>
                            <I.Check /> Erinnerung: {fmtDate(r.sentAt)} {r.dryRun && <span style={{ color: "#b45309" }}>(DRY)</span>}
                          </div>
                        ))}
                      </div>
                    )}
                    {/* PDF names for this email */}
                    {job.pdfNames?.length > 0 && (
                      <div style={{ padding: "0 14px 10px", display: "flex", flexWrap: "wrap", gap: 4 }}>
                        {job.pdfNames.map((name, i) => (
                          <span key={i} style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 10, color: "#64748b", background: "#fff", padding: "2px 6px", borderRadius: 4, border: "1px solid #e2e8f0" }}><I.File /> {name}</span>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ========================
// MAIN APP
// ========================
export default function App() {
  const [activeTab, setActiveTab] = useState("verarbeiten");
  const [kunden, setKunden] = useState([]);
  const [kundenLoading, setKundenLoading] = useState(true);
  const [jobs, setJobs] = useState([]);
  const [selectedKunde, setSelectedKunde] = useState(null);
  const [gewerk, setGewerk] = useState(""); const [region, setRegion] = useState("");
  const [pdfs, setPdfs] = useState([]);
  const [emailTyp, setEmailTyp] = useState("regular");
  const [anrede, setAnrede] = useState("Herr"); const [empfaengerName, setEmpfaengerName] = useState("");
  const [kundeEmail, setKundeEmail] = useState("");
  const [signaturPerson, setSignaturPerson] = useState("Dennis Engel");
  const [isEditing, setIsEditing] = useState(false); const [emailBody, setEmailBody] = useState("");
  const [showConfirm, setShowConfirm] = useState(false); const [sending, setSending] = useState(false);
  const [previewId, setPreviewId] = useState(null); const [previewPdfUrl, setPreviewPdfUrl] = useState(null); const [previewPdfName, setPreviewPdfName] = useState(null);
  const [generating, setGenerating] = useState(false);
  const [reminderJob, setReminderJob] = useState(null); const [reminderSending, setReminderSending] = useState(false);
  const [toast, setToast] = useState(null);
  const [serverStatus, setServerStatus] = useState(null);
  const [vorlage, setVorlage] = useState(null);
  const [vorlageUploading, setVorlageUploading] = useState(false);
  const [pipelineRunning, setPipelineRunning] = useState(false);
  const fileRef = useRef(null);
  const vorlageRef = useRef(null);

  const kunde = selectedKunde ? kunden.find((k) => k.raw === selectedKunde) : null;

  const loadKunden = useCallback(async () => { setKundenLoading(true); try { const d = await api.getKunden(); setKunden([...(d.aktive || []), ...(d.interessenten || [])]); } catch (e) { showToast(e.message, null, true); } finally { setKundenLoading(false); } }, []);
  const loadJobs = useCallback(async () => { try { setJobs(await api.getJobs()); } catch (e) { console.error(e); } }, []);

  const loadVorlage = useCallback(async () => { try { setVorlage(await api.getVorlage()); } catch (e) { console.error(e); } }, []);
  useEffect(() => { api.health().then(setServerStatus).catch(() => {}); loadKunden(); loadJobs(); loadVorlage(); }, [loadKunden, loadJobs, loadVorlage]);

  useEffect(() => { if (kunde) { setEmpfaengerName(kunde.contactName || kunde.name.split(" ").pop()); if (kunde.email) setKundeEmail(kunde.email); else setKundeEmail(""); } }, [kunde]);

  useEffect(() => { if (!isEditing) { setEmailBody(EMAIL_TEMPLATES[emailTyp](anrede, empfaengerName || "xxx", SIGNATUREN[signaturPerson]?.full || "")); } }, [emailTyp, anrede, empfaengerName, signaturPerson, isEditing]);

  const showToast = (msg, sub, err) => { setToast({ msg, sub, err }); setTimeout(() => setToast(null), 4000); };
  const handleVorlageUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setVorlageUploading(true);
    try {
      await api.uploadVorlage(file);
      await loadVorlage();
      showToast("Vorlage hochgeladen", file.name);
    } catch (err) { showToast("Fehler", err.message, true); }
    finally { setVorlageUploading(false); }
  };

  const handleDrop = (e) => { e.preventDefault(); setPdfs((p) => [...p, ...Array.from(e.dataTransfer.files).filter((f) => f.type === "application/pdf")]); };
  const handleEmailBlur = () => { if (kunde && kundeEmail.trim()) api.saveEmail(kunde.raw, kundeEmail, empfaengerName.trim() || undefined); };
  const handleNameBlur = () => { if (kunde && empfaengerName.trim()) api.saveEmail(kunde.raw, kundeEmail.trim() || undefined, empfaengerName); };

  const emailSubject = emailTyp === "erste" ? `Erste Recherche ${kunde?.name || ""} | KALKU` : `Aktuelle Recherche ${kunde?.name || ""} | KALKU`;
  const canSend = kunde && kundeEmail.trim() && empfaengerName.trim() && pdfs.length > 0 && vorlage?.exists;

  // Step 1: Generate PDF and show preview in confirmation modal
  const handleStartPipeline = async () => {
    setGenerating(true);
    try {
      const result = await api.generatePipeline({ kundeName: kunde.name, gewerk, region, pdfs });
      setPreviewId(result.previewId);
      setPreviewPdfName(result.attachmentName);
      setPreviewPdfUrl(`/api/pipeline/preview/${result.previewId}/${encodeURIComponent(result.attachmentName)}`);
      setShowConfirm(true);
    } catch (e) {
      console.error('Pipeline generation error:', e);
      showToast("Fehler beim Generieren", e.message || String(e), true);
    }
    finally { setGenerating(false); }
  };

  // Step 2: Send the pre-generated email after user reviews PDF
  const handleConfirmSend = async () => {
    setSending(true);
    setPipelineRunning(true);
    try {
      const kundeTyp = kunde.typ === "Interessenten" ? "Interessenten" : "Aktive_Kunden";
      const result = await api.sendPipeline({ previewId, to: kundeEmail, subject: emailSubject, body: emailBody, kundeRaw: kunde.raw, kundeName: kunde.name, kundeId: kunde.id, kundeTyp, anrede, empfaengerName, signaturPerson, emailTyp, gewerk, region });
      setSending(false); setPipelineRunning(false); setShowConfirm(false);
      setPreviewId(null); setPreviewPdfUrl(null); setPreviewPdfName(null);
      const msg = result.dryRun ? "DRY RUN — nicht gesendet" : "Pipeline abgeschlossen!";
      const sub = `${kunde.name} · ${result.generatedFile} · SharePoint: ${result.sharePointPath}`;
      showToast(msg, sub);
      loadJobs();
      setTimeout(() => { setPdfs([]); setSelectedKunde(null); setKundeEmail(""); setEmpfaengerName(""); setIsEditing(false); }, 1500);
    } catch (e) { setSending(false); setPipelineRunning(false); showToast("Fehler", e.message, true); }
  };

  // --- Reminder ---
  const handleSendReminder = (job) => { setReminderJob(job); };

  const handleConfirmReminder = async () => {
    if (!reminderJob) return;
    setReminderSending(true);
    try {
      const txt = buildReminderText(reminderJob.anrede, reminderJob.empfaengerName, reminderJob.signaturPerson);
      const html = buildReminderHtml(reminderJob.anrede, reminderJob.empfaengerName, reminderJob.signaturPerson);
      const subj = `Erinnerung – KALKU | ${reminderJob.kundeName}`;
      const result = await api.sendReminder(reminderJob.id, txt, html, subj);
      setReminderSending(false); setReminderJob(null);
      showToast(result.dryRun ? "DRY RUN — Erinnerung nicht gesendet" : "Erinnerung gesendet!", `An: ${reminderJob.to}`);
      loadJobs();
    } catch (e) { setReminderSending(false); showToast("Fehler", e.message, true); }
  };

  const reminderPreview = reminderJob ? buildReminderText(reminderJob.anrede, reminderJob.empfaengerName, reminderJob.signaturPerson) : "";

  const tabs = [
    { key: "verarbeiten", label: "Verarbeiten & Senden", icon: <I.Send /> },
    { key: "jobs", label: "Jobs", icon: <I.Briefcase />, badge: jobs.length },
    { key: "kunden", label: "Kunden", icon: <I.Users />, badge: kunden.length },
    { key: "einstellungen", label: "Einstellungen", icon: <I.Settings /> },
  ];

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(160deg, #eef4ff 0%, #f0f7ff 40%, #fafbff 100%)", fontFamily: "'DM Sans','Segoe UI',system-ui,sans-serif", color: "#1e293b" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');
        @keyframes dropIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:translateY(0)}}
        @keyframes fadeIn{from{opacity:0}to{opacity:1}}
        @keyframes modalIn{from{opacity:0;transform:scale(.95) translateY(10px)}to{opacity:1;transform:scale(1) translateY(0)}}
        @keyframes slideIn{from{opacity:0;transform:translateX(40px)}to{opacity:1;transform:translateX(0)}}
        input::placeholder,textarea::placeholder{color:#94a3b8}*{box-sizing:border-box}
        ::-webkit-scrollbar{width:6px}::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:3px}
      `}</style>

      {/* Send confirm */}
      <ConfirmModal open={showConfirm} data={{ to: kundeEmail, subject: emailSubject, body: emailBody }} onConfirm={handleConfirmSend} onCancel={() => { setShowConfirm(false); setPreviewId(null); setPreviewPdfUrl(null); setPreviewPdfName(null); }} sending={sending} title="E-Mail senden bestätigen" pdfPreviewUrl={previewPdfUrl} pdfPreviewName={previewPdfName} />

      {/* Reminder confirm */}
      <ConfirmModal open={!!reminderJob} data={{ to: reminderJob?.to, subject: `Erinnerung – KALKU | ${reminderJob?.kundeName || ""}`, body: reminderPreview }} onConfirm={handleConfirmReminder} onCancel={() => setReminderJob(null)} sending={reminderSending} title="Erinnerung senden"
        pdfPreviewUrl={reminderJob?.savedAttachments?.length ? `/api/jobs/${reminderJob.id}/attachments/${encodeURIComponent(reminderJob.savedAttachments[0])}` : null}
        pdfPreviewName={reminderJob?.savedAttachments?.[0] || null}
        sendLabel="Erinnerung senden"
      />

      {toast && <Toast show={true} msg={toast.msg} sub={toast.sub} err={toast.err} />}

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "28px 24px 0" }}>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 24 }}>
          <div>
            <div style={{ display: "inline-flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
              <span style={{ background: "linear-gradient(135deg, #2563eb, #1d4ed8)", color: "#fff", padding: "5px 14px", borderRadius: 6, fontSize: 12, fontWeight: 700 }}>KALKU Tender Tool v2</span>
              {serverStatus?.dryRun && <span style={{ background: "#fef3c7", color: "#b45309", padding: "4px 10px", borderRadius: 6, fontSize: 11, fontWeight: 700 }}>⚠ DRY RUN</span>}
            </div>
            <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800, letterSpacing: "-.02em" }}>Ausschreibungs-Automatisierung</h1>
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            {[{ l: "Kunden", v: kunden.length, c: "#1e293b" }, { l: "Gesendet", v: jobs.length, c: "#16a34a" }].map((s, i) => (
              <div key={i} style={{ padding: "8px 16px", background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", textAlign: "center", minWidth: 72 }}>
                <div style={{ fontSize: 11, color: "#64748b", fontWeight: 500 }}>{s.l}</div>
                <div style={{ fontSize: 17, fontWeight: 700, color: s.c }}>{s.v}</div>
              </div>))}
          </div>
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", gap: 4, marginBottom: 24 }}>
          {tabs.map((t) => (
            <button key={t.key} onClick={() => setActiveTab(t.key)} style={{
              padding: "10px 20px", borderRadius: 10, border: "none", fontSize: 14, fontWeight: activeTab === t.key ? 700 : 500,
              color: activeTab === t.key ? "#fff" : "#64748b",
              background: activeTab === t.key ? "linear-gradient(135deg, #2563eb, #1d4ed8)" : "transparent",
              cursor: "pointer", display: "flex", alignItems: "center", gap: 8,
              boxShadow: activeTab === t.key ? "0 4px 12px rgba(37,99,235,.25)" : "none",
            }}>{t.icon} {t.label}
              {t.badge > 0 && <span style={{ background: activeTab === t.key ? "rgba(255,255,255,.25)" : "#e2e8f0", color: activeTab === t.key ? "#fff" : "#64748b", fontSize: 11, fontWeight: 700, padding: "1px 8px", borderRadius: 10 }}>{t.badge}</span>}
            </button>))}
        </div>

        {/* ===== VERARBEITEN TAB ===== */}
        {activeTab === "verarbeiten" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            {/* LEFT */}
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <div style={{ background: "#fff", borderRadius: 16, padding: 20, border: "1px solid #e8edf5", boxShadow: "0 2px 12px rgba(0,0,0,.04)" }}>
                <h2 style={{ margin: "0 0 14px", fontSize: 14, fontWeight: 700, display: "flex", alignItems: "center", gap: 8 }}><I.Users /> Kunde & Einstellungen</h2>
                <div style={{ marginBottom: 12 }}><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>Kunde (SharePoint)</label><KundenDropdown kunden={kunden} value={selectedKunde} onChange={setSelectedKunde} loading={kundenLoading} onRefresh={loadKunden} /></div>
                <div style={{ marginBottom: 12 }}><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>E-Mail {kunde?.email && <span style={{ color: "#16a34a", fontWeight: 400 }}>(gespeichert)</span>}</label><input value={kundeEmail} onChange={(e) => setKundeEmail(e.target.value)} onBlur={handleEmailBlur} placeholder="E-Mail — wird gespeichert" style={{ width: "100%", padding: "10px 14px", border: `1.5px solid ${kundeEmail ? "#bbf7d0" : "#d1d5db"}`, borderRadius: 10, fontSize: 14, outline: "none", background: kundeEmail ? "#f0fdf4" : "#fff" }} onFocus={(e) => e.target.style.borderColor = "#2563eb"} /></div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  <div><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>Gewerk</label><input value={gewerk} onChange={(e) => setGewerk(e.target.value)} placeholder="z.B. Gebäudereinigung" style={{ width: "100%", padding: "10px 14px", border: "1.5px solid #d1d5db", borderRadius: 10, fontSize: 14, outline: "none" }} /></div>
                  <div><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>Region</label><input value={region} onChange={(e) => setRegion(e.target.value)} placeholder="z.B. 66111 + 50km" style={{ width: "100%", padding: "10px 14px", border: "1.5px solid #d1d5db", borderRadius: 10, fontSize: 14, outline: "none" }} /></div>
                </div>
              </div>
              <div style={{ background: "#fff", borderRadius: 16, padding: 20, border: "1px solid #e8edf5", boxShadow: "0 2px 12px rgba(0,0,0,.04)" }}>
                <h2 style={{ margin: "0 0 14px", fontSize: 14, fontWeight: 700, display: "flex", alignItems: "center", gap: 8 }}><I.File /> PDFs</h2>
                <div onDrop={handleDrop} onDragOver={(e) => e.preventDefault()} onClick={() => fileRef.current?.click()} style={{ border: "2px dashed #cbd5e1", borderRadius: 12, padding: pdfs.length ? "12px" : "28px", textAlign: "center", cursor: "pointer", background: "#fafbff" }}>
                  <input ref={fileRef} type="file" accept=".pdf" multiple onChange={(e) => setPdfs((p) => [...p, ...Array.from(e.target.files).filter((f) => f.type === "application/pdf")])} style={{ display: "none" }} />
                  {!pdfs.length ? <><div style={{ color: "#94a3b8" }}><I.Upload /></div><p style={{ margin: "6px 0 0", color: "#64748b", fontSize: 13 }}>PDFs hierher ziehen</p></> :
                    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>{pdfs.map((f, i) => (<div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", background: "#fff", borderRadius: 8, border: "1px solid #e2e8f0", textAlign: "left" }} onClick={(e) => e.stopPropagation()}><I.File /><span style={{ flex: 1, fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span><span onClick={(e) => { e.stopPropagation(); setPdfs((p) => p.filter((_, j) => j !== i)); }} style={{ cursor: "pointer", color: "#ef4444" }}><I.Trash /></span></div>))}</div>}
                </div>
              </div>
              {/* VORLAGE UPLOAD */}
              <div style={{ background: "#fff", borderRadius: 16, padding: 20, border: `1px solid ${vorlage?.exists ? "#bbf7d0" : "#fecaca"}`, boxShadow: "0 2px 12px rgba(0,0,0,.04)" }}>
                <h2 style={{ margin: "0 0 10px", fontSize: 14, fontWeight: 700, display: "flex", alignItems: "center", gap: 8 }}><I.File /> DOCX-Vorlage {vorlage?.exists && <span style={{ color: "#16a34a", fontSize: 11, fontWeight: 600, background: "#f0fdf4", padding: "2px 8px", borderRadius: 6 }}>Aktiv</span>}</h2>
                <input ref={vorlageRef} type="file" accept=".docx" onChange={handleVorlageUpload} style={{ display: "none" }} />
                <div onClick={() => vorlageRef.current?.click()} style={{ border: "2px dashed #cbd5e1", borderRadius: 10, padding: "14px", textAlign: "center", cursor: vorlageUploading ? "wait" : "pointer", background: vorlage?.exists ? "#f0fdf4" : "#fff7ed" }}>
                  {vorlageUploading ? <span style={{ fontSize: 13, color: "#64748b" }}>Hochladen...</span> : vorlage?.exists ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 10, justifyContent: "center" }}>
                      <I.Check /><span style={{ fontSize: 13, fontWeight: 600, color: "#16a34a" }}>Vorlage vorhanden</span>
                      <span style={{ fontSize: 11, color: "#64748b" }}>({(vorlage.size / 1024).toFixed(0)} KB)</span>
                      <span style={{ fontSize: 11, color: "#2563eb", fontWeight: 600 }}>Ersetzen</span>
                    </div>
                  ) : (
                    <div><div style={{ color: "#f59e0b" }}><I.Upload /></div><p style={{ margin: "4px 0 0", color: "#b45309", fontSize: 13, fontWeight: 600 }}>DOCX-Vorlage hochladen</p></div>
                  )}
                </div>
              </div>
            </div>
            {/* RIGHT */}
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <div style={{ background: "#fff", borderRadius: 16, padding: 20, border: "1px solid #e8edf5", boxShadow: "0 2px 12px rgba(0,0,0,.04)" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                  <h2 style={{ margin: 0, fontSize: 14, fontWeight: 700, display: "flex", alignItems: "center", gap: 8 }}><I.Mail /> E-Mail</h2>
                  <div style={{ fontSize: 11, color: "#16a34a", background: "#f0fdf4", padding: "4px 10px", borderRadius: 6, border: "1px solid #bbf7d0", fontWeight: 600, display: "flex", alignItems: "center", gap: 4 }}><I.Pipedrive /> BCC aktiv</div>
                </div>
                <div style={{ marginBottom: 12 }}><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>Vorlage</label><div style={{ display: "flex", gap: 6 }}>{[{ k: "regular", l: "Reguläre Recherche", c: "#2563eb", bg: "#eff6ff" }, { k: "erste", l: "Erste Recherche", c: "#7c3aed", bg: "#f5f3ff" }].map((t) => (<button key={t.k} onClick={() => { setEmailTyp(t.k); setIsEditing(false); }} style={{ flex: 1, padding: "10px 12px", borderRadius: 10, textAlign: "left", border: emailTyp === t.k ? `2px solid ${t.c}` : "1.5px solid #d1d5db", background: emailTyp === t.k ? t.bg : "#fff", fontSize: 13, fontWeight: emailTyp === t.k ? 700 : 500, color: emailTyp === t.k ? t.c : "#64748b", cursor: "pointer" }}>{emailTyp === t.k && "✓ "}{t.l}</button>))}</div></div>
                <div style={{ display: "grid", gridTemplateColumns: "120px 1fr", gap: 10, marginBottom: 12 }}>
                  <div><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>Anrede</label><div style={{ display: "flex", gap: 4 }}>{["Herr", "Frau"].map((a) => (<button key={a} onClick={() => { setAnrede(a); setIsEditing(false); }} style={{ flex: 1, padding: "9px 6px", borderRadius: 8, border: anrede === a ? "2px solid #2563eb" : "1.5px solid #d1d5db", background: anrede === a ? "#eff6ff" : "#fff", fontSize: 13, fontWeight: anrede === a ? 700 : 500, color: anrede === a ? "#2563eb" : "#64748b", cursor: "pointer" }}>{a}</button>))}</div></div>
                  <div><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>Nachname {kunde?.contactName && <span style={{ color: "#16a34a", fontWeight: 400 }}>(gespeichert)</span>}</label><input value={empfaengerName} onChange={(e) => { setEmpfaengerName(e.target.value); setIsEditing(false); }} onBlur={handleNameBlur} placeholder="z.B. Müller" style={{ width: "100%", padding: "9px 14px", border: `1.5px solid ${kunde?.contactName ? "#bbf7d0" : "#d1d5db"}`, borderRadius: 10, fontSize: 14, outline: "none", background: kunde?.contactName ? "#f0fdf4" : "#fff" }} /></div>
                </div>
                <div><label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#475569", marginBottom: 6 }}>Signatur</label><div style={{ display: "flex", gap: 6 }}>{Object.entries(SIGNATUREN).map(([n, d]) => (<button key={n} onClick={() => { setSignaturPerson(n); setIsEditing(false); }} style={{ flex: 1, padding: "10px 8px", borderRadius: 10, textAlign: "center", border: signaturPerson === n ? "2px solid #2563eb" : "1.5px solid #d1d5db", background: signaturPerson === n ? "#eff6ff" : "#fff", cursor: "pointer" }}><div style={{ fontSize: 13, fontWeight: signaturPerson === n ? 700 : 500, color: signaturPerson === n ? "#2563eb" : "#1e293b" }}>{n}</div><div style={{ fontSize: 10, color: "#94a3b8", marginTop: 2 }}>{d.rolle}</div></button>))}</div></div>
              </div>
              <div style={{ background: "#fff", borderRadius: 16, border: "1px solid #e8edf5", boxShadow: "0 2px 12px rgba(0,0,0,.04)", overflow: "hidden", flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 20px", background: "#f8fafc", borderBottom: "1px solid #e2e8f0" }}>
                  <div style={{ fontSize: 13 }}><strong style={{ color: "#475569" }}>Betreff:</strong> {emailSubject}</div>
                  <button onClick={() => setIsEditing(!isEditing)} style={{ border: "none", background: "none", cursor: "pointer", color: "#2563eb", fontSize: 12, fontWeight: 600, display: "flex", alignItems: "center", gap: 4 }}>{isEditing ? <><I.Eye /> Vorschau</> : <><I.Edit /> Bearbeiten</>}</button>
                </div>
                {isEditing ? <textarea value={emailBody} onChange={(e) => setEmailBody(e.target.value)} style={{ width: "100%", minHeight: 260, padding: 16, border: "none", outline: "none", fontSize: 13, lineHeight: 1.6, resize: "vertical", fontFamily: "inherit", background: "#fffef5" }} />
                  : <div style={{ padding: 16, fontSize: 13, lineHeight: 1.7, whiteSpace: "pre-wrap", maxHeight: 300, overflow: "auto" }}>{emailBody}</div>}
                <div style={{ padding: "10px 20px", borderTop: "1px solid #f1f5f9", background: "#fafbff", display: "flex", alignItems: "center", gap: 12, fontSize: 12, color: "#64748b" }}>
                  <span>Von: <strong>kundenservice@kalku.de</strong></span><span style={{ color: "#e2e8f0" }}>|</span><span>An: <strong>{kundeEmail || "—"}</strong></span><span style={{ color: "#e2e8f0" }}>|</span><span style={{ color: "#16a34a" }}><I.Pipedrive /> BCC</span><span style={{ color: "#e2e8f0" }}>|</span><span>{pdfs.length} PDF(s)</span>
                </div>
              </div>
            </div>
            <div style={{ gridColumn: "1/-1" }}>
              <button onClick={handleStartPipeline} disabled={!canSend || generating} style={{ width: "100%", padding: "16px", borderRadius: 14, border: "none", background: generating ? "#86efac" : canSend ? "linear-gradient(135deg, #16a34a, #15803d)" : "#d1d5db", fontSize: 16, fontWeight: 700, color: "#fff", cursor: generating ? "wait" : canSend ? "pointer" : "not-allowed", display: "flex", alignItems: "center", justifyContent: "center", gap: 10, boxShadow: canSend && !generating ? "0 6px 20px rgba(22,163,74,.3)" : "none" }}>
                {generating ? "⏳ PDF wird generiert..." : <><I.Send />{canSend ? `Pipeline starten — ${pdfs.length} PDF(s) → DOCX → E-Mail an ${kunde?.name}` : [!kunde && "Kunde wählen", !kundeEmail && kunde && "E-Mail eingeben", pdfs.length === 0 && "PDFs hochladen", !vorlage?.exists && "Vorlage hochladen"].filter(Boolean).join(" · ")}</>}
              </button>
            </div>
          </div>
        )}

        {/* ===== JOBS TAB ===== */}
        {activeTab === "jobs" && <JobsTab jobs={jobs} onRefresh={loadJobs} onSendReminder={handleSendReminder} />}

        {/* ===== KUNDEN TAB ===== */}
        {activeTab === "kunden" && <KundenTab kunden={kunden} jobs={jobs} onSendReminder={handleSendReminder} />}

        {activeTab === "einstellungen" && (
          <div style={{ background: "#fff", borderRadius: 16, padding: 48, textAlign: "center", border: "1px solid #e8edf5" }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>⚙️</div>
            <h3 style={{ margin: "0 0 8px", fontSize: 18, fontWeight: 700 }}>Einstellungen</h3>
            <p style={{ color: "#64748b", fontSize: 14 }}>Kommt bald</p>
          </div>
        )}
        <div style={{ height: 40 }} />
      </div>
    </div>
  );
}
