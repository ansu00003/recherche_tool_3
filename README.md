# KALKU Tender Tool v2

Ausschreibungs-Automatisierung mit SharePoint, Pipedrive BCC & SMTP.

## Setup

```bash
# 1. Dependencies installieren
npm install

# 2. .env anpassen (Pipedrive BCC, SMTP, SharePoint Credentials)
#    DRY_RUN=true  → Emails werden NICHT gesendet (Testmodus)
#    DRY_RUN=false → Emails werden WIRKLICH gesendet

# 3. Starten (Backend + Frontend parallel)
npm run dev
```

- **Frontend**: http://localhost:5173
- **Backend**: http://localhost:3001

## Architektur

```
Frontend (React/Vite :5173)
    ↓ /api/*
Backend (Express :3001)
    ├── GET  /api/kunden      → SharePoint (MS Graph) → Aktive_Kunden + Interessenten Ordner
    ├── POST /api/kunden/email → E-Mail für Kunden speichern (data/customer-emails.json)
    ├── POST /api/send         → SMTP senden + Pipedrive Smart BCC
    └── GET  /api/health       → Server-Status
```

## Features

### Kunden aus SharePoint
- Liest Ordner aus `Aktive_Kunden/` und `Interessenten/` im SharePoint Drive
- Format: `270_Hassan_Turen` → Suchbar nach Nr (270), Name (Hassan, Turen), oder komplett
- Filter: Alle / Aktive Kunden / Interessenten

### E-Mail Caching
- Beim ersten Mal: E-Mail manuell eingeben
- Wird automatisch gespeichert in `data/customer-emails.json`
- Beim nächsten Mal: E-Mail wird automatisch ausgefüllt

### E-Mail Vorlagen
- **Reguläre Recherche** — für bestehende Kunden
- **Erste Recherche** — Willkommens-E-Mail für neue Interessenten
- Herr/Frau Auswahl + Nachname
- 3 Signaturen: Dennis Engel, Julian Kallenborn, Anna Buxbaum
- E-Mail ist vor dem Senden editierbar

### Pipedrive Integration
- Jede E-Mail bekommt automatisch `kalku@pipedrivemail.com` als BCC
- Pipedrive erkennt den Kontakt und loggt die E-Mail dort
- Erscheint im Pipedrive-Kontakt mit ✉️ Icon genau wie bisher

### SMTP
- Versendet über `kundenservice@kalku.de`
- SMTP Auth: `recherche@kalku.de` @ kasserver.com
- PDF-Anhänge werden als Multipart mitgeschickt

## DRY_RUN Modus
Solange `DRY_RUN=true` in der `.env`:
- Alles funktioniert normal (UI, Bestätigung, etc.)
- Die E-Mail wird aber NICHT gesendet
- Im Terminal siehst du was gesendet WORDEN WÄRE
- E-Mail-Cache wird trotzdem aktualisiert

→ Wenn alles passt: `DRY_RUN=false` setzen und Server neustarten.

## SharePoint Ordner-Struktur
```
SharePoint Drive Root/
├── Aktive_Kunden/
│   ├── 270_Hassan_Turen/
│   ├── 271_Mueller_Bau/
│   └── ...
└── Interessenten/
    ├── 300_Schmidt_GmbH/
    └── ...
```

## Hinweis: FROM-Adresse
Die SMTP-Zugangsdaten sind für `recherche@kalku.de`, aber die FROM-Adresse ist `kundenservice@kalku.de`.
Falls der Mailserver das nicht erlaubt, muss entweder:
- Ein Alias für `kundenservice@` auf dem Mailserver eingerichtet werden
- Oder die FROM-Adresse in `.env` auf `recherche@kalku.de` geändert werden
