require('dotenv').config();
const express = require('express');
const session = require('express-session');
const { google } = require('googleapis');
const pdf = require('pdf-parse');
const ExcelJS = require('exceljs');
const archiver = require('archiver');
const path = require('path');
const fs = require('fs');
const os = require('os');

const app = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(__dirname));
app.use(session({
  secret: process.env.SESSION_SECRET || 'invooro-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

// ── OAuth2 client ─────────────────────────────────────────────────────────────
function createOAuth2Client() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
}

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'https://www.googleapis.com/auth/userinfo.email',
  'https://www.googleapis.com/auth/userinfo.profile'
];

// ── Auth routes ───────────────────────────────────────────────────────────────
app.get('/auth/google', (req, res) => {
  const oauth2Client = createOAuth2Client();
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'
  });
  res.redirect(url);
});

app.get('/auth/google/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.redirect('/app.html?error=access_denied');

  try {
    const oauth2Client = createOAuth2Client();
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client });
    const { data: userInfo } = await oauth2.userinfo.get();

    req.session.tokens = tokens;
    req.session.user = {
      email: userInfo.email,
      name: userInfo.name,
      picture: userInfo.picture
    };

    res.redirect('/app.html?logged_in=true');
  } catch (err) {
    console.error('OAuth callback error:', err);
    res.redirect('/app.html?error=auth_failed');
  }
});

app.get('/auth/logout', (req, res) => {
  req.session.destroy();
  res.json({ success: true });
});

app.get('/api/me', (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: 'Not authenticated' });
  res.json(req.session.user);
});

// ── Gmail scanning ────────────────────────────────────────────────────────────
function requireAuth(req, res, next) {
  if (!req.session.tokens) return res.status(401).json({ error: 'Non connecté' });
  next();
}

// Detect if a message looks like an invoice
function isInvoiceRelated(subject, from, snippet) {
  const keywords = [
    'facture', 'invoice', 'reçu', 'receipt', 'devis', 'avoir',
    'paiement', 'payment', 'commande', 'order', 'abonnement', 'subscription',
    'achat', 'purchase', 'confirmation', 'quittance'
  ];
  const text = `${subject} ${from} ${snippet}`.toLowerCase();
  return keywords.some(kw => text.includes(kw));
}

// Parse invoice data from PDF text using heuristics
function parseInvoiceFromText(text, filename, emailDate, emailFrom) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const fullText = text.toLowerCase();

  // Extract amounts
  const amountRegex = /(\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{2})?)\s*€/g;
  const amounts = [];
  let m;
  while ((m = amountRegex.exec(text)) !== null) {
    const val = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
    if (!isNaN(val) && val > 0) amounts.push(val);
  }
  amounts.sort((a, b) => b - a);

  let ttc = null, ht = null, tva = null;

  // Try to find labeled amounts
  const ttcMatch = text.match(/(?:total\s*(?:ttc|tva\s*comprise|à\s*payer)|montant\s*total)[^\d]*(\d[\d\s,.]*)\s*€/i);
  const htMatch = text.match(/(?:total\s*ht|montant\s*ht|sous[\s-]?total)[^\d]*(\d[\d\s,.]*)\s*€/i);
  const tvaMatch = text.match(/(?:tva|taxe)[^\d]*(\d[\d\s,.]*)\s*€/i);

  if (ttcMatch) ttc = parseFloat(ttcMatch[1].replace(/\s/g, '').replace(',', '.'));
  if (htMatch) ht = parseFloat(htMatch[1].replace(/\s/g, '').replace(',', '.'));
  if (tvaMatch) tva = parseFloat(tvaMatch[1].replace(/\s/g, '').replace(',', '.'));

  // Fallback: largest amount = TTC
  if (!ttc && amounts.length > 0) ttc = amounts[0];
  if (!ht && ttc && tva) ht = ttc - tva;
  if (!tva && ht && ttc) tva = ttc - ht;

  // Extract invoice number
  const numMatch = text.match(/(?:facture|invoice|n[°o]\.?|ref\.?|référence)[^\w]*([A-Z0-9][\w-]{2,20})/i);
  const invoiceNum = numMatch ? numMatch[1] : null;

  // Extract date from PDF text
  const dateMatch = text.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  let invoiceDate = emailDate;
  if (dateMatch) {
    const [, d, mo, y] = dateMatch;
    const year = y.length === 2 ? `20${y}` : y;
    invoiceDate = `${year}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }

  // Extract vendor from email or first lines
  const vendor = emailFrom.replace(/<.*>/, '').replace(/"/g, '').trim() || lines[0] || 'Inconnu';

  return {
    date: invoiceDate,
    fournisseur: vendor,
    numero: invoiceNum || `FAC-${Date.now()}`,
    ht: ht ? Math.round(ht * 100) / 100 : null,
    tva: tva ? Math.round(tva * 100) / 100 : null,
    ttc: ttc ? Math.round(ttc * 100) / 100 : null,
    fichier: filename
  };
}

app.post('/api/scan', requireAuth, async (req, res) => {
  const { dateFrom, dateTo } = req.body;

  try {
    const oauth2Client = createOAuth2Client();
    oauth2Client.setCredentials(req.session.tokens);

    // Refresh token if needed
    oauth2Client.on('tokens', (tokens) => {
      req.session.tokens = { ...req.session.tokens, ...tokens };
    });

    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    // Build Gmail query
    const fromDate = dateFrom ? new Date(dateFrom) : new Date(Date.now() - 90 * 24 * 60 * 60 * 1000);
    const toDate = dateTo ? new Date(dateTo) : new Date();
    const afterEpoch = Math.floor(fromDate.getTime() / 1000);
    const beforeEpoch = Math.floor(toDate.getTime() / 1000) + 86400;

    const query = `has:attachment filename:pdf after:${afterEpoch} before:${beforeEpoch}`;

    const listResponse = await gmail.users.messages.list({
      userId: 'me',
      q: query,
      maxResults: 100
    });

    const messages = listResponse.data.messages || [];
    const invoices = [];
    const errors = [];

    for (const msgRef of messages) {
      try {
        const msg = await gmail.users.messages.get({
          userId: 'me',
          id: msgRef.id,
          format: 'full'
        });

        const headers = msg.data.payload.headers;
        const subject = headers.find(h => h.name === 'Subject')?.value || '';
        const from = headers.find(h => h.name === 'From')?.value || '';
        const dateHeader = headers.find(h => h.name === 'Date')?.value || '';
        const emailDate = dateHeader ? new Date(dateHeader).toISOString().split('T')[0] : new Date().toISOString().split('T')[0];

        if (!isInvoiceRelated(subject, from, msg.data.snippet || '')) continue;

        // Find PDF attachments
        const parts = [];
        function collectParts(part) {
          if (part.filename && part.filename.toLowerCase().endsWith('.pdf')) {
            parts.push(part);
          }
          if (part.parts) part.parts.forEach(collectParts);
        }
        collectParts(msg.data.payload);

        for (const part of parts) {
          try {
            const attId = part.body.attachmentId;
            if (!attId) continue;

            const att = await gmail.users.messages.attachments.get({
              userId: 'me',
              messageId: msgRef.id,
              id: attId
            });

            const data = Buffer.from(att.data.data, 'base64');

            let invoiceData;
            try {
              const parsed = await pdf(data);
              invoiceData = parseInvoiceFromText(parsed.text, part.filename, emailDate, from);
            } catch {
              invoiceData = {
                date: emailDate,
                fournisseur: from.replace(/<.*>/, '').trim(),
                numero: `FAC-${msgRef.id.slice(-8)}`,
                ht: null, tva: null, ttc: null,
                fichier: part.filename
              };
            }

            invoiceData.emailId = msgRef.id;
            invoiceData.attachmentId = attId;
            invoiceData.subject = subject;
            invoices.push(invoiceData);
          } catch (attErr) {
            errors.push(`Pièce jointe ignorée: ${part.filename}`);
          }
        }
      } catch (msgErr) {
        errors.push(`Email ignoré: ${msgRef.id}`);
      }
    }

    // Sort by date desc
    invoices.sort((a, b) => new Date(b.date) - new Date(a.date));

    res.json({ invoices, total: invoices.length, errors });
  } catch (err) {
    console.error('Scan error:', err);
    res.status(500).json({ error: 'Erreur lors du scan: ' + err.message });
  }
});

// ── Excel export ──────────────────────────────────────────────────────────────
app.post('/api/export/excel', requireAuth, async (req, res) => {
  const { invoices, period } = req.body;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Invooro';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Factures', {
    pageSetup: { paperSize: 9, orientation: 'landscape' }
  });

  // Header styling
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  const borderStyle = { style: 'thin', color: { argb: 'FFE5E7EB' } };
  const allBorders = { top: borderStyle, left: borderStyle, bottom: borderStyle, right: borderStyle };

  // Title row
  sheet.mergeCells('A1:H1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = `Invooro — Récapitulatif des factures${period ? ' · ' + period : ''}`;
  titleCell.font = { bold: true, size: 13, color: { argb: 'FF1F2937' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getRow(1).height = 30;

  // Column headers
  const headers = ['Date', 'Fournisseur', 'N° Facture', 'Objet email', 'Montant HT (€)', 'TVA (€)', 'Montant TTC (€)', 'Fichier PDF'];
  const headerRow = sheet.addRow(headers);
  headerRow.eachCell(cell => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = allBorders;
  });
  sheet.getRow(2).height = 22;

  // Column widths
  sheet.columns = [
    { key: 'date', width: 14 },
    { key: 'fournisseur', width: 30 },
    { key: 'numero', width: 18 },
    { key: 'subject', width: 35 },
    { key: 'ht', width: 16 },
    { key: 'tva', width: 12 },
    { key: 'ttc', width: 16 },
    { key: 'fichier', width: 28 }
  ];

  // Data rows
  let totalHT = 0, totalTVA = 0, totalTTC = 0;
  invoices.forEach((inv, i) => {
    const row = sheet.addRow([
      inv.date || '',
      inv.fournisseur || '',
      inv.numero || '',
      inv.subject || '',
      inv.ht != null ? inv.ht : '',
      inv.tva != null ? inv.tva : '',
      inv.ttc != null ? inv.ttc : '',
      inv.fichier || ''
    ]);

    const isEven = i % 2 === 0;
    row.eachCell((cell, colNum) => {
      cell.border = allBorders;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFFFFFFF' : 'FFF9FAFB' } };
      if ([5, 6, 7].includes(colNum)) {
        cell.numFmt = '#,##0.00 "€"';
        cell.alignment = { horizontal: 'right' };
      }
    });

    if (inv.ht) totalHT += inv.ht;
    if (inv.tva) totalTVA += inv.tva;
    if (inv.ttc) totalTTC += inv.ttc;
  });

  // Totals row
  const totalsRow = sheet.addRow(['', 'TOTAL', '', '', totalHT, totalTVA, totalTTC, '']);
  totalsRow.eachCell((cell, colNum) => {
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
    cell.border = allBorders;
    if ([5, 6, 7].includes(colNum)) {
      cell.numFmt = '#,##0.00 "€"';
      cell.alignment = { horizontal: 'right' };
    }
  });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="invooro-factures-${new Date().toISOString().split('T')[0]}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
});

// ── PDF ZIP export ────────────────────────────────────────────────────────────
app.post('/api/export/zip', requireAuth, async (req, res) => {
  const { invoices } = req.body;

  try {
    const oauth2Client = createOAuth2Client();
    oauth2Client.setCredentials(req.session.tokens);
    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="invooro-pdfs-${new Date().toISOString().split('T')[0]}.zip"`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);

    for (const inv of invoices) {
      if (!inv.emailId || !inv.attachmentId) continue;
      try {
        const att = await gmail.users.messages.attachments.get({
          userId: 'me',
          messageId: inv.emailId,
          id: inv.attachmentId
        });
        const data = Buffer.from(att.data.data, 'base64');
        const safeName = inv.fichier.replace(/[^a-z0-9._-]/gi, '_');
        archive.append(data, { name: `${inv.date}_${safeName}` });
      } catch (e) {
        console.error(`Skipping ${inv.fichier}:`, e.message);
      }
    }

    archive.finalize();
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Health check ──────────────────────────────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', version: '1.0.0' });
});

// ── Serve frontend ────────────────────────────────────────────────────────────
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/app', (req, res) => res.sendFile(path.join(__dirname, 'app.html')));

app.listen(PORT, () => {
  console.log(`\n✅ Invooro backend lancé sur http://localhost:${PORT}`);
  console.log(`   → Landing page : http://localhost:${PORT}/`);
  console.log(`   → App          : http://localhost:${PORT}/app.html`);
  console.log(`   → OAuth start  : http://localhost:${PORT}/auth/google\n`);
});
