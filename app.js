/* Note de frais — Option A (browser only)
   - Generates a recap PDF page with PDF-Lib
   - Appends line-linked attachments first (optional), then the rest
   - Supports attachments: PDF, JPG, PNG
*/
const { PDFDocument, StandardFonts, rgb } = window.PDFLib || {};

const state = {
  lineRows: [],
  justifs: [], // { id, file, name, type, size, source: 'list'|'line', lineIndex? }
};

const DEPLACEMENT_TARIF_EUR_KM = 0.3;

const generated = {
  bytes: null,
  filename: null,
  url: null
};

const fmtEUR = (n) => {
  const v = Number.isFinite(n) ? n : 0;
  return v.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const el = (id) => document.getElementById(id);

function setStatus(msg) {
  el('status').textContent = msg || '';
}

function clearGeneratedPreview() {
  if (generated.url) {
    URL.revokeObjectURL(generated.url);
  }
  generated.bytes = null;
  generated.filename = null;
  generated.url = null;

  const preview = el('pdfPreview');
  const wrap = el('pdfPreviewWrap');
  if (preview) preview.src = '';
  if (wrap) wrap.hidden = true;

  const dl = el('downloadBtn');
  if (dl) dl.disabled = true;
}

function setGeneratedPreview(bytes, filename) {
  clearGeneratedPreview();
  generated.bytes = bytes;
  generated.filename = filename;

  const blob = new Blob([bytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  generated.url = url;

  const preview = el('pdfPreview');
  const wrap = el('pdfPreviewWrap');
  if (preview) preview.src = url;
  if (wrap) wrap.hidden = false;

  const dl = el('downloadBtn');
  if (dl) dl.disabled = false;
}

function bytesToHuman(bytes) {
  const u = ['B','KB','MB','GB'];
  let i=0, b=bytes;
  while (b>=1024 && i<u.length-1){ b/=1024; i++; }
  return `${b.toFixed(i===0?0:1)} ${u[i]}`;
}

function mkId() {
  return Math.random().toString(16).slice(2) + Date.now().toString(16);
}

/* ---------- Table rows ---------- */
function addLineRow(prefill = {}) {
  const tbody = el('fraisBody');
  const tr = document.createElement('tr');

  const lineIndex = state.lineRows.length;

  tr.innerHTML = `
    <td>
      <select class="cat">
        <option value="Déplacement">Déplacement</option>
        <option value="Péage">Péage</option>
        <option value="Repas">Repas</option>
        <option value="Hôtel">Hôtel</option>
        <option value="Parking">Parking</option>
        <option value="Divers">Divers</option>
      </select>
    </td>
    <td><input type="text" class="desc" placeholder="Ex: Rouen → Le Havre" value="${prefill.desc || ''}"></td>
    <td class="num"><input type="number" step="1" min="0" class="km" placeholder="0" value="${prefill.km || ''}"></td>
    <td class="num"><input type="number" step="0.01" min="0" class="amt" placeholder="0.00" value="${prefill.amt || ''}"></td>
    <td>
      <input type="file" class="lineJustif" accept=".pdf,image/*" />
      <div class="small muted lineJustifName"></div>
    </td>
    <td class="right"><button class="btn btn-danger delRow">Suppr</button></td>
  `;

  // set category
  tr.querySelector('.cat').value = prefill.cat || 'Déplacement';

  // Events
  const kmEl = tr.querySelector('.km');
  const amtEl = tr.querySelector('.amt');
  const catEl = tr.querySelector('.cat');

  const applyDeplacementRules = () => {
    const cat = (catEl.value || '').trim();
    const isDep = cat === 'Déplacement';
    if (isDep) {
      kmEl.disabled = false;
      kmEl.style.display = '';
      const km = parseFloat(kmEl.value || '0') || 0;
      const amt = km * DEPLACEMENT_TARIF_EUR_KM;
      amtEl.value = amt ? amt.toFixed(2) : '';
      amtEl.disabled = true;
    } else {
      // switching away from Déplacement: user must enter amount manually
      if (amtEl.disabled) amtEl.value = '';
      amtEl.disabled = false;
      kmEl.value = '';
      kmEl.disabled = true;
      kmEl.style.display = 'none';
    }
    updateTotal();
  };

  amtEl.addEventListener('input', updateTotal);
  kmEl.addEventListener('input', applyDeplacementRules);
  catEl.addEventListener('change', applyDeplacementRules);
  tr.querySelector('.delRow').addEventListener('click', () => {
    // remove any linked justif for that line
    const row = state.lineRows[lineIndex];
    if (row?.linkedJustifId) {
      state.justifs = state.justifs.filter(j => j.id !== row.linkedJustifId);
    }
    tr.remove();
    state.lineRows[lineIndex] = null;
    updateTotal();
    renderJustifsList();
  });

  tr.querySelector('.lineJustif').addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    const nameEl = tr.querySelector('.lineJustifName');

    // Remove previous linked justif if any
    const row = state.lineRows[lineIndex];
    if (row?.linkedJustifId) {
      state.justifs = state.justifs.filter(j => j.id !== row.linkedJustifId);
      row.linkedJustifId = null;
    }

    if (!file) {
      nameEl.textContent = '';
      renderJustifsList();
      return;
    }

    const id = mkId();
    const j = {
      id,
      file,
      name: file.name,
      type: file.type || (file.name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : ''),
      size: file.size,
      source: 'line',
      lineIndex
    };
    state.justifs.push(j);
    row.linkedJustifId = id;
    nameEl.textContent = file.name;
    renderJustifsList();
  });

  tbody.appendChild(tr);

  state.lineRows.push({
    tr,
    linkedJustifId: null
  });

  applyDeplacementRules();
}

function getLinesData() {
  const rows = [];
  state.lineRows.forEach((r) => {
    if (!r || !r.tr || !r.tr.isConnected) return;
    const cat = r.tr.querySelector('.cat').value || '';
    const desc = r.tr.querySelector('.desc').value || '';
    const km = parseFloat(r.tr.querySelector('.km')?.value || '0') || 0;
    const amt = parseFloat(r.tr.querySelector('.amt').value || '0') || 0;
    rows.push({ cat, desc, km, amt });
  });
  return rows;
}

function updateTotal() {
  const lines = getLinesData();
  const total = lines.reduce((s, l) => s + (Number.isFinite(l.amt) ? l.amt : 0), 0);
  el('totalEuros').textContent = fmtEUR(total);
}

/* ---------- Justifs list ---------- */
function addJustifsFromPicker(files) {
  if (!files || files.length === 0) return;
  for (const file of files) {
    const id = mkId();
    state.justifs.push({
      id,
      file,
      name: file.name,
      type: file.type || (file.name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : ''),
      size: file.size,
      source: 'list'
    });
  }
  renderJustifsList();
}

function renderJustifsList() {
  const list = el('justifsList');
  if (!list) return;
  list.innerHTML = '';

  // Keep stable order: line-linked first by line index, then list-added
  const linked = state.justifs
    .filter(j => j.source === 'line')
    .sort((a,b) => (a.lineIndex ?? 0) - (b.lineIndex ?? 0));
  const added = state.justifs.filter(j => j.source === 'list');

  const all = [...linked, ...added];

  if (all.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'muted small';
    empty.textContent = 'Aucun justificatif ajouté.';
    list.appendChild(empty);
    return;
  }

  for (const j of all) {
    const item = document.createElement('div');
    item.className = 'item';

    const badge = j.source === 'line'
      ? `Ligne #${(j.lineIndex ?? 0) + 1}`
      : 'Ajouté';
    const mime = j.type || 'fichier';

    item.innerHTML = `
      <div class="meta">
        <div class="name">${escapeHtml(j.name)}</div>
        <div class="sub">${escapeHtml(mime)} • ${bytesToHuman(j.size)} • <span class="badge">${badge}</span></div>
      </div>
      <div class="actions">
        <button class="btn btn-ghost preview">Aperçu</button>
        <button class="btn btn-danger remove">Suppr</button>
      </div>
    `;

    item.querySelector('.remove').addEventListener('click', () => {
      // if it's linked to a row, clear row pointer + UI label
      if (j.source === 'line') {
        const row = state.lineRows[j.lineIndex];
        if (row) {
          row.linkedJustifId = null;
          const nameEl = row.tr?.querySelector('.lineJustifName');
          const inputEl = row.tr?.querySelector('.lineJustif');
          if (nameEl) nameEl.textContent = '';
          if (inputEl) inputEl.value = '';
        }
      }
      state.justifs = state.justifs.filter(x => x.id !== j.id);
      renderJustifsList();
    });

    item.querySelector('.preview').addEventListener('click', async () => {
      // open in new tab
      const url = URL.createObjectURL(j.file);
      window.open(url, '_blank', 'noopener');
      // Note: browser will revoke on refresh; OK.
    });

    list.appendChild(item);
  }
}

function escapeHtml(s) {
  return (s ?? '').toString()
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#039;');
}

/* ---------- PDF generation ---------- */
async function createRecapPdfDefault({ nom, lines, total }) {
  const pdfDoc = await PDFDocument.create();
  let page = pdfDoc.addPage([595.28, 841.89]); // A4 in points (approx)

  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const margin = 40;
  const w = page.getWidth();
  const h = page.getHeight();

  // Title
  page.drawText('NOTE DE FRAIS', { x: margin, y: h - margin - 10, size: 20, font: fontBold, color: rgb(0.15,0.2,0.3) });

  // Info box
  const boxTop = h - margin - 50;
  const boxH = 95;
  page.drawRectangle({ x: margin, y: boxTop - boxH, width: w - margin*2, height: boxH, borderColor: rgb(0.75,0.8,0.88), borderWidth: 1 });

  const info = [
    ['Nom / Prénom', nom || '—'],
    ['Généré le', new Date().toLocaleString('fr-FR')]
  ];

  let ix = margin + 12;
  let iy = boxTop - 22;
  const col1 = 150;
  for (const [k, v] of info) {
    page.drawText(k + ' :', { x: ix, y: iy, size: 10, font: fontBold, color: rgb(0.25,0.3,0.4) });
    page.drawText(String(v), { x: ix + col1, y: iy, size: 10, font, color: rgb(0.15,0.2,0.3) });
    iy -= 16;
    if (iy < boxTop - boxH + 18) { // overflow to right half if needed
      ix = margin + 320;
      iy = boxTop - 22;
    }
  }

  // Table header
  let y = boxTop - boxH - 30;
  page.drawText('Détails', { x: margin, y, size: 13, font: fontBold, color: rgb(0.15,0.2,0.3) });
  y -= 18;

  const tableX = margin;
  const tableW = w - margin*2;
  const rowH = 18;

  const cols = [
    { key:'cat', label:'Catégorie', width: 110 },
    { key:'desc', label:'Description', width: tableW - (110+55+55+80) },
    { key:'tarif', label:'Tarif', width: 55, align:'right' },
    { key:'km', label:'Km', width: 55, align:'right' },
    { key:'amt', label:'Montant', width: 80, align:'right' }
  ];

  // Header background
  page.drawRectangle({ x: tableX, y: y - rowH + 4, width: tableW, height: rowH, color: rgb(0.92,0.95,0.99) });
  let cx = tableX + 6;
  for (const c of cols) {
    page.drawText(c.label, { x: cx, y: y, size: 10, font: fontBold, color: rgb(0.2,0.25,0.35) });
    cx += c.width;
  }
  y -= rowH;

  // Rows (auto paginate if too many)
  const maxY = margin + 80;
  for (let i=0;i<lines.length;i++) {
    if (y < maxY) {
      // new page
      const p = pdfDoc.addPage([595.28, 841.89]);
      y = p.getHeight() - margin - 20;
      // repeat header
      p.drawText('NOTE DE FRAIS (suite)', { x: margin, y, size: 13, font: fontBold, color: rgb(0.15,0.2,0.3) });
      y -= 18;
      p.drawRectangle({ x: tableX, y: y - rowH + 4, width: tableW, height: rowH, color: rgb(0.92,0.95,0.99) });
      cx = tableX + 6;
      for (const c of cols) {
        p.drawText(c.label, { x: cx, y: y, size: 10, font: fontBold, color: rgb(0.2,0.25,0.35) });
        cx += c.width;
      }
      y -= rowH;
      // switch current page reference
      page = p;
    }

    const line = lines[i];
    const p = pdfDoc.getPages()[pdfDoc.getPageCount()-1];

    // row line
    p.drawLine({ start: {x: tableX, y: y+4}, end: {x: tableX+tableW, y: y+4}, thickness: 1, color: rgb(0.92,0.93,0.96) });

    cx = tableX + 6;
    const isDep = ((line.cat || '').trim() === 'Déplacement');
    const tarif = isDep ? '0,30' : '';
    const kms = isDep ? String(Math.round(line.km || 0)) : '';
    const cells = [
      line.cat || '',
      line.desc || '',
      tarif,
      kms,
      fmtEUR(line.amt || 0)
    ];
    for (let j=0;j<cols.length;j++) {
      const c = cols[j];
      const text = String(cells[j] ?? '');
      if (c.align === 'right') {
        const tw = font.widthOfTextAtSize(text, 10);
        p.drawText(text, { x: cx + c.width - 6 - tw, y, size: 10, font, color: rgb(0.15,0.2,0.3) });
      } else {
        // naive wrap for desc
        if (c.key === 'desc' && text.length > 55) {
          p.drawText(text.slice(0, 55) + '…', { x: cx, y, size: 10, font, color: rgb(0.15,0.2,0.3) });
        } else {
          p.drawText(text, { x: cx, y, size: 10, font, color: rgb(0.15,0.2,0.3) });
        }
      }
      cx += c.width;
    }

    y -= rowH;
  }

  // Total box (on last page)
  const pages = pdfDoc.getPages();
  const last = pages[pages.length - 1];
  const ty = Math.max(margin + 30, y - 10);
  last.drawRectangle({ x: tableX, y: ty, width: tableW, height: 32, borderColor: rgb(0.75,0.8,0.88), borderWidth: 1, color: rgb(0.98,0.99,1) });
  last.drawText('TOTAL', { x: tableX + 10, y: ty + 10, size: 12, font: fontBold, color: rgb(0.15,0.2,0.3) });
  const totalStr = fmtEUR(total) + ' €';
  const tw = fontBold.widthOfTextAtSize(totalStr, 14);
  last.drawText(totalStr, { x: tableX + tableW - 10 - tw, y: ty + 8, size: 14, font: fontBold, color: rgb(0.1,0.15,0.25) });

  return pdfDoc;
}

async function createRecapPdf({ nom, adresse, motif, lieu, dateMission, renonceIndemnites, lines, total }) {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([595.28, 841.89]); // A4 portrait

  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const w = page.getWidth();
  const h = page.getHeight();

  const safe = (s) => (s ?? '').toString().trim();
  const black = rgb(0, 0, 0);
  const gray = rgb(0.25, 0.25, 0.25);
  const green = rgb(0.12, 0.45, 0.18);
  const red = rgb(0.82, 0.05, 0.05);

  const wrapText = (text, maxWidth, size, usedFont) => {
    const t = safe(text);
    if (!t) return [''];
    const words = t.split(/\s+/).filter(Boolean);
    const out = [];
    let cur = '';
    for (const word of words) {
      const cand = cur ? `${cur} ${word}` : word;
      if (usedFont.widthOfTextAtSize(cand, size) <= maxWidth) {
        cur = cand;
      } else {
        if (cur) out.push(cur);
        cur = word;
      }
    }
    if (cur) out.push(cur);
    return out.length ? out : [''];
  };

  const centerText = (text, y, size, usedFont, color = black) => {
    const tw = usedFont.widthOfTextAtSize(text, size);
    page.drawText(text, { x: (w - tw) / 2, y, size, font: usedFont, color });
  };

  const drawStamp = (text, yCenter) => {
    const size = 11;
    const paddingX = 14;
    const paddingY = 10;
    const tw = fontBold.widthOfTextAtSize(text, size);
    const boxW = tw + paddingX * 2;
    const boxH = size + paddingY * 2;
    const x = (w - boxW) / 2;
    const y = yCenter - boxH / 2;
    page.drawRectangle({ x, y, width: boxW, height: boxH, borderColor: red, borderWidth: 2, color: rgb(1, 1, 1), opacity: 0.0 });
    page.drawText(text, { x: x + paddingX, y: y + paddingY, size, font: fontBold, color: red });
  };

  const missionStr = (() => {
    if (!dateMission) return '';
    try {
      const d = new Date(dateMission);
      return Number.isNaN(d.getTime()) ? safe(dateMission) : d.toLocaleDateString('fr-FR');
    } catch {
      return safe(dateMission);
    }
  })();

  const nowStr = new Date().toLocaleDateString('fr-FR');

  const toUint8 = (ab) => new Uint8Array(ab);

  const svgToPngBytes = async (svgText) => {
    const svgBlob = new Blob([svgText], { type: 'image/svg+xml' });
    const svgUrl = URL.createObjectURL(svgBlob);
    try {
      const img = new Image();
      const loaded = new Promise((resolve, reject) => {
        img.onload = () => resolve();
        img.onerror = (e) => reject(e);
      });
      img.src = svgUrl;
      await loaded;

      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, img.naturalWidth || 600);
      canvas.height = Math.max(1, img.naturalHeight || 200);
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);

      const dataUrl = canvas.toDataURL('image/png');
      const b64 = dataUrl.split(',')[1] || '';
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
      return bytes;
    } finally {
      URL.revokeObjectURL(svgUrl);
    }
  };

  const loadEmbedImage = async (url) => {
    if (window.location?.protocol === 'file:') {
      throw new Error('Logo loading disabled on file://');
    }
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`fetch ${url} failed: ${res.status}`);
    if (url.toLowerCase().endsWith('.svg')) {
      const svgText = await res.text();
      const pngBytes = await svgToPngBytes(svgText);
      return pdfDoc.embedPng(pngBytes);
    }
    const bytes = await res.arrayBuffer();
    const u8 = toUint8(bytes);
    if (url.toLowerCase().endsWith('.png')) return pdfDoc.embedPng(u8);
    return pdfDoc.embedJpg(u8);
  };

  // --- Header ---
  const logoY = h - 90;
  const logoH = 55;
  const logoW = 110;
  const logoGap = 18;
  const slots = [];
  let lx = 40;
  for (let i=0;i<4;i++) {
    slots.push({ x: lx, y: logoY, w: logoW, h: logoH });
    lx += logoW + logoGap;
  }

  const logoCandidates = [
    './logos/logo_RBFC.svg',
    './logos/Logo_FFCO.svg',
    './logos/Logo FFCO.png',
    './logos/Logo ANS.png',
  ];

  const embedded = [];
  for (const url of logoCandidates) {
    try {
      const img = await loadEmbedImage(url);
      embedded.push(img);
    } catch {
      // ignore
    }
  }

  for (let i=0;i<slots.length;i++) {
    const s = slots[i];
    const img = embedded[i];
    if (!img) {
      page.drawRectangle({ x: s.x, y: s.y, width: s.w, height: s.h, borderColor: rgb(0.85,0.85,0.85), borderWidth: 1 });
      continue;
    }
    const dims = img.scale(1);
    const scale = Math.min(s.w / dims.width, s.h / dims.height);
    const dw = dims.width * scale;
    const dh = dims.height * scale;
    const x = s.x + (s.w - dw) / 2;
    const y = s.y + (s.h - dh) / 2;
    page.drawImage(img, { x, y, width: dw, height: dh });
  }

  // President / email (center)
  centerText('Président : Francis MOINE 8, les drillons 89570 BEUGNON', h - 115, 9, fontBold, gray);
  centerText('Émail : ligue.bourgogne-franche-comte@ffcorientation.fr', h - 128, 9, fontBold, gray);

  // Treasurer block (left)
  page.drawText('Trésorière : Clotilde PERRIN', { x: 55, y: h - 160, size: 9, font: fontBold, color: gray });
  const treLines = [
    '24 rue des Justices Bât 35 Appart 181',
    '25000 Besançon',
    'Tel : 06 80 54 53 40'
  ];
  let ty = h - 172;
  for (const t of treLines) {
    page.drawText(t, { x: 55, y: ty, size: 9, font, color: gray });
    ty -= 12;
  }

  // --- Titles (moved up to free space for the table) ---
  centerText('ANNEXE I', h - 215, 14, fontBold, black);
  centerText('État de remboursement des frais de déplacement pour mission', h - 240, 14, fontBold, black);

  // --- Info block (moved up) ---
  const infoX = 70;
  let y = h - 285;
  const labelSize = 10.5;
  const valueSize = 10.5;
  const valueX = 260;
  const maxValueW = w - valueX - 60;

  const drawInfo = (label, value) => {
    page.drawText(label, { x: infoX, y, size: labelSize, font: fontBold, color: black });
    const linesW = wrapText(value || '—', maxValueW, valueSize, font);
    page.drawText(linesW[0], { x: valueX, y, size: valueSize, font, color: black });
    y -= 16;
    if (linesW.length > 1) {
      page.drawText(linesW[1], { x: valueX, y, size: valueSize, font, color: black });
      y -= 16;
    }
  };

  drawInfo('NOM - PRÉNOM et ADRESSE :', [safe(nom), safe(adresse)].filter(Boolean).join(', '));
  drawInfo('Motif du déplacement :', safe(motif));
  drawInfo('Lieu du déplacement :', safe(lieu));
  drawInfo('Date de la mission :', missionStr || nowStr);

  // --- Green rules block ---
  y -= 10;
  const rules = [
    'Utiliser un seul état par déplacement ou mission à envoyer par courrier au trésorier (email :',
    'compta.bfco@gmail.com) dans le mois qui suit la mission (accompagné d\'un RIB).',
    'Joindre les justificatifs (invitation - péage - etc...).',
    '- L\'indemnité est fixée chaque année par le Comité Directeur (0,30€/km). Elle peut être modifiée par anticipation',
    'lors d\'un Comité Directeur si le besoin s\'en fait sentir.',
    '- Frais de péage sur justificatif.',
    '- Frais de restauration : remboursement sur justificatif plafonné à 18,00 € par repas.',
    '- Frais d\'hébergement : remboursement sur facture plafonné à 40,00 € par nuitée et 8,00€ pour le petit déjeuner.'
  ];
  const ruleSize = 8.5;
  for (const r of rules) {
    const rl = wrapText(r, w - 2*infoX, ruleSize, font);
    for (const line of rl) {
      page.drawText(line, { x: infoX, y, size: ruleSize, font, color: green });
      y -= 11;
    }
  }

  const drawTableHeader = (p, startY) => {
    const headerTopY = startY;
    p.drawRectangle({ x: tableX, y: headerTopY - headerH, width: tableW, height: headerH, borderColor: gray, borderWidth: 1 });
    let cx2 = tableX;
    for (const c of cols) {
      p.drawLine({ start: { x: cx2, y: headerTopY }, end: { x: cx2, y: headerTopY - headerH }, thickness: 1, color: gray });
      p.drawText(c.label, { x: cx2 + 6, y: headerTopY - 15, size: 9.5, font: fontBold, color: black });
      cx2 += c.w;
    }
    p.drawLine({ start: { x: tableX + tableW, y: headerTopY }, end: { x: tableX + tableW, y: headerTopY - headerH }, thickness: 1, color: gray });
    return headerTopY - headerH;
  };

  const drawTableRow = (p, rowTopY, l) => {
    p.drawRectangle({ x: tableX, y: rowTopY - rowH, width: tableW, height: rowH, borderColor: gray, borderWidth: 1 });
    let cx2 = tableX;
    const cat = l ? safe(l.cat) : '';
    const desc = l ? safe(l.desc) : '';
    const isDep = l && (safe(l.cat) === 'Déplacement');
    const tarif = isDep ? '0,30€' : '';
    const kms = isDep ? String(Math.round(l.km || 0)) : '';
    const montant = l ? `${fmtEUR(l.amt || 0)}€` : '';
    const cells = [cat, desc, tarif, kms, montant];

    for (let j=0;j<cols.length;j++) {
      const c = cols[j];
      p.drawLine({ start: { x: cx2, y: rowTopY }, end: { x: cx2, y: rowTopY - rowH }, thickness: 1, color: gray });
      const txt = safe(cells[j]);
      if (c.align === 'right') {
        const tw = font.widthOfTextAtSize(txt, 10);
        p.drawText(txt, { x: cx2 + c.w - 6 - tw, y: rowTopY - 14, size: 10, font, color: black });
      } else {
        const t = wrapText(txt, c.w - 12, 10, font)[0];
        p.drawText(t, { x: cx2 + 6, y: rowTopY - 14, size: 10, font, color: black });
      }
      cx2 += c.w;
    }
    p.drawLine({ start: { x: tableX + tableW, y: rowTopY }, end: { x: tableX + tableW, y: rowTopY - rowH }, thickness: 1, color: gray });
    return rowTopY - rowH;
  };

  // --- Table ---
  y -= 12;
  const tableX = 85;
  const tableW = w - tableX*2;
  const headerH = 22;
  const rowH = 20;
  const cols = [
    { key: 'cat', label: 'Catégorie', w: tableW * 0.20 },
    { key: 'desc', label: 'Description', w: tableW * 0.32 },
    { key: 'tarif', label: 'Tarif', w: tableW * 0.16 },
    { key: 'km', label: 'Kilomètres', w: tableW * 0.16 },
    { key: 'amt', label: 'MONTANT', w: tableW * 0.16, align: 'right' }
  ];
  y = drawTableHeader(page, y);

  // Fit as many rows as possible on the first page without overlapping the bottom blocks
  const minYAfterTable = 250;
  const maxRowsFirstPage = Math.max(0, Math.floor((y - minYAfterTable) / rowH));
  const safeLines = Array.isArray(lines) ? lines : [];
  const firstLines = safeLines.slice(0, Math.min(maxRowsFirstPage, safeLines.length));
  let remaining = safeLines.slice(firstLines.length);

  for (const l of firstLines) {
    y = drawTableRow(page, y, l);
  }

  // --- Total box ---
  y -= 18;
  const totalBoxW = 320;
  const totalBoxH = 34;
  const totalX = (w - totalBoxW) / 2;
  page.drawRectangle({ x: totalX, y: y - totalBoxH, width: totalBoxW, height: totalBoxH, borderColor: gray, borderWidth: 1 });
  page.drawText('Total général :', { x: totalX + 14, y: y - 22, size: 12, font: fontBold, color: black });
  const totalStr = `€ ${fmtEUR(total)}`;
  const tw = fontBold.widthOfTextAtSize(totalStr, 12);
  page.drawText(totalStr, { x: totalX + totalBoxW - 14 - tw, y: y - 22, size: 12, font: fontBold, color: black });

  // --- Certification (red) ---
  y -= 50;
  centerText('Je certifie ne pas me faire rembourser mes frais plusieurs fois.', y, 10.5, fontBold, red);

  if (renonceIndemnites) {
    drawStamp('Je renonce au paiement de ces indemnités', y - 22);
  }

  // --- Signature boxes ---
  const boxY = 55;
  const boxH = 78;
  const boxW = 210;
  const leftX = 70;
  const rightX = w - 70 - boxW;
  page.drawRectangle({ x: leftX, y: boxY, width: boxW, height: boxH, borderColor: gray, borderWidth: 1 });
  page.drawRectangle({ x: rightX, y: boxY, width: boxW, height: boxH, borderColor: gray, borderWidth: 1 });
  page.drawText('Nom prénom et Date', { x: leftX + 40, y: boxY + boxH - 20, size: 10, font: fontBold, color: black });
  page.drawText(`${safe(nom) || '—'}, ${nowStr}`, { x: leftX + 16, y: boxY + boxH - 36, size: 10, font, color: black });
  page.drawText('Le trésorier :', { x: rightX + 14, y: boxY + boxH - 22, size: 10, font: fontBold, color: black });

  // Continuation pages: table only, until all lines are rendered (1:1 with web table)
  while (remaining.length > 0) {
    const p = pdfDoc.addPage([595.28, 841.89]);
    const pw = p.getWidth();
    const ph = p.getHeight();
    const title = 'ANNEXE I (suite)';
    const tw = fontBold.widthOfTextAtSize(title, 12);
    p.drawText(title, { x: (pw - tw) / 2, y: ph - 50, size: 12, font: fontBold, color: black });

    let py = ph - 75;
    py = drawTableHeader(p, py);

    const minY = 50;
    const rowsFit = Math.max(0, Math.floor((py - minY) / rowH));
    const chunk = remaining.slice(0, rowsFit);
    remaining = remaining.slice(chunk.length);
    for (const l of chunk) {
      py = drawTableRow(p, py, l);
    }

    // safety: prevent infinite loop if something goes wrong with geometry
    if (chunk.length === 0) break;
  }

  return pdfDoc;
}

async function appendPdf(finalDoc, file) {
  const bytes = await file.arrayBuffer();
  const donor = await PDFDocument.load(bytes);
  const donorPages = await finalDoc.copyPages(donor, donor.getPageIndices());
  donorPages.forEach(p => finalDoc.addPage(p));
}

async function appendImageAsPage(finalDoc, file) {
  const bytes = await file.arrayBuffer();
  const isPng = (file.type === 'image/png') || file.name.toLowerCase().endsWith('.png');
  const img = isPng ? await finalDoc.embedPng(bytes) : await finalDoc.embedJpg(bytes);

  const page = finalDoc.addPage([595.28, 841.89]); // A4 portrait
  const margin = 24;
  const maxW = page.getWidth() - margin*2;
  const maxH = page.getHeight() - margin*2;

  const dims = img.scale(1);
  let scale = Math.min(maxW / dims.width, maxH / dims.height);
  // avoid upscaling too much
  scale = Math.min(scale, 1.2);

  const w = dims.width * scale;
  const h = dims.height * scale;

  const x = (page.getWidth() - w) / 2;
  const y = (page.getHeight() - h) / 2;

  page.drawImage(img, { x, y, width: w, height: h });
}

function isPdf(file) {
  return (file.type === 'application/pdf') || file.name.toLowerCase().endsWith('.pdf');
}

function isImage(file) {
  const t = file.type || '';
  return t.startsWith('image/') || /\.(png|jpe?g)$/i.test(file.name);
}

async function generateFinalPdf() {
  if (!window.PDFLib) {
    alert('PDF-Lib n\'est pas chargé. Vérifie ta connexion internet (CDN unpkg).');
    return;
  }

  setStatus('Préparation…');
  clearGeneratedPreview();

  const nom = el('nom').value.trim();
  const adresse = (el('adresse')?.value || '').trim();
  const motif = (el('motif')?.value || '').trim();
  const lieu = (el('lieu')?.value || '').trim();
  const dateMission = (el('dateMission')?.value || '').trim();
  const renonceIndemnites = !!el('renonceIndemnites')?.checked;

  const lines = getLinesData();
  const total = lines.reduce((s, l) => s + (l.amt || 0), 0);

  // Build recap
  setStatus('Génération du récap PDF…');
  const recapDoc = await createRecapPdf({ nom, adresse, motif, lieu, dateMission, renonceIndemnites, lines, total });

  // Create final doc + copy recap pages
  const finalDoc = await PDFDocument.create();
  const recapPages = await finalDoc.copyPages(recapDoc, recapDoc.getPageIndices());
  recapPages.forEach(p => finalDoc.addPage(p));

  // Collect attachments
  const linked = state.justifs
    .filter(j => j.source === 'line')
    .sort((a,b) => (a.lineIndex ?? 0) - (b.lineIndex ?? 0));

  const added = state.justifs.filter(j => j.source === 'list');

  const ordered = [...linked, ...added];

  if (ordered.length > 0) {
    setStatus('Ajout des justificatifs…');
  }

  for (let i=0;i<ordered.length;i++) {
    const j = ordered[i];
    setStatus(`Ajout justificatif ${i+1}/${ordered.length} : ${j.name}`);
    if (isPdf(j.file)) {
      await appendPdf(finalDoc, j.file);
    } else if (isImage(j.file)) {
      await appendImageAsPage(finalDoc, j.file);
    } else {
      console.warn('Format non supporté:', j.file);
    }
  }

  setStatus('Finalisation…');
  const bytes = await finalDoc.save();

  const missionDate = (dateMission || '').toString().trim();
  const baseParts = [
    (nom || '').toString().trim(),
    missionDate
  ].filter(Boolean);
  const rawBase = baseParts.length ? baseParts.join('_') : 'note_de_frais';
  const filenameBase = rawBase.replace(/[^a-z0-9_-]+/gi, '_');
  const outName = `${filenameBase}_complet.pdf`;

  setGeneratedPreview(bytes, outName);
  setStatus(`✅ PDF prêt : ${outName}`);
}

function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/* ---------- init ---------- */
function resetAll() {
  // clear table
  el('fraisBody').innerHTML = '';
  state.lineRows = [];
  // clear justifs
  state.justifs = [];
  const justifsInput = el('justifsInput');
  if (justifsInput) justifsInput.value = '';
  clearGeneratedPreview();
  renderJustifsList();
  updateTotal();
  setStatus('');
}

window.addEventListener('DOMContentLoaded', () => {
  // default rows
  addLineRow({ cat: 'Déplacement' });
  addLineRow({ cat: 'Repas' });

  el('addRowBtn').addEventListener('click', () => addLineRow());
  el('clearBtn').addEventListener('click', resetAll);

  const addJustifsBtn = el('addJustifsBtn');
  const justifsInput = el('justifsInput');
  if (addJustifsBtn && justifsInput) {
    addJustifsBtn.addEventListener('click', () => {
      const files = justifsInput.files;
      addJustifsFromPicker(files);
      justifsInput.value = '';
    });
  }

  el('generateBtn').addEventListener('click', () => {
    generateFinalPdf().catch(err => {
      console.error(err);
      setStatus('❌ Erreur pendant la génération. Ouvre la console (F12) pour le détail.');
      alert('Erreur pendant la génération du PDF. Détail dans la console (F12).');
    });
  });

  const downloadBtn = el('downloadBtn');
  if (downloadBtn) {
    downloadBtn.addEventListener('click', () => {
      if (!generated.bytes || !generated.filename) return;
      downloadBytes(generated.bytes, generated.filename);
    });
  }

  renderJustifsList();
  updateTotal();
});
