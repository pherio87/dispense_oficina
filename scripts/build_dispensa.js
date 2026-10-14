#!/usr/bin/env node
// build_dispensa.js — da Markdown a .docx con il layout standard Oficina.
//
// Uso:  node build_dispensa.js dispensa.md [-o dispensa.docx] [--brand brand.json]
//
// Il testo si scrive in Markdown (vedi references/sintassi.md); pandoc lo legge e questo script
// costruisce il documento con docx-js. Tutte le misure del layout sono in questo file e in
// ../assets/brand.json (colori), così ogni dispensa esce uguale.
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const JSZip = require(require.resolve('jszip', { paths: [path.dirname(require.resolve('docx')), __dirname] }));
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType, PageNumber, TabStopType, Tab,
  ImageRun, VerticalAlign, LineRuleType, HeightRule, SimpleField, ExternalHyperlink,
  VerticalAlignSection, PageBreak, TableLayoutType,
} = require('docx');

// ---------------------------------------------------------------- argomenti
const argv = process.argv.slice(2);
const opt = (name, def) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : def; };
const WITH_VALUE = ['-o', '--brand', '--qualifiche'];
const input = argv.find((a, i) => !a.startsWith('-') && !WITH_VALUE.includes(argv[i - 1]));
if (!input) { console.error('Uso: node build_dispensa.js dispensa.md [-o out.docx] [--brand brand.json] [--qualifiche cartella]'); process.exit(2); }
const OUT = opt('-o', input.replace(/\.md$/i, '') + '.docx');
const BRAND_FILE = opt('--brand', path.join(__dirname, '..', 'assets', 'brand.json'));
const QUAL_DIR = opt('--qualifiche', path.join(__dirname, '..', 'references', 'qualifiche'));
const BASE_DIR = path.dirname(path.resolve(input));

// ---------------------------------------------------------------- geometria (DXA: 567 = 1 cm)
const PAGE_W = 11906, PAGE_H = 16838;        // A4
const MARGIN = 1134;                          // 2 cm su tutti i lati: spazio per il logo, le note e i fori
const HF_DIST = 709;                          // 1,25 cm: intestazione e piè di pagina dal bordo (come Word)
const TEXT_W = PAGE_W - 2 * MARGIN;           // 9638 = 17 cm
const TEXT_H = PAGE_H - 2 * MARGIN;           // 14570 = 25,7 cm
const CM = 567, PX_PER_CM = 96 / 2.54;
const MAX_IMG_H_CM = 12;                      // un'immagine + didascalia deve stare in pagina con un po' di testo
const SPLIT_ABOVE_CM = 23;                    // oltre questa altezza stimata un esercizio può spezzarsi

// ---------------------------------------------------------------- colori (solo da brand.json)
const brand = JSON.parse(fs.readFileSync(BRAND_FILE, 'utf8'));
const hex2rgb = (h) => [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16));
const rgb2hex = (a) => a.map(v => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
const mix = (h, other, t) => rgb2hex(hex2rgb(h).map((v, i) => v * (1 - t) + hex2rgb(other)[i] * t));
const lum = (h) => { const [r, g, b] = hex2rgb(h).map(v => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); }); return 0.2126 * r + 0.7152 * g + 0.0722 * b; };
const contrast = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
function readable(h, bg = 'FFFFFF') {          // il testo colorato deve restare leggibile (WCAG 4,5:1)
  let c = h, n = 0;
  while (contrast(c, bg) < 4.5 && n++ < 20) c = mix(c, '000000', 0.08);
  if (c !== h) console.warn(`ATTENZIONE: ${h} ha contrasto ${contrast(h, bg).toFixed(1)} su ${bg}: uso ${c}`);
  return c;
}
const INK = readable(brand.primary || '000000');            // testi, titoli, filetti
const ACCENT = brand.accent || 'E30613';                     // accento (rosso del logo)
const ON_ACCENT = contrast('FFFFFF', ACCENT) >= contrast('000000', ACCENT) ? 'FFFFFF' : '000000';
const ON_INK = contrast('FFFFFF', INK) >= contrast('000000', INK) ? 'FFFFFF' : '000000';
// Blu: titoli di capitolo e sezione, intestazioni di tabella, riquadro "Da ricordare" (informare).
// Il rosso resta per ciò che segnala: filetto delle unità, numeri degli elenchi, riquadro "Attenzione".
const SECOND = readable(brand.secondary || brand.primary || '000000');
const ON_SECOND = contrast('FFFFFF', SECOND) >= contrast('000000', SECOND) ? 'FFFFFF' : '000000';

// ---------------------------------------------------------------- scala tipografica (mezzi punti)
// Un solo carattere (Arial) e tre misure: 12 pt (testo, elenchi, tabelle, didascalie), 16 pt (intestazioni), 20 pt (titoli e unità).
const T = { font: 'Arial', body: 24, h1: 40, h2: 32, h3: 32, small: 24, tbl: 24, comp: 20, line: 360, after: 200 };
// comp: 10 pt, solo per capacità e conoscenze nel riquadro Competenze (testi lunghi del repertorio, non testo di studio)
const auto = (n) => ({ line: n, lineRule: LineRuleType.AUTO });
const line = (color, size = 4, space = 1) => ({ style: BorderStyle.SINGLE, size, space, color });

// ---------------------------------------------------------------- stato e utilità
const ctx = { fig: 0, tab: 0, ex: 0, list: 0, warnings: [],
  comp: { qual: null, unit: null, units: [], used: new Map() },   // riquadri Competenze: qualifica corrente, unità, codici usati
 stats: { units: 0, chapters: 0, figures: 0, tables: 0, exercises: 0, splitEx: 0, images: 0 } };
const warn = (m) => { ctx.warnings.push(m); };

// ---------------------------------------------------------------- qualifiche RER (references/qualifiche/*.md)
// Ogni file: "# Titolo", "## UCn – titolo", "### Capacità" / "### Conoscenze", righe "| n.m | testo |".
// Nella dispensa si scrivono solo i codici: il testo viene sempre da qui, così non si ricopia a mano.
const norm = (s) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
let QUALS = null;
function parseQualifica(file) {
  const q = { file: path.basename(file), title: '', ucs: {}, codes: {} };
  let uc = null, kind = null;
  for (const raw of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
    const l = raw.trim(); let m;
    if (!q.title && (m = l.match(/^#\s+(.+)$/))) q.title = m[1].trim();
    else if ((m = l.match(/^##\s+UC\s*(\d+)\s*[–—-]\s*(.+)$/i))) { uc = { n: m[1], title: m[2].trim(), cap: [], con: [] }; q.ucs[uc.n] = uc; kind = null; }
    else if (/^##\s/.test(l)) { uc = null; kind = null; }
    else if (/^###\s+Capacit/i.test(l)) kind = 'cap';
    else if (/^###\s+Conoscenz/i.test(l)) kind = 'con';
    else if (uc && kind && (m = l.match(/^\|\s*(\d+\.\d+)\s*\|\s*(.+?)\s*\|$/))) {
      const it = { code: m[1], text: m[2], kind, uc: uc.n };
      if (m[1].split('.')[0] !== uc.n) warn(`${q.file}: il codice ${m[1]} è sotto UC${uc.n}`);
      if (q.codes[m[1]]) warn(`${q.file}: codice ${m[1]} ripetuto`);
      uc[kind].push(it); q.codes[m[1]] = it;
    }
  }
  return q;
}
function loadQualifiche() {
  if (QUALS) return QUALS;
  QUALS = fs.existsSync(QUAL_DIR) ? fs.readdirSync(QUAL_DIR).filter(f => /\.md$/i.test(f)).sort().map(f => parseQualifica(path.join(QUAL_DIR, f))) : [];
  if (!QUALS.length) warn(`Nessuna qualifica trovata in ${QUAL_DIR}`);
  return QUALS;
}
// "04", "4", "impianti elettrici", "Operatore impianti elettrici": numero del file oppure parte del titolo
function findQualifica(id) {
  const qs = loadQualifiche(), s = String(id).trim();
  if (/^\d+$/.test(s)) return qs.filter(q => { const m = q.file.match(/^(\d+)/); return m && +m[1] === +s; });
  const n = norm(s), exact = qs.filter(q => norm(q.title) === n);
  return exact.length ? exact : qs.filter(q => norm(q.title).includes(n) || norm(q.file).includes(n));
}
// "1.2 1.4-1.6, UC3" -> ['1.2','1.4','1.5','1.6','UC3']
function parseCodes(s) {
  const out = [];
  for (const tok of String(s || '').split(/[\s,;]+/).filter(Boolean)) {
    const r = tok.match(/^(\d+)\.(\d+)[–—-](?:(\d+)\.)?(\d+)$/);
    if (r && (!r[3] || r[3] === r[1])) { for (let k = +r[2]; k <= +r[4]; k++) out.push(`${r[1]}.${k}`); }
    else out.push(/^uc\d+$/i.test(tok) ? tok.toUpperCase() : tok);
  }
  return out;
}

function imageSize(buf) {
  if (buf.slice(1, 4).toString() === 'PNG') return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  if (buf[0] === 0xff && buf[1] === 0xd8) {
    let i = 2;
    while (i < buf.length) {
      if (buf[i] !== 0xff) { i++; continue; }
      const m = buf[i + 1];
      if (m >= 0xc0 && m <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(m)) return { h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
      i += 2 + buf.readUInt16BE(i + 2);
    }
  }
  throw new Error('formato immagine non riconosciuto (usa PNG o JPG)');
}
const imgType = (buf) => (buf.slice(1, 4).toString() === 'PNG' ? 'png' : 'jpg');

// ---------------------------------------------------------------- inline pandoc -> segmenti -> TextRun
const kvOf = (attr) => Object.fromEntries((attr && attr[2]) || []);
function flat(list, props = {}, out = []) {
  for (const n of list || []) {
    switch (n.t) {
      case 'Str': out.push({ k: 'text', text: n.c, props }); break;
      case 'Space': case 'SoftBreak': out.push({ k: 'text', text: ' ', props }); break;
      case 'LineBreak': out.push({ k: 'br' }); break;
      case 'Strong': flat(n.c, { ...props, bold: true }, out); break;
      case 'Emph': flat(n.c, { ...props, italics: true }, out); break;
      case 'Underline': flat(n.c, { ...props, underline: {} }, out); break;
      case 'Strikeout': flat(n.c, { ...props, strike: true }, out); break;
      case 'Superscript': flat(n.c, { ...props, superScript: true }, out); break;
      case 'Subscript': flat(n.c, { ...props, subScript: true }, out); break;
      case 'SmallCaps': case 'Span': flat(n.t === 'Span' ? n.c[1] : n.c, props, out); break;
      case 'Quoted': { const [q, c] = n.c; const [a, b] = q.t === 'SingleQuote' ? ['‘', '’'] : ['“', '”'];
        out.push({ k: 'text', text: a, props }); flat(c, props, out); out.push({ k: 'text', text: b, props }); break; }
      case 'Code': out.push({ k: 'text', text: n.c[1], props }); break;
      case 'Math': out.push({ k: 'text', text: n.c[1], props }); break;
      case 'Cite': flat(n.c[1], props, out); break;
      case 'Link': out.push({ k: 'link', url: n.c[2][0], inner: flat(n.c[1], { ...props, underline: {} }, []) }); break;
      case 'Image': out.push({ k: 'img', node: n }); break;
      case 'Note': warn('Nota a piè di pagina ignorata (non prevista nel layout)'); break;
      case 'RawInline': break;
      default: warn(`Elemento inline non gestito: ${n.t}`);
    }
  }
  return out;
}
function toRuns(items, base = {}) {
  const merged = [];
  for (const it of items) {
    const last = merged[merged.length - 1];
    if (it.k === 'text' && last && last.k === 'text' && JSON.stringify(last.props) === JSON.stringify(it.props)) last.text += it.text;
    else merged.push({ ...it });
  }
  return merged.map(it => {
    if (it.k === 'br') return new TextRun({ ...base, break: 1 });
    if (it.k === 'link') return new ExternalHyperlink({ link: it.url, children: toRuns(it.inner, base) });
    if (it.k === 'img') return null;
    return new TextRun({ ...base, ...it.props, text: it.text });
  }).filter(Boolean);
}
const runsOf = (inlines, base) => toRuns(flat(inlines), base);
const plain = (inlines) => flat(inlines).filter(i => i.k === 'text').map(i => i.text).join('');
const blockText = (b) => {
  if (['Para', 'Plain'].includes(b.t)) return plain(b.c);
  if (b.t === 'Header') return plain(b.c[2]);
  return '';
};

// ---------------------------------------------------------------- stili
function makeStyles() {
  const small = { font: T.font, size: T.small, color: INK };
  const callout = (id, name, o) => ({ id, name, basedOn: 'Normal', next: 'Normal', run: o.run,
    paragraph: { border: o.border, shading: o.shading, indent: { left: 340, right: 340 },
                 spacing: { before: 0, after: 0, ...auto(T.line) }, alignment: AlignmentType.LEFT, keepLines: true } });
  const box = (c, w, sp) => ({ top: line(c, w, sp), bottom: line(c, w, sp), right: line(c, w, sp) });
  return {
    default: {
      // Title e Heading 1-3 si ridefiniscono qui: in paragraphStyles docx-js li duplicherebbe.
      heading1: { run: { font: T.font, size: T.h1, bold: true, color: INK },
                  paragraph: { spacing: { before: 0, after: 280, ...auto(276) }, keepNext: true, keepLines: true, outlineLevel: 0,
                               pageBreakBefore: true, alignment: AlignmentType.LEFT, border: { bottom: line(ACCENT, 24, 6) } } },
      heading2: { run: { font: T.font, size: T.h2, bold: true, color: SECOND },
                  paragraph: { spacing: { before: 0, after: 240, ...auto(276) }, keepNext: true, keepLines: true, outlineLevel: 1,
                               pageBreakBefore: true, alignment: AlignmentType.LEFT } },
      heading3: { run: { font: T.font, size: T.h3, bold: true, color: INK },
                  paragraph: { spacing: { before: 400, after: 140, ...auto(276) }, keepNext: true, keepLines: true, outlineLevel: 2,
                               alignment: AlignmentType.LEFT } },
      // I livelli 4-6 e i collegamenti non servono, ma esistono nel file: li allineo alla palette
      // perché chi li applica a mano in Word non trovi blu fuori marchio.
      heading4: { run: { font: T.font, size: T.h3, bold: true, color: INK }, paragraph: { spacing: { before: 400, after: 140, ...auto(276) }, keepNext: true, outlineLevel: 3 } },
      heading5: { run: { font: T.font, size: T.h3, bold: true, color: INK }, paragraph: { spacing: { before: 400, after: 140, ...auto(276) }, keepNext: true, outlineLevel: 4 } },
      heading6: { run: { font: T.font, size: T.h3, bold: true, color: INK }, paragraph: { spacing: { before: 400, after: 140, ...auto(276) }, keepNext: true, outlineLevel: 5 } },
      hyperlink: { run: { color: INK, underline: {} } },
      document: { run: { font: T.font, size: T.body, color: INK, language: { value: 'it-IT' } },
                  paragraph: { spacing: { ...auto(T.line), after: T.after }, alignment: AlignmentType.LEFT } },
    },
    paragraphStyles: [
      { id: 'ListItem', name: 'Elenco', basedOn: 'Normal', paragraph: { spacing: { after: 100 } } },
      { id: 'Caption', name: 'caption', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: small,
        paragraph: { spacing: { before: 80, after: 320 }, alignment: AlignmentType.CENTER, keepLines: true } },
      { id: 'CaptionTable', name: 'Didascalia tabella', basedOn: 'Normal', next: 'Normal', run: small,
        paragraph: { spacing: { before: 160, after: 100 }, keepNext: true, keepLines: true } },
      // Riquadro Competenze: è una tabella (codici e testi del repertorio sono dati), con stili propri
      { id: 'CompTitle', name: 'Competenze: titolo', basedOn: 'Normal', run: { font: T.font, size: T.body, bold: true, color: ON_SECOND },
        paragraph: { spacing: { before: 0, after: 0, ...auto(276) }, keepNext: true, keepLines: true } },
      { id: 'CompText', name: 'Competenze: voce', basedOn: 'Normal', run: { font: T.font, size: T.comp, color: INK },
        paragraph: { spacing: { before: 0, after: 80, ...auto(276) }, keepLines: true,
                     indent: { left: 600, hanging: 600 }, tabStops: [{ type: TabStopType.LEFT, position: 600 }] } },
      { id: 'TableText', name: 'Testo tabella', basedOn: 'Normal', run: { font: T.font, size: T.tbl, color: INK },
        paragraph: { spacing: { before: 0, after: 0, ...auto(252) }, alignment: AlignmentType.LEFT } },
      // Da ricordare: riquadro bianco con filetto e barra a sinistra nel colore secondario
      callout('CalloutInfo', 'Riquadro: da ricordare', {
        run: { color: INK },
        border: { ...box(SECOND, 8, 8), left: line(SECOND, 36, 8) } }),
      // Attenzione: riquadro pieno nel colore d'accento, testo a contrasto
      callout('CalloutWarn', 'Riquadro: attenzione', {
        run: { color: ON_ACCENT },
        border: { ...box(ACCENT, 8, 8), left: line(ACCENT, 8, 8) },
        shading: { type: ShadingType.CLEAR, fill: ACCENT, color: 'auto' } }),
    ],
  };
}
function makeNumbering() {
  const lvl = (level, fmt, text, left, hanging) => ({ level, format: fmt, text, alignment: AlignmentType.LEFT,
    style: { run: { color: INK, bold: true }, paragraph: { indent: { left, hanging } } } });
  const list = (ref, f0, t0, f1, t1) => ({ reference: ref, levels: [lvl(0, f0, t0, 567, 397), lvl(1, f1, t1, 1134, 397)] });
  return { config: [
    list('bullets', LevelFormat.BULLET, '•', LevelFormat.BULLET, '–'),
    list('steps', LevelFormat.DECIMAL, '%1.', LevelFormat.LOWER_LETTER, '%2)'),
    list('abc', LevelFormat.UPPER_LETTER, '%1)', LevelFormat.DECIMAL, '%2.'),
  ] };
}

// ---------------------------------------------------------------- mattoni
// Ogni "item" ha build(keep): keep=true significa "resta sulla stessa pagina del blocco successivo".
// docx-js non permette di modificare un paragrafo già creato, perciò i blocchi si costruiscono all'ultimo.
const para = (children, o = {}) => new Paragraph({ children, ...o });
const spacer = (keep) => new Paragraph({ keepNext: keep, spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.EXACT }, children: [] });
const item = (build, hint = false) => ({ build, hint });

function listItems(l, refName, level, instance, out) {
  for (const blocks of l) {
    const first = blocks.find(b => ['Plain', 'Para'].includes(b.t));
    out.push({ level, ref: refName, instance, inlines: first ? first.c : [] });
    for (const b of blocks) {
      if (b === first) continue;
      if (b.t === 'BulletList') listItems(b.c, 'bullets', Math.min(level + 1, 1), instance, out);
      else if (b.t === 'OrderedList') listItems(b.c[1], 'steps', Math.min(level + 1, 1), instance, out);
    }
  }
}
function listItem(block) {
  const rows = [];
  if (block.t === 'BulletList') listItems(block.c, 'bullets', 0, ++ctx.list, rows);
  else {
    const [, style] = block.c[0];
    const ref = ['UpperAlpha', 'LowerAlpha'].includes(style.t) ? 'abc' : 'steps';
    listItems(block.c[1], ref, 0, ++ctx.list, rows);
  }
  const short = rows.length <= (block.t === 'BulletList' ? 7 : 10);   // elenchi brevi e procedure restano insieme
  const gap = ctx.inExercise ? 320 : undefined;                          // negli esercizi: aria tra una domanda e la successiva
  return item((keep) => rows.map((r, i) => new Paragraph({
    style: 'ListItem', keepNext: i < rows.length - 1 ? (short || keep) : keep, keepLines: true,
    ...(gap && i === rows.length - 1 ? { spacing: { after: gap } } : {}),
    numbering: { reference: r.ref, level: r.level, instance: r.instance }, children: runsOf(r.inlines) })), false);
}

function headingItem(level, inlines) {
  if (level === 1) { ctx.ex = 0; ctx.stats.units++; ctx.comp.unit = { name: plain(inlines), has: false, box: null }; ctx.comp.units.push(ctx.comp.unit); }
  if (level === 2) ctx.stats.chapters++;
  const n = Math.min(level, 3);
  if (level > 3) warn(`Titolo di livello ${level} trattato come livello 3: ${plain(inlines)}`);
  const it = item(() => [new Paragraph({ style: `Heading${n}`, children: runsOf(inlines) })]);
  if (level === 1) {
    it.unit = plain(inlines);
  }
  return it;
}

function paraItem(b) {
  const imgs = flat(b.c).filter(i => i.k === 'img');
  const onlyImgs = imgs.length && flat(b.c).every(i => i.k === 'img' || (i.k === 'text' && !i.text.trim()));
  if (onlyImgs) return imgs.length === 1 ? figureItem(imgs[0].node, null) : imageRowItem(imgs.map(i => i.node), null);
  const text = plain(b.c).trim();
  return item((keep) => [new Paragraph({ keepNext: keep, keepLines: true, children: runsOf(b.c) })], text.endsWith(':'));
}

function loadImage(node) {
  const url = node.c[2][0];
  const file = path.resolve(BASE_DIR, decodeURI(url));
  if (!fs.existsSync(file)) { warn(`Immagine mancante: ${url}`); return null; }
  const buf = fs.readFileSync(file);
  const size = imageSize(buf);
  ctx.stats.images++;
  return { buf, size, kv: kvOf(node.c[0]), name: path.basename(file), type: imgType(buf), url };
}
function fitCm(img, widthCm, maxW = TEXT_W / CM) {
  let w = widthCm, h = (w * img.size.h) / img.size.w;
  if (w > maxW) { w = maxW; h = (w * img.size.h) / img.size.w; }
  if (h > MAX_IMG_H_CM) { h = MAX_IMG_H_CM; w = (h * img.size.w) / img.size.h; }
  return { w, h };
}
function widthFor(img, availCm) {
  const v = img.kv.width;
  if (v && v.endsWith('%')) return (parseFloat(v) / 100) * availCm;
  if (v && v.endsWith('cm')) return parseFloat(v);
  return Math.min((img.size.w / 150) * 2.54, availCm);          // dimensione naturale a 150 dpi, senza uscire dai margini
}
function imageRun(img, widthCm, altFallback) {
  const { w, h } = fitCm(img, widthCm);
  const alt = img.kv.alt || altFallback || '';
  if (!alt) warn(`Immagine senza testo alternativo né didascalia: ${img.url}`);
  return new ImageRun({ type: img.type, data: img.buf, transformation: { width: Math.round(w * PX_PER_CM), height: Math.round(h * PX_PER_CM) },
    altText: { title: (alt || img.name).slice(0, 120), description: alt || img.name, name: img.name } });
}
function captionPara(label, seqName, n, capInlines, credit, keep) {
  const kids = [new TextRun({ text: label + ' ' }), new SimpleField(`SEQ ${seqName} \\* ARABIC`, String(n)), new TextRun({ text: ' – ' }), ...runsOf(capInlines)];
  if (credit) kids.push(new TextRun({ text: 'Fonte: ' + credit, break: 1 }));
  return new Paragraph({ style: 'Caption', keepNext: keep, children: kids });
}
function figureItem(node, capInlines) {
  const img = loadImage(node);
  if (!img) return item(() => [para([new TextRun({ text: `[IMMAGINE MANCANTE: ${node.c[2][0]}]`, bold: true, color: ACCENT })])]);
  const hasCap = capInlines && capInlines.length;
  ctx.stats.figures += hasCap ? 1 : 0;
  const n = hasCap ? ++ctx.fig : 0;
  return item((keep) => [
    new Paragraph({ keepNext: !!hasCap || keep, alignment: AlignmentType.CENTER, spacing: { before: 240, after: hasCap ? 0 : 240 },
      children: [imageRun(img, widthFor(img, TEXT_W / CM), hasCap ? plain(capInlines) : '')] }),
    ...(hasCap ? [captionPara('Figura', 'Figura', n, capInlines, img.kv.fonte, keep)] : []),
  ]);
}
function imageRowItem(nodes, capInlines, creditKv) {
  const imgs = nodes.map(loadImage).filter(Boolean);
  if (!imgs.length) return item(() => []);
  const hasCap = capInlines && capInlines.length;
  const n = hasCap ? ++ctx.fig : 0; ctx.stats.figures += hasCap ? 1 : 0;
  const each = (TEXT_W / CM - 1.4) / imgs.length;   // margine di sicurezza: la riga non deve mai andare a capo
  return item((keep) => {
    const kids = [];
    imgs.forEach((im, i) => { if (i) kids.push(new TextRun({ text: '    ' })); kids.push(imageRun(im, Math.min(widthFor({ ...im, kv: {} }, each), each), hasCap ? plain(capInlines) : '')); });
    return [new Paragraph({ keepNext: !!hasCap || keep, alignment: AlignmentType.CENTER, spacing: { before: 240, after: hasCap ? 0 : 240 }, children: kids }),
      ...(hasCap ? [captionPara('Figura', 'Figura', n, capInlines, creditKv, keep)] : [])];
  });
}

// ---- riquadri
function calloutItem(kind, title, blocks) {
  const style = kind === 'attenzione' ? 'CalloutWarn' : 'CalloutInfo';
  const lines = [];
  for (const b of blocks) {
    if (['Para', 'Plain'].includes(b.t)) lines.push(b.c);
    else if (b.t === 'BulletList' || b.t === 'OrderedList') {
      const rows = []; listItems(b.t === 'BulletList' ? b.c : b.c[1], 'bullets', 0, 0, rows);
      rows.forEach((r, i) => lines.push([{ t: 'Str', c: (b.t === 'BulletList' ? '–' : `${i + 1}.`) + ' ' }, ...r.inlines]));
    } else warn(`Nei riquadri sono ammessi solo paragrafi ed elenchi (trovato ${b.t})`);
  }
  return item((keep) => [
    new Paragraph({ style, keepNext: true, spacing: { before: 400, after: 100 }, children: [new TextRun({ text: title, bold: true, ...(kind === 'attenzione' ? {} : { color: SECOND }) })] }),
    ...lines.map((l, i) => new Paragraph({ style, keepNext: i < lines.length - 1 ? true : keep,
      ...(i === lines.length - 1 ? { spacing: { before: 0, after: 400 } } : {}), children: runsOf(l) })),
  ]);
}

// ---- riquadro Competenze: qualifica, unità di competenza, capacità e conoscenze (testi dal repertorio)
const byCode = (a, b) => { const [x1, y1] = a.code.split('.').map(Number), [x2, y2] = b.code.split('.').map(Number); return x1 - x2 || y1 - y2; };
function resolveCompetenze(kv) {
  const C = ctx.comp;
  if (kv.qualifica) {
    const found = findQualifica(kv.qualifica);
    if (found.length === 1) C.qual = found[0];
    else warn(found.length ? `qualifica="${kv.qualifica}" è ambigua: ${found.map(q => q.file).join(', ')}` : `qualifica="${kv.qualifica}" non trovata in ${QUAL_DIR}`);
  }
  const q = C.qual;
  if (!q) { warn('Riquadro Competenze senza qualifica: aggiungi qualifica="04" (numero del file) al primo riquadro'); return null; }
  const codes = parseCodes(kv.codici);
  if (!codes.length) warn(`Riquadro Competenze senza codici (${C.unit ? C.unit.name : 'prima della prima unità'}): aggiungi codici="1.2 1.6"`);
  const groups = {}, missing = [];
  const group = (n) => (groups[n] = groups[n] || { uc: q.ucs[n], cap: [], con: [] });
  for (const c of codes) {
    const m = c.match(/^UC(\d+)$/);
    if (m) { if (q.ucs[m[1]]) group(m[1]); else missing.push(c); continue; }
    const it = q.codes[c];
    if (!it) { missing.push(c); continue; }
    const g = group(it.uc);
    if (!g[it.kind].includes(it)) g[it.kind].push(it);
    if (it.text.includes('[…]')) warn(`${q.title}, ${it.code}: il testo regionale è incompleto (segno […]); controllalo prima di stampare`);
  }
  if (missing.length) warn(`${q.title}: codici inesistenti ${missing.join(', ')} (vedi ${q.file})`);
  const ordered = Object.keys(groups).sort((a, b) => a - b).map(n => ({ ...groups[n], cap: groups[n].cap.sort(byCode), con: groups[n].con.sort(byCode) }));
  if (!C.used.has(q)) C.used.set(q, new Set());
  ordered.forEach(g => { C.used.get(q).add('UC' + g.uc.n); [...g.cap, ...g.con].forEach(it => C.used.get(q).add(it.code)); });
  return { q, groups: ordered, missing };
}
function competenzeItem(kv) {
  const C = ctx.comp;
  if (C.unit) C.unit.has = true; else warn('Riquadro Competenze prima della prima unità (#)');
  let data;
  if (kv.riepilogo && kv.riepilogo !== 'false') {                  // nella verifica: ripete il riquadro dell'unità
    data = C.unit && C.unit.box;
    if (!data) { warn(`Riquadro Competenze con riepilogo ma senza un riquadro prima nell'unità (${C.unit ? C.unit.name : '?'})`); return item(() => []); }
  } else {
    data = resolveCompetenze(kv);
    if (data && C.unit && !C.unit.box) C.unit.box = data;
  }
  const R = (text, o = {}) => new TextRun({ text, ...o });
  const W = TEXT_W, HALF = Math.round(W / 2);
  const B = (c, sz) => ({ style: BorderStyle.SINGLE, size: sz, color: c });
  const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  // keep: true finché non è l'ultima riga, così il riquadro resta intero (se non ci sta, passa alla pagina dopo)
  const cell = (children, o = {}) => new TableCell({ verticalAlign: VerticalAlign.TOP,
    margins: { top: o.top ?? 100, bottom: o.bottom ?? 100, left: 180, right: 180 },
    width: { size: o.span ? W : HALF, type: WidthType.DXA }, ...(o.span ? { columnSpan: 2 } : {}),
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    borders: { top: o.bTop || NONE, bottom: o.bBottom || NONE, left: o.bLeft || B(SECOND, 12), right: o.bRight || B(SECOND, 12) },
    children });
  const label = (t, keep) => new Paragraph({ style: 'CompText', keepNext: keep, indent: { left: 0, hanging: 0 }, spacing: { after: 80 },
    children: [R(t.toUpperCase(), { bold: true, color: SECOND })] });
  const entry = (it, keep, last) => new Paragraph({ style: 'CompText', keepNext: keep, ...(last ? { spacing: { after: 0 } } : {}),
    children: [R(it.code, { bold: true, color: SECOND }), new TextRun({ children: [new Tab()] }), R(it.text)] });
  const listCell = (title, list, keep, o) => cell([label(title, true), ...list.map((it, i) => entry(it, keep, i === list.length - 1))], o);
  const title = kv.title || 'COMPETENZE';
  return item((keep) => {
    const rows = [];
    const k = (last) => (last ? keep : true);
    rows.push(new TableRow({ tableHeader: true, cantSplit: true, children: [cell([
      new Paragraph({ style: 'CompTitle', children: [R(title)] }),
      ...(data ? [new Paragraph({ style: 'CompTitle', children: [R('Qualifica: ', { size: T.comp }), R(data.q.title, { size: T.comp })] })] : []),
    ], { span: true, fill: SECOND, top: 140, bottom: 140, bTop: B(SECOND, 12), bLeft: B(SECOND, 12), bRight: B(SECOND, 12) })] }));
    const groups = data ? data.groups : [];
    if (!data) rows.push(new TableRow({ cantSplit: true, children: [cell([new Paragraph({ style: 'TableText', children: [R('[QUALIFICA O CODICI MANCANTI]', { bold: true, color: ACCENT })] })], { span: true, bBottom: B(SECOND, 12) })] }));
    groups.forEach((g, gi) => {
      const lastG = gi === groups.length - 1 && !(data.missing.length);
      rows.push(new TableRow({ cantSplit: true, children: [cell([new Paragraph({ style: 'TableText', keepNext: true, children: [
        R(`Unità di competenza ${g.uc.n} (UC${g.uc.n})`, { bold: true, color: SECOND }), R(' – ' + g.uc.title, { bold: true }) ] })],
        { span: true, top: 160, bottom: 60, bTop: gi ? B(SECOND, 6) : NONE,
          ...(lastG && !g.cap.length && !g.con.length ? { bBottom: B(SECOND, 12), bottom: 160 } : {}) })] }));
      const bot = lastG ? { bBottom: B(SECOND, 12) } : {};
      const kk = k(lastG);
      if (g.cap.length && g.con.length) rows.push(new TableRow({ cantSplit: true, children: [
        listCell('Capacità', g.cap, kk, { ...bot, top: 60, bottom: 160, bRight: B(SECOND, 4) }),
        listCell('Conoscenze', g.con, kk, { ...bot, top: 60, bottom: 160, bLeft: B(SECOND, 4) }) ] }));
      else if (g.cap.length || g.con.length) rows.push(new TableRow({ cantSplit: true, children: [
        listCell(g.cap.length ? 'Capacità' : 'Conoscenze', g.cap.length ? g.cap : g.con, kk, { ...bot, span: true, top: 60, bottom: 160 })] }));
    });
    if (data && data.missing.length) rows.push(new TableRow({ cantSplit: true, children: [cell([new Paragraph({ style: 'TableText',
      children: [R('Codici non trovati: ' + data.missing.join(', '), { bold: true, color: ACCENT })] })], { span: true, bTop: B(SECOND, 6), bBottom: B(SECOND, 12) })] }));
    const tbl = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [HALF, W - HALF], layout: TableLayoutType.FIXED, rows,
      borders: { top: B(SECOND, 12), bottom: B(SECOND, 12), left: B(SECOND, 12), right: B(SECOND, 12), insideHorizontal: NONE, insideVertical: NONE } });
    return [spacer(true), tbl, spacer(keep)];
  });
}

// ---- tabelle
function tableItem(b) {
  const [, cap, colspecs, head, bodies] = b.c;
  const capInlines = (cap[1][0] && ['Plain', 'Para'].includes(cap[1][0].t)) ? cap[1][0].c : null;
  const cellText = (c) => c[4].map(blockText).join(' ');
  const headRows = head[1].map(r => r[1]);
  const bodyRows = bodies.flatMap(tb => [...tb[2], ...tb[3]]).map(r => r[1]);
  const nCols = colspecs.length;
  // larghezze: da pandoc se specificate, altrimenti in proporzione al contenuto
  let widths;
  if (colspecs.every(cs => cs[1].t === 'ColWidth')) widths = colspecs.map(cs => cs[1].c);
  else {
    const lens = Array.from({ length: nCols }, (_, i) => Math.max(...[...headRows, ...bodyRows].map(r => (r[i] ? cellText(r[i]).length : 0)), 4));
    widths = lens.map(l => Math.max(8, Math.min(l, 40)));
  }
  const sum = widths.reduce((a, c) => a + c, 0);
  let cols = widths.map(w => Math.max(1100, Math.round((w / sum) * TEXT_W)));
  const diff = TEXT_W - cols.reduce((a, c) => a + c, 0); cols[cols.indexOf(Math.max(...cols))] += diff;
  const aligns = colspecs.map(cs => ({ AlignRight: AlignmentType.RIGHT, AlignCenter: AlignmentType.CENTER }[cs[0].t] || AlignmentType.LEFT));
  const borders = { top: line(INK), bottom: line(INK), left: line(INK), right: line(INK) };
  const mkCell = (c, i, isHead, keep, minH) => new TableCell({
    width: { size: cols[i], type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    shading: isHead ? { type: ShadingType.CLEAR, fill: SECOND, color: 'auto' } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: (c[4].length ? c[4] : [{ t: 'Plain', c: [] }]).map(blk => new Paragraph({ style: 'TableText', keepNext: keep, alignment: aligns[i],
      children: runsOf(['Plain', 'Para'].includes(blk.t) ? blk.c : [{ t: 'Str', c: blockText(blk) }], isHead ? { bold: true, color: ON_SECOND } : {}) })),
  });
  const chain = bodyRows.length <= 14;                               // le tabelle brevi non si spezzano
  const n = capInlines ? ++ctx.tab : 0; ctx.stats.tables++;
  return item((keep) => {
    const rows = [];
    headRows.forEach(r => rows.push(new TableRow({ tableHeader: true, cantSplit: true, children: r.map((c, i) => mkCell(c, i, true, true)) })));
    bodyRows.forEach((r, ri) => {
      const fillIn = r.some(c => !cellText(c).trim());               // tabelle da completare: righe alte per scrivere
      rows.push(new TableRow({ cantSplit: true, ...(fillIn ? { height: { value: 760, rule: HeightRule.AT_LEAST } } : {}),
        children: r.map((c, i) => mkCell(c, i, false, ri < bodyRows.length - 1 ? chain || keep : keep)) }));
    });
    return [
      ...(capInlines ? [new Paragraph({ style: 'CaptionTable', children: [new TextRun({ text: 'Tabella ' }), new SimpleField('SEQ Tabella \\* ARABIC', String(n)), new TextRun({ text: ' – ' }), ...runsOf(capInlines)] })] : []),
      new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: cols, layout: TableLayoutType.FIXED, rows }),
      spacer(keep),
    ];
  });
}

// ---- esercizi
const vfItem = (b) => item((keep) => b.c.map((blocks, i) => {
  const first = blocks.find(x => ['Plain', 'Para'].includes(x.t));
  return new Paragraph({ keepNext: i < b.c.length - 1 ? true : keep, keepLines: true, spacing: { after: 160 },
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W }],
    children: [...runsOf(first ? first.c : []), new TextRun({ children: [new Tab(), 'V □    F □'] })] });
}));
const linesItem = (n) => item((keep) => Array.from({ length: n }, (_, i) => new Paragraph({
  keepNext: i < n - 1 ? true : keep, spacing: { before: 0, after: 0, line: 620, lineRule: LineRuleType.EXACT },
  border: { bottom: line(INK), between: line(INK) }, children: [] })));

function estimateCm(blocks) {                                        // stima grossolana dell'altezza, per decidere se spezzare
  let cm = 1.6;
  const lines = (t, w = 90) => Math.max(1, Math.ceil(t.length / w));
  for (const b of blocks) {
    if (['Para', 'Plain'].includes(b.t)) cm += lines(plain(b.c)) * 0.64 + 0.35 + (flat(b.c).some(i => i.k === 'img') ? MAX_IMG_H_CM * 0.7 : 0);
    else if (b.t === 'BulletList') b.c.forEach(x => { cm += lines(x.map(blockText).join(' '), 81) * 0.64 + 0.18; });
    else if (b.t === 'OrderedList') b.c[1].forEach(x => { cm += lines(x.map(blockText).join(' '), 81) * 0.64 + 0.18; });
    else if (b.t === 'Table') cm += (b.c[4].reduce((a, tb) => a + tb[3].length, 0) + 1) * 1.1 + 1;
    else if (b.t === 'Figure') cm += MAX_IMG_H_CM * 0.75 + 1.3;
    else if (b.t === 'Div') {
      const kv = kvOf(b.c[0]), cl = b.c[0][1];
      if (cl.includes('righe')) cm += (parseInt(kv.n || '3', 10)) * 1.09;
      else if (cl.includes('vf')) cm += b.c[1].reduce((a, x) => a + (x.t === 'BulletList' ? x.c.length * 0.9 : 0), 0);
      else if (cl.includes('coppia')) cm += 7;
      else cm += estimateCm(b.c[1]);
    }
  }
  return cm;
}
function exerciseItems(attr, blocks) {
  const kv = kvOf(attr), n = ++ctx.ex; ctx.stats.exercises++;
  const title = `Esercizio ${n}` + (kv.tipo ? ` · ${kv.tipo}` : '');
  const est = estimateCm(blocks);
  const split = kv.split === 'true' || est > SPLIT_ABOVE_CM;
  if (split) { ctx.stats.splitEx++; console.log(`Nota: ${title} è lungo (≈${est.toFixed(0)} cm): le domande restano intere ma l'esercizio può continuare sulla pagina dopo.`); }
  // gruppi: un nuovo gruppo inizia a un paragrafo che segue un blocco non-paragrafo
  const groups = []; let prevPara = false;
  for (const b of blocks) {
    const isPara = ['Para', 'Plain'].includes(b.t);
    if (!groups.length || (isPara && !prevPara)) groups.push([]);
    groups[groups.length - 1].push(b); prevPara = isPara;
  }
  ctx.inExercise = true;
  const groupItems = groups.map(g => g.flatMap(blockToItems));
  ctx.inExercise = false;
  const head = item(() => [new Paragraph({ style: 'Heading3', children: [new TextRun(title)] })], true);
  if (!split) return [item((keep) => [].concat(head.build(true), ...groupItems.flat().map((it, i, arr) => it.build(i < arr.length - 1 ? true : keep))))];
  return groupItems.map((its, gi) => item((keep) => [].concat(...(gi === 0 ? [head.build(true)] : []),
    ...its.map((it, i) => it.build(i < its.length - 1 ? true : keep && false)))));
}

// ---------------------------------------------------------------- blocchi pandoc -> item
function blockToItems(b) {
  switch (b.t) {
    case 'Header': return [headingItem(b.c[0], b.c[2])];
    case 'Para': case 'Plain': return [paraItem(b)];
    case 'BulletList': case 'OrderedList': return [listItem(b)];
    case 'Table': return [tableItem(b)];
    case 'Figure': {
      const img = b.c[2].flatMap(x => (x.c || [])).find(i => i && i.t === 'Image');
      const cap = b.c[1][1][0] ? b.c[1][1][0].c : null;
      return img ? [figureItem(img, cap)] : [];
    }
    case 'DefinitionList': return b.c.map(([term, defs]) => item((keep) => [new Paragraph({ keepLines: true, keepNext: keep,
      children: [...runsOf(term, { bold: true }), new TextRun({ text: ' – ' }), ...runsOf(defs.flat().flatMap(x => ['Para', 'Plain'].includes(x.t) ? x.c : []))] })]));
    case 'BlockQuote': return b.c.flatMap(blockToItems);
    case 'CodeBlock': return [item((keep) => [new Paragraph({ keepNext: keep, children: [new TextRun(b.c[1])] })])];
    case 'LineBlock': return b.c.map(l => item((keep) => [new Paragraph({ keepNext: keep, children: runsOf(l) })]));
    case 'HorizontalRule': case 'RawBlock': return [];
    case 'Div': {
      const [attr, inner] = b.c, cl = attr[1], kv = kvOf(attr);
      if (cl.includes('ricorda')) return [calloutItem('ricorda', kv.title || 'DA RICORDARE', inner)];
      if (cl.includes('attenzione')) return [calloutItem('attenzione', kv.title || 'ATTENZIONE', inner)];
      if (cl.includes('competenze')) return [competenzeItem(kv)];
      if (cl.includes('esercizio')) return exerciseItems(attr, inner);
      if (cl.includes('vf')) return inner.filter(x => x.t === 'BulletList').map(vfItem);
      if (cl.includes('righe')) return [linesItem(parseInt(kv.n || '3', 10))];
      if (cl.includes('coppia')) {
        const imgs = inner.flatMap(x => flat(x.c || []).filter(i => i.k === 'img').map(i => i.node));
        const capText = kv.caption ? [{ t: 'Str', c: kv.caption }] : null;
        return [imageRowItem(imgs, capText, kv.fonte)];
      }
      if (cl.includes('nuovapagina')) return [item(() => [new Paragraph({ children: [new PageBreak()] })])];
      return inner.flatMap(blockToItems);
    }
    default: warn(`Blocco non gestito: ${b.t}`); return [];
  }
}

// docx-js scrive i bordi come top/bottom/left/right; lo schema OOXML vuole top/left/bottom/right/between.
async function fixBorderOrder(buf) {
  const order = ['top', 'left', 'bottom', 'right', 'between', 'bar'];
  const tag = (x) => x.match(/<w:(\w+)/)[1];
  const fix = (xml) => xml.replace(/<w:pBdr>([\s\S]*?)<\/w:pBdr>/g, (m, inner) => {
    const els = inner.match(/<w:(?:top|left|bottom|right|between|bar)\b[^>]*\/>/g) || [];
    els.sort((a, b) => order.indexOf(tag(a)) - order.indexOf(tag(b)));
    return '<w:pBdr>' + els.join('') + '</w:pBdr>';
  });
  const zip = await JSZip.loadAsync(buf);
  for (const name of Object.keys(zip.files)) {
    if (/^word\/[^/]+\.xml$/.test(name)) zip.file(name, fix(await zip.file(name).async('string')));
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------- main
(async () => {
  let ast;
  try { ast = JSON.parse(execFileSync('pandoc', [input, '-f', 'markdown', '-t', 'json'], { maxBuffer: 1 << 28, encoding: 'utf8' })); }
  catch (e) {
    const msg = String(e.stderr || e.message);
    console.error('pandoc non ha letto il file:', msg);
    if (/YAML/.test(msg)) console.error('Suggerimento: nell’intestazione tra --- i valori con due punti (:) o altri simboli vanno tra virgolette, per esempio  title: "Impianti: introduzione".');
    process.exit(1);
  }

  const items = ast.blocks.flatMap(blockToItems);
  if (!items.length) items.push(item(() => [new Paragraph({ children: [] })]));
  const title = path.basename(input).replace(/\.[^.]+$/, '');
  const pageProps = { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, header: HF_DIST, footer: HF_DIST } } };
  const doc = new Document({
    creator: 'Oficina', title, description: 'Dispensa Oficina',
    styles: makeStyles(), numbering: makeNumbering(),
    // Solo il contenuto: niente copertina, intestazione né piè di pagina.
    sections: [{ properties: pageProps, children: items.flatMap(it => [].concat(it.build(!!it.hint))) }],
  });
  fs.writeFileSync(OUT, await fixBorderOrder(await Packer.toBuffer(doc)));
  const s = ctx.stats;
  console.log(`Scritto ${OUT}\n  ${s.units} unità · ${s.chapters} capitoli · ${s.figures} figure (${s.images} immagini) · ${s.tables} tabelle · ${s.exercises} esercizi (${s.splitEx} lunghi)`);
  const C = ctx.comp;
  const without = C.units.filter(u => !u.has).map(u => u.name);
  if (without.length) warn(`Unità senza riquadro Competenze: ${without.join('; ')}`);
  for (const [q, used] of C.used) {
    console.log(`Competenze – ${q.title}:`);
    for (const n of Object.keys(q.ucs).sort((a, b) => a - b)) {
      const uc = q.ucs[n], codes = [...uc.cap, ...uc.con].filter(it => used.has(it.code)).map(it => it.code);
      console.log(`  UC${n} – ${uc.title}: ` + (codes.length ? codes.join(', ') : used.has('UC' + n) ? 'solo il titolo' : 'non trattata'));
    }
  }
  if (ctx.warnings.length) { console.log('Avvisi:'); [...new Set(ctx.warnings)].forEach(w => console.log('  - ' + w)); }
})();
