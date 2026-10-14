#!/usr/bin/env python3
"""Controlla che una dispensa .docx rispetti il layout standard Oficina.

Uso:  python check_dispensa.py dispensa.docx [--veloce] [--brand ../assets/brand.json]

--veloce salta il rendering in PDF (più rapido): restano i controlli sul file.

Controlli sul file (XML): A4, margini 1,27 cm (stretti), un solo font, misure ammesse (12/16/20 pt; 10 pt nel riquadro Competenze),
nessun grigio, immagini entro i margini e con testo alternativo.
Riquadro Competenze in ogni unità (avviso se manca; errore se ha codici inesistenti).
Controlli sul rendering (PDF via LibreOffice): unità e capitoli in cima a una pagina nuova, didascalie sulla stessa pagina dell'immagine, esercizi spezzati (da tenere rari).
Esce con codice 1 se un controllo obbligatorio fallisce.
"""
import json, os, re, subprocess, sys, zipfile

docx = sys.argv[1]
here = os.path.dirname(os.path.abspath(__file__))
brand_path = sys.argv[sys.argv.index("--brand") + 1] if "--brand" in sys.argv else os.path.join(here, "..", "assets", "brand.json")
brand = json.load(open(brand_path, encoding="utf-8"))
ALLOWED_COLORS = {"000000", "FFFFFF", "AUTO"} | {brand[k].upper() for k in ("primary", "accent", "secondary") if k in brand}
ALLOWED_SIZES = {12, 16, 20, 10}   # 10 pt solo per capacità e conoscenze nel riquadro Competenze
fails, warns, lines_out = [], [], []


def ok(msg): lines_out.append(f"  ✔ {msg}")
def ko(msg): fails.append(msg); lines_out.append(f"  ✘ {msg}")
def wa(msg): warns.append(msg); lines_out.append(f"  ! {msg}")


z = zipfile.ZipFile(docx)
read = lambda n: z.read(n).decode("utf-8")
parts = {n: read(n) for n in z.namelist() if re.match(r"word/(document|styles|header\d*|footer\d*)\.xml$", n)}
doc, sty = parts["word/document.xml"], parts["word/styles.xml"]
allxml = "".join(parts.values())

print(f"\nControllo: {os.path.basename(docx)}")
print("Pagina e margini")
sects = re.findall(r"<w:sectPr.*?</w:sectPr>", doc, re.S)
good = True
for i, s in enumerate(sects, 1):
    pg = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"', s)
    mar = dict(re.findall(r'w:(top|right|bottom|left|header|footer)="(\d+)"', re.search(r"<w:pgMar[^>]*>", s).group(0)))
    if (pg.group(1), pg.group(2)) != ("11906", "16838"): ko(f"sezione {i}: pagina non A4"); good = False
    if any(abs(int(mar[k]) - 720) > 3 for k in ("top", "right", "bottom", "left")): ko(f"sezione {i}: margini diversi da 1,27 cm {mar}"); good = False
if good: ok("A4, margini 1,27 cm (stretti)")

print("Caratteri e colori")
fonts = sorted(set(re.findall(r'w:ascii="([^"]+)"', allxml)))
ok(f"un solo font: {fonts[0]}") if fonts == ["Arial"] else ko(f"font diversi da Arial: {fonts}")
used = set(re.findall(r'<w:pStyle w:val="([^"]+)"', doc + parts.get("word/header1.xml", "")))
used |= {m for n, x in parts.items() if "header" in n or "footer" in n for m in re.findall(r'<w:pStyle w:val="([^"]+)"', x)}
sz = set(re.findall(r'<w:sz w:val="(\d+)"', doc))
sz |= set(re.findall(r'<w:sz w:val="(\d+)"', re.search(r"<w:docDefaults>.*?</w:docDefaults>", sty, re.S).group(0)))
for sid in used:
    m = re.search(r'<w:style [^>]*w:styleId="%s".*?</w:style>' % re.escape(sid), sty, re.S)
    if m: sz |= set(re.findall(r'<w:sz w:val="(\d+)"', m.group(0)))
sizes = sorted({int(v) / 2 for v in sz})
bad = [s for s in sizes if s not in ALLOWED_SIZES]
ok(f"misure usate (pt): {sizes}") if not bad else ko(f"misure fuori scala: {bad}")
colors = {c.upper() for c in re.findall(r'w:(?:color|fill) w:val="([0-9A-Fa-f]{6}|auto)"', allxml)} | \
         {c.upper() for c in re.findall(r'w:color="([0-9A-Fa-f]{6}|auto)"', allxml)} | \
         {c.upper() for c in re.findall(r'w:fill="([0-9A-Fa-f]{6}|auto)"', allxml)}
greys = [c for c in colors if c != "AUTO" and c[0:2] == c[2:4] == c[4:6] and c not in ("000000", "FFFFFF")]
other = sorted(colors - ALLOWED_COLORS - set(greys))
ok(f"colori usati: {sorted(colors)}") if not greys and not other else ko(f"colori non ammessi — grigi: {greys}, altri: {other}")

print("Immagini")
exts = re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', doc)
wide = [round(int(cx) / 360000, 1) for cx, cy in exts if int(cx) > 6645910 + 3600]
ok(f"{len(exts)} immagini, tutte entro i 18,4 cm di larghezza utile") if not wide else ko(f"immagini oltre i margini (cm): {wide}")
noalt = len(re.findall(r'<wp:docPr [^>]*?(?<!descr)/?>', doc)) - len(re.findall(r'<wp:docPr [^>]*descr="[^"]+"', doc))
fname_alt = len(re.findall(r'<wp:docPr [^>]*descr="[^"]*\.(?:png|jpe?g)"', doc))
if noalt > 0: ko(f"{noalt} immagini senza testo alternativo")
elif fname_alt: ko(f"{fname_alt} immagini con il nome del file al posto del testo alternativo: scrivi alt=\"...\"")
else: ok("tutte le immagini hanno un testo alternativo")

print("Testo e tabelle")
txt = lambda x: "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x))
tbls = [t for t in re.findall(r"<w:tbl>.*?</w:tbl>", doc, re.S) if 'w:val="CompTitle"' not in t]   # il riquadro Competenze non conta
tot, intbl = len(txt(doc)), len(txt("".join(tbls)))
(ok if intbl / max(tot, 1) < 0.35 else wa)(f"testo dentro tabelle: {100 * intbl / max(tot, 1):.0f}% (le tabelle servono per i dati)")

print("Competenze")
units, cur = [], None
for p in re.findall(r"<w:p[ >].*?</w:p>", doc, re.S):
    st = re.search(r'<w:pStyle w:val="([^"]+)"', p)
    st = st.group(1) if st else ""
    if st == "Heading1":
        cur = [txt(p), False]; units.append(cur)
    elif st == "CompTitle" and cur:
        cur[1] = True
missing = [u[0] for u in units if not u[1]]
if not units: wa("nessuna unità (#) nel documento")
elif not missing: ok(f"{len(units)} unità, tutte con il riquadro Competenze")
else: wa("unità senza riquadro Competenze: " + "; ".join(missing))
if "NON TROVATI" in txt(doc) or "Codici non trovati" in txt(doc) or "CODICI MANCANTI" in txt(doc):
    ko("un riquadro Competenze contiene codici inesistenti o non ha la qualifica")

# ---- rendering
pdf = os.path.splitext(docx)[0] + ".pdf"
soffice = "/mnt/skills/public/docx/scripts/office/soffice.py"
cmd = ["python3", soffice, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(os.path.abspath(docx)), docx] if os.path.exists(soffice) \
    else ["soffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(os.path.abspath(docx)), docx]
if "--veloce" not in sys.argv: subprocess.run(cmd, capture_output=True)
if "--veloce" in sys.argv:
    print("Impaginazione: saltata (--veloce)")
elif not os.path.exists(pdf):
    wa("PDF non generato: salto i controlli di impaginazione")
else:
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
    pages = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", pdf, "-"], capture_output=True, text=True).stdout
        pages.append([l.strip() for l in t.splitlines() if l.strip()])
    print(f"Impaginazione ({n} pagine)")
    hdr = 0
    heads, notop = [], []
    for pi, L in enumerate(pages, 1):
        for li, l in enumerate(L):
            if li >= hdr and re.match(r"(Unità|Capitolo) \d+ –", l):
                heads.append((pi, li, l))
                if li != hdr: notop.append(f"p{pi}: {l}")
    ok(f"{len(heads)} tra unità e capitoli, tutti in cima a una pagina nuova") if not notop else ko("titoli non in cima alla pagina: " + "; ".join(notop))
    # esercizi: quali occupano più pagine
    allh = [(pi, li, l) for pi, L in enumerate(pages, 1) for li, l in enumerate(L) if li >= hdr and re.match(r"(Unità|Capitolo) \d+ –|Esercizio \d+", l)]
    spans = []
    for k, (pi, li, l) in enumerate(allh):
        if not l.startswith("Esercizio"): continue
        nxt = allh[k + 1] if k + 1 < len(allh) else (n + 1, hdr, "")
        end = nxt[0] - 1 if nxt[1] == hdr and nxt[0] > pi else nxt[0]
        if end > pi: spans.append(f"{l[:40]} (p{pi}–p{end})")
    n_ex = sum(1 for h in allh if h[2].startswith("Esercizio"))
    if not spans: ok(f"{n_ex} esercizi, tutti su una sola pagina")
    elif len(spans) <= max(1, n_ex // 4): ok(f"{n_ex} esercizi, {len(spans)} su più pagine (eccezione prevista per i lunghi): " + "; ".join(spans))
    else: wa(f"{len(spans)} esercizi su {n_ex} occupano più pagine: " + "; ".join(spans))
    # didascalie sulla stessa pagina dell'immagine 
    imgs = subprocess.run(["pdfimages", "-list", pdf], capture_output=True, text=True).stdout.splitlines()[2:]
    per_page = {}
    for r in imgs:
        c = r.split()
        if len(c) > 2: per_page[int(c[0])] = per_page.get(int(c[0]), 0) + 1
    miss = []
    for pi, L in enumerate(pages, 1):
        for l in L:
            if re.match(r"Figura \d+ –", l) and per_page.get(pi, 0) < 1: miss.append(f"{l[:30]} (p{pi})")
    ok("ogni didascalia “Figura N” è nella stessa pagina della sua immagine") if not miss else ko("didascalie senza immagine nella pagina: " + "; ".join(miss))

print("\n".join(lines_out))
print(f"\nRisultato: {'TUTTO OK' if not fails else str(len(fails)) + ' controlli falliti'}" + (f", {len(warns)} avvisi" if warns else ""))
sys.exit(1 if fails else 0)
