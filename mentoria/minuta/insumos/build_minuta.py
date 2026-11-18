# -*- coding: utf-8 -*-
"""Monta a minuta (.docx) a partir das seções em Markdown:
ABNT (A4, margens 3/2, Arial 12, 1,5), notas de rodapé reais do Word,
citações longas com recuo de 4 cm, validação das chamadas autor-data contra a
bibliografia e lista de REFERÊNCIAS só com as obras citadas.
Uso: build_minuta.py <dir_secoes> <bib_marcada.json> <saida.docx> <relatorio.json>"""
import sys, os, re, json, glob, zipfile, shutil, tempfile, unicodedata
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

sec_dir, bib_json, out_docx, rep_json = sys.argv[1:5]
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def strip_acc(s):
    s = unicodedata.normalize('NFKD', s)
    return "".join(c for c in s if not unicodedata.combining(c))

# ---------- 1) leitura das seções ----------
files = sorted(glob.glob(os.path.join(sec_dir, '*.md')))
blocks = []          # (tipo, texto) ; tipos: h1, h2, chaptitle, p, quote, pagebreak
notes = {}           # global id -> texto
gid = 0
rol_blocks = []
for f in files:
    raw = open(f, encoding='utf-8').read().replace('\r', '')
    defs = {}
    body_lines = []
    for ln in raw.split('\n'):
        m = re.match(r'^\[\^(\w+)\]:\s*(.*)$', ln)
        if m:
            defs[m.group(1)] = m.group(2).strip(); last = m.group(1)
        else:
            body_lines.append(ln)
    local2g = {}
    def repl(m):
        global gid
        k = m.group(1)
        if k not in local2g:
            gid += 1; local2g[k] = gid
        return f'§FN{local2g[k]}§'
    text = '\n'.join(body_lines)
    text = re.sub(r'\[\^(\w+)\]', repl, text)
    for k, g in local2g.items():
        notes[g] = defs.get(k, '[nota sem texto]')
    # parágrafos
    target = blocks
    prev_h1_cap = False
    for para in re.split(r'\n\s*\n', text):
        para = para.strip()
        if not para:
            continue
        if para.startswith('# '):
            h = para[2:].strip()
            if h.upper().startswith('ROL DE CATEGORIAS'):
                target = rol_blocks
                target.append(('h1', h)); continue
            target = blocks
            target.append(('h1', h)); prev_h1_cap = h.upper().startswith('CAPÍTULO'); continue
        if para.startswith('### '):
            target.append(('h3', para[4:].strip())); prev_h1_cap = False; continue
        if para.startswith('## '):
            target.append(('h2', para[3:].strip())); prev_h1_cap = False; continue
        if prev_h1_cap:
            target.append(('chaptitle', para)); prev_h1_cap = False; continue
        if para.startswith('>'):
            q = ' '.join(l.lstrip('> ').strip() for l in para.split('\n'))
            target.append(('quote', q)); continue
        target.append(('p', ' '.join(l.strip() for l in para.split('\n'))))

# ---------- 2) validação das chamadas autor-data ----------
bib = json.load(open(bib_json, encoding='utf-8'))
bib_plain = [e.replace('«', '').replace('»', '') for e in bib]
ALIAS = {'CNJ': 'CONSELHO NACIONAL DE JUSTICA', 'OEA': 'ORGANIZACAO DOS ESTADOS AMERICANOS',
         'FIOCRUZ': 'FUNDACAO OSWALDO CRUZ', 'STJ': 'SUPERIOR TRIBUNAL DE JUSTICA',
         'ANADEP': 'MOURA', 'DPU': 'ESTEVES'}
def authors_segment(entry):
    m = re.match(r"^(.*?[a-zà-ÿ]{2,}\)?)\. ", entry)        # termina no 1º ponto após prenome
    if m and len(m.group(1)) < 260:
        return m.group(1)
    return entry.split('. ', 1)[0]                            # entidades em caixa alta
def entry_matches(surname, year, entry):
    E = strip_acc(entry).upper()
    A = strip_acc(authors_segment(entry)).upper()
    s = ALIAS.get(strip_acc(surname).upper(), strip_acc(surname).upper())
    yr = year[:4]
    return bool(re.search(r'(^|[^A-Z])' + re.escape(s) + r'([^A-Z]|$)', A)) and (yr in E)

all_text = '\n'.join(t for _, t in blocks + rol_blocks) + '\n' + '\n'.join(notes.values())
cites = set()
# parentéticas: (SOBRENOME, 2019) ; (A; B, 2004) ; (A, 2010; B, 2008, p. 3)
for grp in re.findall(r'\(([^()]*?\b(?:1[5-9]|20)\d\d[a-z]?\b[^()]*)\)', all_text):
    parts = [p.strip() for p in grp.split(';')]
    pending = []
    for p in parts:
        m = re.match(r"^(?:cf\.\s*|apud\s+)?([A-ZÀ-Ý][A-ZÀ-Ý'\-\s\.&]+?)(?:,\s*((?:[A-Z]\.\s?)+))?(?:\s+et al\.)?,\s*((?:1[5-9]|20)\d\d[a-z]?)", p)
        if m:
            sur = m.group(1).strip().split(' ')[0]
            cites.add((sur, m.group(3), (m.group(2) or '').strip()))
            for a in pending: cites.add((a, m.group(3), ''))
            pending = []
        elif re.match(r"^[A-ZÀ-Ý][A-ZÀ-Ý'\-\s\.]+$", p):
            pending.append(p.split(' ')[0])
# narrativas: Sobrenome (2019) / Sobrenome et al. (2006)
for m in re.finditer(r"\b([A-ZÀ-Ý][a-zà-ÿ'\-]+(?:\s(?:de|da|do|dos)\s[A-ZÀ-Ý][a-zà-ÿ]+)?)(?:\s+et al\.)?\s+\(((?:1[5-9]|20)\d\d[a-z]?)(?:,\s*p\.[^)]*)?\)", all_text):
    cites.add((m.group(1).split(' ')[-1], m.group(2), ''))
cited_idx, unmatched = set(), []
for s, y, ini in sorted(cites):
    hits = [i for i, e in enumerate(bib_plain) if entry_matches(s, y, e)]
    if ini:
        hits = [i for i in hits if (strip_acc(s).upper() + ', ' + ini[0]) in strip_acc(bib_plain[i]).upper()]
    if hits:
        cited_idx.update(hits)
    else:
        unmatched.append(f'{s}, {ini + ", " if ini else ""}{y}')

# ---------- 3) documento ----------
doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
sec.top_margin, sec.left_margin = Cm(3), Cm(3)
sec.bottom_margin, sec.right_margin = Cm(2), Cm(2)
st = doc.styles['Normal']
st.font.name = 'Arial'; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
pf = st.paragraph_format
pf.line_spacing = 1.5; pf.space_after = Pt(0); pf.space_before = Pt(0)
fnt = doc.styles.add_style('Footnote Text', WD_STYLE_TYPE.PARAGRAPH)
fnt.base_style = st; fnt.font.size = Pt(10)
fnt.paragraph_format.line_spacing = 1.0; fnt.paragraph_format.first_line_indent = Cm(0)
fnt.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
fnr = doc.styles.add_style('Footnote Reference', WD_STYLE_TYPE.CHARACTER)
fnr.font.superscript = True

# cabeçalho com número de página
hp = sec.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
def add_field(par, instr, size=10):
    r = par.add_run(); r.font.size = Pt(size)
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = ' '
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    for e in (f1, it, f2, t, f3): r._r.append(e)
add_field(hp, 'PAGE')

def add_inline(par, text, size=None, bold_all=False):
    # tokens: §FNn§, **negrito**, *itálico*
    for tok in re.split(r'(§FN\d+§|\*\*[^*]+\*\*|\*[^*]+\*)', text):
        if not tok:
            continue
        m = re.match(r'§FN(\d+)§', tok)
        if m:
            r = par.add_run(f'§FN{m.group(1)}§'); continue
        if tok.startswith('**'):
            r = par.add_run(tok[2:-2]); r.bold = True
        elif tok.startswith('*') and len(tok) > 2:
            r = par.add_run(tok[1:-1]); r.italic = True
            if bold_all: r.bold = True
        else:
            r = par.add_run(tok)
            if bold_all: r.bold = True
        if size: r.font.size = Pt(size)

def body_par(text):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(1.25)
    add_inline(p, text); return p

def heading(text, level, pagebreak=False, center=False):
    p = doc.add_paragraph()
    if pagebreak: p.paragraph_format.page_break_before = True
    p.paragraph_format.space_before = Pt(12 if level >= 2 else 0)
    p.paragraph_format.space_after = Pt(12); p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    add_inline(p, text.upper() if level == 1 else text, bold_all=True)
    # estilo de título para o sumário automático
    pPr = p._p.get_or_add_pPr(); ol = OxmlElement('w:outlineLvl'); ol.set(qn('w:val'), str(level - 1)); pPr.append(ol)
    return p

# folha de rosto
for t, b, sz in [('UNIVERSIDADE DO VALE DO ITAJAÍ – UNIVALI', True, 12),
                 ('PROGRAMA DE PÓS-GRADUAÇÃO STRICTO SENSU EM CIÊNCIA JURÍDICA – PPCJ', True, 12),
                 ('CURSO DE DOUTORADO EM CIÊNCIA JURÍDICA – CDCJ', True, 12), ('', False, 12), ('', False, 12),
                 ('DEFENSORIA PÚBLICA, RACISMO CLIMÁTICO E CIDADANIA EMANCIPATÓRIA RACIAL: A EDUCAÇÃO PARA A SUSTENTABILIDADE ÉTNICO-RACIAL COMO FUNÇÃO INSTITUCIONAL NO ESTADO DEMOCRÁTICO DE DIREITO SOCIOAMBIENTAL', True, 14),
                 ('', False, 12), ('Primeira minuta: Introdução e Capítulos 1 e 2', False, 12), ('', False, 12),
                 ('FÁBIO ROBERTO DE OLIVEIRA SANTOS', True, 12), ('', False, 12),
                 ('Orientadora: Professora Doutora Denise Schmitt Siqueira Garcia', False, 12), ('', False, 12), ('', False, 12),
                 ('Itajaí, 2026', False, 12)]:
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if t:
        r = p.add_run(t); r.bold = b; r.font.size = Pt(sz)

# rol de categorias (pré-textual) e sumário
if rol_blocks:
    for kind, t in rol_blocks:
        if kind == 'h1':
            heading(t, 1, pagebreak=True, center=True)
        else:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_after = Pt(6); add_inline(p, t)
p = heading('Sumário', 1, pagebreak=True, center=True)
p._p.get_or_add_pPr().remove(p._p.get_or_add_pPr().find(qn('w:outlineLvl')))
tp = doc.add_paragraph(); add_field(tp, 'TOC \\o "1-2" \\h \\z \\u', size=12)

for kind, t in blocks:
    if kind == 'h1':
        heading(t, 1, pagebreak=True, center=True)
    elif kind == 'chaptitle':
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(18); p.paragraph_format.keep_with_next = True
        add_inline(p, t.upper(), bold_all=True)
    elif kind == 'h2':
        heading(t, 2)
    elif kind == 'h3':
        heading(t, 3)
    elif kind == 'quote':
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        q = p.paragraph_format; q.left_indent = Cm(4); q.line_spacing = 1.0
        q.space_before = Pt(12); q.space_after = Pt(12)
        add_inline(p, t, size=10)
    else:
        body_par(t)

# referências
heading('Referências', 1, pagebreak=True, center=True)
def sk(e): return strip_acc(e.replace('«', '').replace('»', '')).upper()
LETTER_GROUPS = [('FREIRE', '2023'), ('MOREIRA', '2019'), ('MBEMBE', '2018'), ('BRASIL', '2004')]
letter_of = {}
for sur, yr in LETTER_GROUPS:
    grp = sorted([i for i in cited_idx if strip_acc(bib_plain[i]).upper().startswith(sur) and
                  yr in bib_plain[i].split('Disponível')[0]], key=lambda i: sk(bib[i]))
    if sur == 'BRASIL':
        grp = [i for i in grp if 'CNE/CP' in bib_plain[i]]
    if len(grp) > 1:
        for k_, i in enumerate(grp): letter_of[i] = (yr, 'abcdefgh'[k_])
def add_letter(e, i):
    if i not in letter_of: return e
    yr, L = letter_of[i]
    head, sep, tail = e.partition('Disponível')
    pos = head.rfind(yr)
    return head[:pos] + yr + L + head[pos + len(yr):] + sep + tail
for i in sorted(cited_idx, key=lambda i: sk(bib[i])):
    bib[i] = add_letter(bib[i], i)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(12)
    for part in re.split(r'(«[^»]*»)', bib[i]):
        if part:
            r = p.add_run(part.strip('«»')); r.bold = part.startswith('«')

tmp = tempfile.mkdtemp()
base_docx = os.path.join(tmp, 'base.docx'); doc.save(base_docx)

# ---------- 4) pós-processamento: notas de rodapé reais ----------
ext = os.path.join(tmp, 'x'); os.makedirs(ext)
with zipfile.ZipFile(base_docx) as z: z.extractall(ext)
dpath = os.path.join(ext, 'word', 'document.xml')
dt = etree.parse(dpath); droot = dt.getroot()
def wq(t): return f'{{{W}}}{t}'
n_refs = 0
for r in list(droot.iter(wq('r'))):
    t = r.find(wq('t'))
    if t is not None and t.text and re.fullmatch(r'§FN\d+§', t.text):
        n = re.fullmatch(r'§FN(\d+)§', t.text).group(1)
        r.remove(t)
        rpr = r.find(wq('rPr'))
        if rpr is None:
            rpr = etree.SubElement(r, wq('rPr')); r.insert(0, rpr)
        rs = etree.SubElement(rpr, wq('rStyle')); rs.set(wq('val'), 'FootnoteReference'); rpr.insert(0, rs)
        fr = etree.SubElement(r, wq('footnoteReference')); fr.set(wq('id'), n)
        n_refs += 1
dt.write(dpath, xml_declaration=True, encoding='UTF-8', standalone=True)

fnroot = etree.Element(wq('footnotes'), nsmap={'w': W})
for i, typ in ((-1, 'separator'), (0, 'continuationSeparator')):
    fn = etree.SubElement(fnroot, wq('footnote')); fn.set(wq('type'), typ); fn.set(wq('id'), str(i))
    p = etree.SubElement(fn, wq('p')); r = etree.SubElement(p, wq('r'))
    etree.SubElement(r, wq('separator' if i == -1 else 'continuationSeparator'))
def fn_runs(p, text):
    for tok in re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text):
        if not tok: continue
        r = etree.SubElement(p, wq('r')); rpr = etree.SubElement(r, wq('rPr'))
        if tok.startswith('**'): etree.SubElement(rpr, wq('b')); tok = tok[2:-2]
        elif tok.startswith('*') and len(tok) > 2: etree.SubElement(rpr, wq('i')); tok = tok[1:-1]
        t = etree.SubElement(r, wq('t')); t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); t.text = tok
for g in sorted(notes):
    fn = etree.SubElement(fnroot, wq('footnote')); fn.set(wq('id'), str(g))
    p = etree.SubElement(fn, wq('p')); ppr = etree.SubElement(p, wq('pPr'))
    ps = etree.SubElement(ppr, wq('pStyle')); ps.set(wq('val'), 'FootnoteText')
    r = etree.SubElement(p, wq('r')); rpr = etree.SubElement(r, wq('rPr'))
    rs = etree.SubElement(rpr, wq('rStyle')); rs.set(wq('val'), 'FootnoteReference')
    etree.SubElement(r, wq('footnoteRef'))
    r2 = etree.SubElement(p, wq('r')); t = etree.SubElement(r2, wq('t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); t.text = ' '
    fn_runs(p, notes[g])
etree.ElementTree(fnroot).write(os.path.join(ext, 'word', 'footnotes.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)

rp = os.path.join(ext, 'word', '_rels', 'document.xml.rels'); rt = etree.parse(rp)
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
ids = [e.get('Id') for e in rt.getroot()]; k = 1
while f'rId{k}' in ids: k += 1
e = etree.SubElement(rt.getroot(), f'{{{PR}}}Relationship'); e.set('Id', f'rId{k}')
e.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'); e.set('Target', 'footnotes.xml')
rt.write(rp, xml_declaration=True, encoding='UTF-8', standalone=True)
cp = os.path.join(ext, '[Content_Types].xml'); ct = etree.parse(cp)
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
o = etree.SubElement(ct.getroot(), f'{{{CT}}}Override'); o.set('PartName', '/word/footnotes.xml')
o.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml')
ct.write(cp, xml_declaration=True, encoding='UTF-8', standalone=True)
sp = os.path.join(ext, 'word', 'settings.xml'); stt = etree.parse(sp); sr = stt.getroot()
fpr = etree.Element(wq('footnotePr'))
for i in (-1, 0):
    x = etree.SubElement(fpr, wq('footnote')); x.set(wq('id'), str(i))
sr.insert(0, fpr)
uf = etree.Element(wq('updateFields')); uf.set(wq('val'), 'true'); sr.insert(0, uf)
stt.write(sp, xml_declaration=True, encoding='UTF-8', standalone=True)

if os.path.exists(out_docx): os.remove(out_docx)
with zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(cp, '[Content_Types].xml')
    for rd, _, fs in os.walk(ext):
        for fnm in fs:
            fp = os.path.join(rd, fnm); arc = os.path.relpath(fp, ext)
            if arc != '[Content_Types].xml': z.write(fp, arc)
shutil.rmtree(tmp)

words = sum(len(re.sub(r'§FN\d+§', '', t).split()) for k_, t in blocks if k_ in ('p', 'quote'))
json.dump({'arquivos': [os.path.basename(f) for f in files], 'notas': len(notes), 'chamadas_de_nota': n_refs,
           'palavras_corpo': words, 'obras_citadas': len(cited_idx), 'citacoes_sem_correspondencia': unmatched, 'letras': {bib_plain[i][:60]: ''.join(v) for i, v in letter_of.items()}},
          open(rep_json, 'w'), ensure_ascii=False, indent=1)
print(json.dumps({'notas': len(notes), 'palavras_corpo': words, 'obras_citadas': len(cited_idx),
                  'sem_correspondencia': len(unmatched)}, ensure_ascii=False))
