# -*- coding: utf-8 -*-
"""Inserção rastreada (w:ins) de entradas bibliográficas e comentários do Word,
autor 'Fábio Portela'. Arial 12; trecho entre «...» em negrito."""
import copy, re, unicodedata
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def w(t): return f'{{{W}}}{t}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
AUTHOR, INITIALS = "Fábio Portela", "FP"
DATE = "2026-11-17T12:00:00Z"

def norm(s):
    return (s or "").replace(" ", " ").replace("‑", "-")

def sortkey(t):
    t = norm(t).replace('«', '').replace('»', '')
    t = unicodedata.normalize('NFKD', t)
    return "".join(c for c in t if not unicodedata.combining(c)).strip().upper()

class Doc:
    def __init__(self, path):
        self.path = path
        self.tree = etree.parse(path)
        self.body = self.tree.getroot().find(w('body'))
        ids = [int(v) for el in self.tree.iter() for k, v in el.attrib.items()
               if k == w('id') and v.isdigit()]
        self.next_id = (max(ids) if ids else 0) + 1000
        self.comments = []  # (id, text)
        self.refresh()

    def nid(self):
        self.next_id += 1
        return str(self.next_id)

    def refresh(self):
        self.paras = self.body.findall(w('p'))

    def ptext(self, p):  # visão aceita aproximada: w:t (delText fica de fora)
        return "".join(t.text or "" for t in p.iter(w('t')))

    def find_para(self, phrase, start=0, exact=False):
        n = norm(phrase)
        for i in range(start, len(self.paras)):
            t = norm(self.ptext(self.paras[i]))
            if (t.strip() == n) if exact else (n in t):
                return i
        raise ValueError(f"não encontrado: {phrase!r}")

    def following_blank(self, p):
        nx = p.getnext()
        if nx is not None and nx.tag == w('p') and self.ptext(nx).strip() == "":
            return nx
        return None

    def _mark(self, tag):
        e = etree.Element(w(tag))
        e.set(w('id'), self.nid()); e.set(w('author'), AUTHOR); e.set(w('date'), DATE)
        return e

    def _mark_para_inserted(self, p):
        ppr = p.find(w('pPr'))
        if ppr is None:
            ppr = etree.Element(w('pPr')); p.insert(0, ppr)
        rpr = ppr.find(w('rPr'))
        if rpr is None:
            rpr = etree.SubElement(ppr, w('rPr'))
        rpr.insert(0, self._mark('ins'))

    @staticmethod
    def _rpr(bold):
        rpr = etree.Element(w('rPr'))
        rf = etree.SubElement(rpr, w('rFonts'))
        for a in ('ascii', 'eastAsia', 'hAnsi', 'cs'):
            rf.set(w(a), 'Arial')
        if bold:
            etree.SubElement(rpr, w('b')); etree.SubElement(rpr, w('bCs'))
        etree.SubElement(rpr, w('sz')).set(w('val'), '24')
        etree.SubElement(rpr, w('szCs')).set(w('val'), '24')
        return rpr

    @staticmethod
    def segments(marked):
        out = []
        for part in re.split(r'(«[^»]*»)', marked):
            if part:
                if part.startswith('«'):
                    out.append((part[1:-1], True))
                else:
                    out.append((part, False))
        return out

    def new_bib_paragraphs(self, marked, entry_model, blank_model, comment=None):
        ep = etree.Element(w('p'))
        ppr = entry_model.find(w('pPr'))
        ep.append(copy.deepcopy(ppr) if ppr is not None else etree.Element(w('pPr')))
        self._mark_para_inserted(ep)
        cid = None
        if comment:
            cid = self.nid()
            crs = etree.SubElement(ep, w('commentRangeStart')); crs.set(w('id'), cid)
        ins = self._mark('ins')
        for text, bold in self.segments(marked):
            r = etree.SubElement(ins, w('r'))
            r.append(self._rpr(bold))
            t = etree.SubElement(r, w('t')); t.set(XS, 'preserve'); t.text = text
        ep.append(ins)
        if comment:
            cre = etree.SubElement(ep, w('commentRangeEnd')); cre.set(w('id'), cid)
            ins2 = self._mark('ins')
            rr = etree.SubElement(ins2, w('r'))
            etree.SubElement(rr, w('commentReference')).set(w('id'), cid)
            ep.append(ins2)
            self.comments.append((cid, comment))
        bp = etree.Element(w('p'))
        bppr = blank_model.find(w('pPr'))
        bp.append(copy.deepcopy(bppr) if bppr is not None else etree.Element(w('pPr')))
        self._mark_para_inserted(bp)
        return [ep, bp]

    def insert_after(self, anchor, elems):
        cur = anchor
        for e in elems:
            cur.addnext(e); cur = e
        self.refresh()

    def insert_before(self, anchor, elems):
        for e in elems:
            anchor.addprevious(e)
        self.refresh()

    def save(self):
        self.tree.write(self.path, xml_declaration=True, encoding='UTF-8', standalone=True)

def write_comments_part(root_dir, comments):
    """Cria/estende word/comments.xml e registra rels e content types."""
    import os
    cpath = os.path.join(root_dir, 'word', 'comments.xml')
    if os.path.exists(cpath):
        ctree = etree.parse(cpath); croot = ctree.getroot()
    else:
        croot = etree.Element(w('comments'), nsmap={'w': W}); ctree = etree.ElementTree(croot)
    for cid, text in comments:
        c = etree.SubElement(croot, w('comment'))
        c.set(w('id'), cid); c.set(w('author'), AUTHOR); c.set(w('date'), DATE); c.set(w('initials'), INITIALS)
        p = etree.SubElement(c, w('p'))
        r0 = etree.SubElement(p, w('r'))
        etree.SubElement(r0, w('annotationRef'))
        r = etree.SubElement(p, w('r'))
        t = etree.SubElement(r, w('t')); t.set(XS, 'preserve'); t.text = text
    ctree.write(cpath, xml_declaration=True, encoding='UTF-8', standalone=True)
    # rels
    rp = os.path.join(root_dir, 'word', '_rels', 'document.xml.rels')
    rt = etree.parse(rp); rr = rt.getroot()
    REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
    if not any(e.get('Type') == REL for e in rr):
        ids = [e.get('Id') for e in rr]
        n = 1
        while f'rId{n}' in ids: n += 1
        e = etree.SubElement(rr, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
        e.set('Id', f'rId{n}'); e.set('Type', REL); e.set('Target', 'comments.xml')
        rt.write(rp, xml_declaration=True, encoding='UTF-8', standalone=True)
    # content types
    cp = os.path.join(root_dir, '[Content_Types].xml')
    ct = etree.parse(cp); cr = ct.getroot()
    CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    if not any(e.get('PartName') == '/word/comments.xml' for e in cr):
        e = etree.SubElement(cr, '{http://schemas.openxmlformats.org/package/2006/content-types}Override')
        e.set('PartName', '/word/comments.xml'); e.set('ContentType', CT)
        ct.write(cp, xml_declaration=True, encoding='UTF-8', standalone=True)
