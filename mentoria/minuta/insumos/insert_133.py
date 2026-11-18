# -*- coding: utf-8 -*-
"""Insere as 133 referências do relatório na seção 9 (bibliografia adicional),
em ordem alfabética, como marcas de revisão de 'Fábio Portela', Arial 12,
destaque ABNT em negrito; pendências de verificação viram comentários do Word."""
import sys, re, json
sys.path.insert(0, "/tmp/claude-0/-home-user-AtheNEO/9146d87c-7f84-54b9-8b57-723cee0e44d2/scratchpad/tools")
from docx_track import Doc, sortkey, write_comments_part

report_path, base = sys.argv[1], sys.argv[2]
lines = open(report_path, encoding='utf-8').read().split('\n')

# ---- 1) parse das 133 entradas ----
items, cur = {}, None
for ln in lines:
    if ln.startswith('## Itens que exigem'):
        break
    m = re.match(r'^(\d{1,3})\. (.*)$', ln)
    if m:
        cur = int(m.group(1)); items[cur] = {'ref': m.group(2), 'extra': []}
    elif cur and ln.strip():
        items[cur]['extra'].append(ln)
assert sorted(items) == list(range(1, 134)), "numeração inesperada"

FLAG = re.compile(r'\s*\[dado não confirmado:\s*([^\]]*)\]')
for n, it in items.items():
    flags = FLAG.findall(it['ref']) + [f for e in it['extra'] for f in FLAG.findall(e)]
    ref = FLAG.sub('', it['ref']).strip()
    assert ref.count('**') == 2, f"negrito inesperado no item {n}"
    ref = re.sub(r'\*\*(.+?)\*\*', r'«\1»', ref)
    ref = re.sub(r'\s{2,}', ' ', ref)
    if not ref.endswith('.'):
        ref += '.'
    it['marked'], it['flags'] = ref, flags

# ---- 2) categorias da seção "Itens que exigem conferência" ----
CAT = {
 "edição, ano ou ISBN divergentes": [1,3,9,10,22,28,32,80,90,91,93,95,102,97],
 "subtítulo, local ou série a conferir na folha de rosto": [14,26,43,59,67,68,73,79,81,82,83,86,118],
 "DOI, páginas ou volume não conferidos": [2,33,35,38,47,49,52,55,63,70,96],
 "URL deduzida, a testar": [25],
 "número de páginas divergente": [54,56,57,60,62],
 "tradutor ou coautoria não confirmados": [40,66,92,131],
 "coletânea não confirmada (item mais frágil da lista)": [88],
 "dados de publicação no DOU a conferir": [75,77,78],
 "confirmação só por fonte secundária ou resultado de busca": [50,71,98,99,100,104,106,108,118],
 "conferir o percentual de 69,3% no PDF de 2022 antes de citar": [103],
 "edição brasileira inexistente ou não confirmada (cita-se o original)": [34,120,127,129],
 "dados da editora obtidos por resultado de busca (site não abriu)": [5,11,13,14,15,23],
}
for n, it in items.items():
    cats = [c for c, nums in CAT.items() if n in nums]
    it['comment'] = None
    if cats or it['flags']:
        parts = []
        if cats:
            parts.append("Conferir antes da versão final: " + "; ".join(cats) + ".")
        if it['flags']:
            parts.append("Pendência registrada na pesquisa: " + " | ".join(f.strip().rstrip('.') for f in it['flags']) + ".")
        it['comment'] = " ".join(parts)

# ---- 3) inserção alfabética na seção 9 ----
d = Doc(f"{base}/word/document.xml")
existing_keys = set()
for p in d.paras:
    t = d.ptext(p).strip()
    if t:
        existing_keys.add(sortkey(t)[:80])
dups = [n for n, it in items.items() if sortkey(it['marked'])[:80] in existing_keys]
assert not dups, f"duplicatas com o documento: {dups}"

def sec9_entries():
    s9 = d.find_para("9 BIBLIOGRAFIA A PESQUISAR", start=400, exact=True)
    return [(sortkey(d.ptext(d.paras[i])), d.paras[i]) for i in range(s9 + 1, len(d.paras))
            if d.ptext(d.paras[i]).strip()]

ent0 = sec9_entries()
model, blank_model = ent0[0][1], d.following_blank(ent0[0][1])
for n in sorted(items, key=lambda k: sortkey(items[k]['marked'])):
    it = items[n]
    key = sortkey(it['marked'])
    ents = sec9_entries()
    new = d.new_bib_paragraphs(it['marked'], model, blank_model, comment=it['comment'])
    cands = [(k, p) for k, p in ents if k <= key]
    if cands:
        pred = max(cands, key=lambda x: x[0])[1]
        b = d.following_blank(pred)
        d.insert_after(b if b is not None else pred, new)
    else:
        d.insert_before(ents[0][1], new)

d.save()
write_comments_part(base, d.comments)
print(json.dumps({"inseridas": len(items), "comentarios": len(d.comments)}, ensure_ascii=False))
