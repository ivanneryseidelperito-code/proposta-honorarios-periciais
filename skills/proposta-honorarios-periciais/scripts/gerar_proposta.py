#!/usr/bin/env python3
"""Preenche o modelo tipográfico de proposta de honorários do perito.

Uso: python3 gerar_proposta.py dados.json saida.docx [--modelo caminho.docx]

Campos do JSON (ver exemplo_dados.json):
  juizo, cidade, processo, parte1_qualif, parte1_nome, parte2_qualif, parte2_nome,
  cpf, id_nomeacao, objeto, data (opcional; padrão = hoje),
  vh {valor, fonte_nome, fonte_url, nota}         -> obrigatórios valor, fonte_nome, fonte_url
  etapas [{descricao, horas, atividades}]          -> um parágrafo de planejamento e uma linha do quadro por etapa
  prazo_dias (prazo estimado para o laudo, em dias)
  valor_final (opcional; gera a linha Desconto/arredondamento)
  ressalva_despesas (bool; padrão true)
"""
import sys, os, json, re, zipfile, copy, datetime
from decimal import Decimal, ROUND_HALF_UP
from xml.sax.saxutils import escape
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
q = lambda t: "{%s}%s" % (W, t)
AQUI = os.path.dirname(os.path.abspath(__file__))
MODELO = os.path.join(AQUI, "..", "modelo", "modelo_proposta_honorarios.docx")
MESES = "janeiro fevereiro março abril maio junho julho agosto setembro outubro novembro dezembro".split()
RESSALVA = (" Igualmente, não estão inclusas as despesas com viagens, hospedagens, transporte, "
            "alimentação e outras eventualmente necessárias, cujo ressarcimento será objeto de "
            "pedido formal, nos termos da NBC PP 01 (R2), item 35.")

D = lambda v: Decimal(str(v)).quantize(Decimal("0.01"), ROUND_HALF_UP)

def brl(v, prefixo=True):
    s = f"{D(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("R$ " + s) if prefixo else s

UN = "zero um dois três quatro cinco seis sete oito nove dez onze doze treze quatorze quinze dezesseis dezessete dezoito dezenove".split()
DZ = "_ _ vinte trinta quarenta cinquenta sessenta setenta oitenta noventa".split()
CT = "_ cento duzentos trezentos quatrocentos quinhentos seiscentos setecentos oitocentos novecentos".split()

def _ate999(n):
    if n == 100: return "cem"
    p = []
    c, r = divmod(n, 100)
    if c: p.append(CT[c])
    if r < 20 and r: p.append(UN[r])
    elif r:
        d, u = divmod(r, 10); p.append(DZ[d] + (" e " + UN[u] if u else ""))
    return " e ".join(p)

def _inteiro(n):
    if n == 0: return "zero"
    grupos = []; esc = [("", ""), ("mil", "mil"), ("milhão", "milhões"), ("bilhão", "bilhões")]
    i = 0
    while n:
        n, g = divmod(n, 1000)
        if g:
            s = _ate999(g)
            if i == 1: s = "mil" if g == 1 else s + " mil"
            elif i > 1: s = s + " " + (esc[i][0] if g == 1 else esc[i][1])
            grupos.insert(0, (g, s))
        i += 1
    out = grupos[0][1]
    for g, s in grupos[1:]:
        out += (" e " if (g < 100 or g % 100 == 0) else " ") + s
    return out

def extenso(v):
    v = D(v); r = int(v); c = int((v - r) * 100)
    partes = []
    if r:
        txt = _inteiro(r)
        de = " de" if r % 1000000 == 0 and r >= 1000000 else ""
        partes.append(txt + de + (" real" if r == 1 else " reais"))
    if c:
        partes.append(_inteiro(c) + (" centavo" if c == 1 else " centavos"))
    s = " e ".join(partes) or "zero real"
    return s[0].upper() + s[1:]

def data_extenso(d):
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"

def sub_tokens_xml(root, mapa):
    for t in root.iter(q("t")):
        if t.text and "{{" in t.text:
            for k, v in mapa.items():
                t.text = t.text.replace("{{%s}}" % k, v)

def main():
    a = sys.argv[1:]
    modelo = MODELO
    if "--modelo" in a:
        i = a.index("--modelo"); modelo = a[i + 1]; del a[i:i + 2]
    dados = json.load(open(a[0], encoding="utf-8")); saida = a[1]

    vh = dados.get("vh") or {}
    falt = [k for k in ("valor", "fonte_nome", "fonte_url") if not vh.get(k)]
    if falt:
        sys.exit("ERRO: valor-hora exige valor, fonte_nome e fonte_url informados pelo perito. Faltando: " + ", ".join(falt))
    VH = D(vh["valor"])
    etapas = [(e["descricao"], D(e["horas"])) for e in dados["etapas"]]
    linhas = [(d, h, D(h * VH)) for d, h in etapas]
    subtotal = sum((x[2] for x in linhas), Decimal("0"))
    total = D(dados["valor_final"]) if dados.get("valor_final") not in (None, "") else subtotal
    desconto = subtotal - total
    if desconto < 0:
        sys.exit("ERRO: valor_final superior ao subtotal calculado (%s)." % brl(subtotal))
    data = dados.get("data") or data_extenso(datetime.date.today())

    mapa = {
        "JUIZO": dados["juizo"], "CIDADE": dados["cidade"], "PROCESSO": dados["processo"],
        "PARTE1_QUALIF": dados.get("parte1_qualif", "Autor:"), "PARTE1_NOME": dados["parte1_nome"],
        "PARTE2_QUALIF": dados.get("parte2_qualif", "Réu:"), "PARTE2_NOME": dados["parte2_nome"],
        "CPF": dados.get("cpf") or "[CPF]", "ID_NOMEACAO": dados["id_nomeacao"], "OBJETO": dados["objeto"],
        "FONTE_VH_NOME": vh["fonte_nome"], "FONTE_VH_URL": vh["fonte_url"],
        "NOTA_VH": " " + (vh.get("nota") or f"Valor da hora técnica adotado: {brl(VH)}, conforme {vh['fonte_nome']} ({vh['fonte_url']})."),
        "TOTAL": brl(total), "TOTAL_TAB": brl(total), "TOTAL_EXTENSO": extenso(total),
        "SUBTOTAL": brl(subtotal), "DESCONTO": brl(desconto), "DATA": data,
        "PRAZO_DIAS": str(dados.get("prazo_dias") or "[A CONFIRMAR]"),
        "RESSALVA_DESPESAS": RESSALVA if dados.get("ressalva_despesas", True) else "",
    }

    z = zipfile.ZipFile(modelo); parts = {n: z.read(n) for n in z.namelist()}
    doc = etree.fromstring(parts["word/document.xml"])
    # tabela de etapas
    for t in doc.iter(q("tbl")):
        txt = "".join(x.text or "" for x in t.iter(q("t")))
        if "{{E_DESC}}" in txt:
            rows = t.findall(q("tr")); proto = rows[1]; pos = t.index(proto)
            for k, (d, h, v) in enumerate(linhas):
                r = copy.deepcopy(proto)
                sub_tokens_xml(r, {"E_DESC": d, "E_HORAS": brl(h, False), "E_VH": brl(VH), "E_VALOR": brl(v)})
                t.insert(pos + k, r)
            t.remove(proto)
            if desconto == 0:
                for r in t.findall(q("tr")):
                    if "{{DESCONTO}}" in "".join(x.text or "" for x in r.iter(q("t"))):
                        t.remove(r)
    # bloco de planejamento: um parágrafo por etapa (título em negrito + atividades)
    for proto in [p for p in doc.iter(q("p")) if "{{PLAN_TITULO}}" in "".join(t.text or "" for t in p.iter(q("t")))]:
        blank = proto.getprevious(); ant = proto
        for n, e in enumerate(dados["etapas"], 1):
            if not e.get("atividades"):
                sys.exit("ERRO: etapa '%s' sem descrição das atividades (campo atividades)." % e["descricao"])
            novo = copy.deepcopy(proto)
            sub_tokens_xml(novo, {"PLAN_TITULO": "Etapa %d – %s:" % (n, e["descricao"]), "PLAN_DESC": e["atividades"].strip()})
            if n > 1:
                b = copy.deepcopy(blank); ant.addnext(b); ant = b
            ant.addnext(novo); ant = novo
        proto.getparent().remove(proto)
    sub_tokens_xml(doc, mapa)
    parts["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    fn = etree.fromstring(parts["word/footnotes.xml"]); sub_tokens_xml(fn, mapa)
    parts["word/footnotes.xml"] = etree.tostring(fn, xml_declaration=True, encoding="UTF-8", standalone=True)
    for n in ("docProps/core.xml", "docProps/app.xml"):
        s = parts[n].decode("utf-8")
        for k, v in mapa.items(): s = s.replace("{{%s}}" % k, escape(v))
        parts[n] = s.encode("utf-8")
    rest = [m for n in parts if n.endswith(".xml") for m in re.findall(r"\{\{[A-Z0-9_]+\}\}", parts[n].decode("utf-8", "ignore"))]
    if rest:
        sys.exit("ERRO: marcadores não preenchidos: " + ", ".join(sorted(set(rest))))
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as o:
        for n in z.namelist(): o.writestr(n, parts[n])
    print(json.dumps({"subtotal": brl(subtotal), "desconto": brl(desconto), "total": brl(total),
                      "horas": brl(sum(h for _, h, _ in linhas), False), "extenso": extenso(total)}, ensure_ascii=False))

if __name__ == "__main__":
    main()
