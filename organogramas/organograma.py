"""Gera organogramas de times no padrão visual do modelo (modelos/Organograma_Time_Youcom.pptx).

Uso: a partir de um script de time (ex.: renner.py), chame build(slides, saida).
Cada slide é um dict:
    title, subtitle        textos do cabeçalho
    top                    (nome, papel) bloco laranja de direção (opcional)
    gray                   [(nome, papel), ...] blocos cinza logo abaixo do topo (opcional)
    columns                [{title, color, weight, items}, ...]
    columns_width          fração da largura útil ocupada pelas colunas, centralizadas (padrão 1)
    legend                 lista com "line", "hiring" e/ou texto livre de rodapé
Itens de coluna:
    ("head", nome, papel)                   cartão de liderança, largura total da coluna
    ("grid", ncols, [(nome, papel[, tipo])]) cartões de equipe; tipo "member" | "hiring" | "summary"
Cartões "head" consecutivos são ligados por uma linha (responde a).
"""
import math
import os
import re
import shutil
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "modelos", "Organograma_Time_Youcom.pptx")

ORANGE, RED, BLACK, WINE, GRAY = "EC6A3B", "B8393A", "141414", "7C2722", "5E5E5E"
GRAY_BG, MUTED, BORDER, BG = "EFEEEC", "5E5E5E", "E3E1DE", "FCFBF9"
FONT = ('<a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/>'
        '<a:cs typeface="Arial" pitchFamily="34" charset="-120"/>')

EMU = 914400
L, R = 812800, 11379200          # margens do modelo
COL_GAP, SUB_GAP, ROW_GAP = 152400, 76200, 63500
LINK_GAP = 152400                # espaço vertical para a linha entre lideranças
CONTENT_BOTTOM = 6235700         # acima da legenda
INS_L, INS_T, INS_R, INS_B = 50800, 76200, 101600, 50800


# ---------------------------------------------------------------- medição de texto (estimativa conservadora)
def _wrap_lines(text, sz_pt, width_emu, bold):
    per_char = sz_pt / 72 * (0.66 if bold else 0.52)  # polegadas por caractere (Arial)
    usable = (width_emu - INS_L - INS_R) / EMU
    lines, cur = 1, 0.0
    for word in text.split(" "):
        w = len(word) * per_char
        if cur == 0:
            cur = w
        elif cur + per_char + w <= usable:
            cur += per_char + w
        else:
            lines += 1
            cur = w
    return lines


def card_height(w, name, role, name_sz=1200):
    n = _wrap_lines(name, name_sz / 100, w, True) + _wrap_lines(role, 12, w, False)
    return int((0.15 + n * 0.195 + 0.03) * EMU)


# ---------------------------------------------------------------- primitivas de XML (mesmas do modelo)
class Slide:
    def __init__(self):
        self.shapes, self._id = [], 4

    def nid(self):
        self._id += 1
        return self._id

    @staticmethod
    def _xfrm(x, y, w, h):
        return f'<a:xfrm><a:off x="{int(x)}" y="{int(y)}"/><a:ext cx="{int(w)}" cy="{int(h)}"/></a:xfrm>'

    def rect(self, x, y, w, h, fill, line=None, lw=19050, dash="solid"):
        i = self.nid()
        ln = (f'<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill><a:prstDash val="{dash}"/></a:ln>'
              if line else "<a:ln></a:ln>")
        self.shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Shape {i}"></p:cNvPr><p:cNvSpPr/><p:nvPr></p:nvPr></p:nvSpPr>'
            f'<p:spPr>{self._xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}</p:spPr></p:sp>')

    def line(self, x1, y1, x2, y2):
        i = self.nid()
        self.shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Shape {i}"></p:cNvPr><p:cNvSpPr/><p:nvPr></p:nvPr></p:nvSpPr>'
            f'<p:spPr>{self._xfrm(x1, y1, x2 - x1, y2 - y1)}<a:prstGeom prst="line"><a:avLst></a:avLst></a:prstGeom>'
            f'<a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="{BLACK}"/></a:solidFill>'
            f'<a:prstDash val="solid"/></a:ln></p:spPr></p:sp>')

    def text(self, x, y, w, h, paras, ins=(INS_L, INS_T, INS_R, INS_B), fill=None, algn="ctr", spc=True):
        """paras: [(texto, sz, cor, negrito)]"""
        i = self.nid()
        fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
        body = ""
        for t, sz, color, bold in paras:
            sp = '<a:spcAft><a:spcPts val="200"/></a:spcAft>' if spc else ""
            b = ' b="1"' if bold else ""
            body += (f'<a:p><a:pPr algn="{algn}" indent="0" marL="0">{sp}<a:buNone/></a:pPr>'
                     f'<a:r><a:rPr lang="pt-BR" sz="{sz}"{b} dirty="0"><a:solidFill><a:srgbClr val="{color}"/>'
                     f'</a:solidFill>{FONT}</a:rPr><a:t>{escape(t)}</a:t></a:r>'
                     f'<a:endParaRPr lang="pt-BR" sz="{sz}" dirty="0"/></a:p>')
        self.shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Text {i}"></p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr></p:nvPr></p:nvSpPr>'
            f'<p:spPr>{self._xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom>{fill_xml}'
            f'<a:ln></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="{ins[0]}" tIns="{ins[1]}" rIns="{ins[2]}" '
            f'bIns="{ins[3]}" rtlCol="0" anchor="ctr"></a:bodyPr><a:lstStyle/>{body}</p:txBody></p:sp>')

    def card(self, x, y, w, h, name, role, kind="member", accent=ORANGE, name_sz=1200):
        algn, name_color, role_color = "l", BLACK, MUTED
        if kind == "member":
            self.rect(x, y, w, h, "FFFFFF", BORDER, 12700)
        elif kind == "summary":
            self.rect(x, y, w, h, GRAY_BG, BORDER, 12700)
        elif kind == "hiring":
            self.rect(x, y, w, h, "FFFFFF", ORANGE, 19050, "dash")
        elif kind == "head":
            self.rect(x, y, w, h, "FFFFFF", accent, 19050)
            algn = "ctr"
        elif kind == "top":
            self.rect(x, y, w, h, ORANGE)
            algn, name_color, role_color = "ctr", "FFFFFF", "FFFFFF"
        elif kind == "gray":
            self.rect(x, y, w, h, GRAY_BG)
            algn = "ctr"
        self.text(x, y, w, h, [(name, name_sz, name_color, True), (role, 1200, role_color, False)], algn=algn)


# ---------------------------------------------------------------- layout de um slide
def _render(spec, header_xml):
    s = Slide()
    s.shapes.extend(header_xml)
    width = R - L
    cx = L + width // 2

    # bloco de direção + caixas cinza
    y = 1168400
    anchor_bottom = None
    if spec.get("top"):
        name, role = spec["top"]
        w = max(3302000, min(6096000, int((max(len(name) * 13 * 0.66, len(role) * 12 * 0.52) / 72 + 0.4) * EMU)))
        h = 558800
        s.card(cx - w // 2, y, w, h, name, role, kind="top", name_sz=1300)
        anchor_bottom = y + h
        y += h + 101600
        gray = spec.get("gray", [])
        if len(gray) == 1:
            name, role = gray[0]
            gw = max(3302000, int((max(len(name) * 13 * 0.66, len(role) * 12 * 0.52) / 72 + 0.4) * EMU))
            s.line(cx, anchor_bottom, cx, y)
            s.card(cx - gw // 2, y, gw, h, name, role, kind="gray", name_sz=1300)
            anchor_bottom = y + h
            y += h + 101600
        elif len(gray) == 2:
            gw = 3302000
            s.card(L, y, gw, h, *gray[0], kind="gray", name_sz=1300)
            s.card(R - gw, y, gw, h, *gray[1], kind="gray", name_sz=1300)
            s.line(L + gw, y + h // 2, R - gw, y + h // 2)
            y += h + 101600
        bus_y = y + 25400
        y = bus_y + 127000
    else:
        y = 1295400

    # colunas
    cols = spec["columns"]
    area_w = int(width * spec.get("columns_width", 1))
    total_w = area_w - COL_GAP * (len(cols) - 1)
    weights = sum(c.get("weight", 1) for c in cols)
    xs, x = [], cx - area_w // 2
    for c in cols:
        w = int(total_w * c.get("weight", 1) / weights)
        xs.append((x, w))
        x += w + COL_GAP

    hdr_sz = 1600
    while hdr_sz > 1200 and any(_wrap_lines(c["title"], hdr_sz / 100, w, True) > 1 for c, (_, w) in zip(cols, xs)):
        hdr_sz -= 100
    two = any(_wrap_lines(c["title"], hdr_sz / 100, w, True) > 1 for c, (_, w) in zip(cols, xs))
    hdr_h = 594360 if two else 419100

    if spec.get("top"):
        s.line(cx, anchor_bottom, cx, bus_y)
        centers = [x + w // 2 for x, w in xs]
        s.line(min(centers + [cx]), bus_y, max(centers + [cx]), bus_y)
        for c in centers:
            s.line(c, bus_y, c, y)

    bottoms = []
    for c, (x, w) in zip(cols, xs):
        s.text(x, y, w, hdr_h, [(c["title"], hdr_sz, "FFFFFF", True)], ins=(50800,) * 4, fill=c["color"], spc=False)
        cy = y + hdr_h + 76200
        prev_head = None
        for item in c["items"]:
            if item[0] == "head":
                _, name, role = item
                if prev_head is not None:
                    s.line(x + w // 2, prev_head, x + w // 2, cy + LINK_GAP - ROW_GAP)
                    cy += LINK_GAP - ROW_GAP
                h = card_height(w, name, role)
                s.card(x, cy, w, h, name, role, kind="head", accent=c["color"])
                prev_head = cy + h
                cy += h + ROW_GAP
            elif item[0] == "grid":
                _, ncols, people = item
                prev_head = None
                sw = (w - SUB_GAP * (ncols - 1)) // ncols
                for r in range(math.ceil(len(people) / ncols)):
                    row = people[r * ncols:(r + 1) * ncols]
                    h = max(card_height(sw, p[0], p[1]) for p in row)
                    for k, p in enumerate(row):
                        s.card(x + k * (sw + SUB_GAP), cy, sw, h, p[0], p[1], kind=p[2] if len(p) > 2 else "member")
                    cy += h + ROW_GAP
        bottoms.append(cy)

    # legenda (canto inferior direito, como no modelo)
    legend = spec.get("legend", ["line"])
    lx, ly = R, 6311900
    for item in reversed(legend):
        if item == "hiring":
            s.text(lx - 1270000, ly, 1270000, 228600, [("Em contratação", 1200, BLACK, False)], ins=(0,) * 4, algn="l", spc=False)
            s.rect(lx - 1270000 - 355600, ly + 31750, 254000, 165100, BG, ORANGE, 19050, "dash")
            lx -= 1270000 + 355600 + 304800
        elif item == "line":
            s.text(lx - 1905000, ly, 1905000, 228600, [("Liderança (responde a)", 1200, BLACK, False)], ins=(0,) * 4, algn="l", spc=False)
            s.line(lx - 1905000 - 431800, ly + 114300, lx - 1905000 - 76200, ly + 114300)
            lx -= 1905000 + 431800 + 304800
    note = [i for i in legend if i not in ("line", "hiring")]
    if note:
        s.text(L, ly, lx - L, 228600, [(note[0], 1100, MUTED, False)], ins=(0,) * 4, algn="l", spc=False)

    over = max(bottoms) - ROW_GAP
    if over > CONTENT_BOTTOM:
        raise ValueError(f'"{spec["title"]}" não cabe no slide ({over / EMU:.2f}" > {CONTENT_BOTTOM / EMU:.2f}")')
    return s.shapes


# ---------------------------------------------------------------- montagem do pacote
def build(slides, out_path, doc_title, template=TEMPLATE):
    work = out_path + ".work"
    shutil.rmtree(work, ignore_errors=True)
    zipfile.ZipFile(template).extractall(work)
    sl = os.path.join(work, "ppt", "slides")
    tpl = open(os.path.join(sl, "slide1.xml"), encoding="utf-8").read()
    start, end = tpl.index("<p:sp>"), tpl.rindex("</p:sp>") + len("</p:sp>")
    header = re.findall(r"<p:sp>.*?</p:sp>", tpl, re.S)[:3]

    ct_path = os.path.join(work, "[Content_Types].xml")
    ct = open(ct_path, encoding="utf-8").read()
    rels_path = os.path.join(work, "ppt", "_rels", "presentation.xml.rels")
    rels = open(rels_path, encoding="utf-8").read()
    pres_path = os.path.join(work, "ppt", "presentation.xml")
    pres = open(pres_path, encoding="utf-8").read()
    slide_rel = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/'
                 'package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                 'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')

    ids = ['<p:sldId id="256" r:id="rId2"/>']
    for n, spec in enumerate(slides, start=1):
        hdr = [header[0],
               re.sub(r"<a:t>[^<]*</a:t>", f"<a:t>{escape(spec['subtitle'])}</a:t>", header[1]),
               re.sub(r"<a:t>[^<]*</a:t>", f"<a:t>{escape(spec['title'])}</a:t>", header[2])]
        xml = tpl[:start] + "".join(_render(spec, hdr)) + tpl[end:]
        open(os.path.join(sl, f"slide{n}.xml"), "w", encoding="utf-8").write(xml)
        if n > 1:
            rid = f"rId{100 + n}"
            open(os.path.join(sl, "_rels", f"slide{n}.xml.rels"), "w", encoding="utf-8").write(slide_rel)
            ct = ct.replace("</Types>", f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="application/'
                                        f'vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
            rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                                f'officeDocument/2006/relationships/slide" Target="slides/slide{n}.xml"/></Relationships>')
            ids.append(f'<p:sldId id="{255 + n}" r:id="{rid}"/>')
    pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", "<p:sldIdLst>" + "".join(ids) + "</p:sldIdLst>", pres)
    open(ct_path, "w", encoding="utf-8").write(ct)
    open(rels_path, "w", encoding="utf-8").write(rels)
    open(pres_path, "w", encoding="utf-8").write(pres)

    core = os.path.join(work, "docProps", "core.xml")
    c = open(core, encoding="utf-8").read()
    c = re.sub(r"<dc:title>.*?</dc:title>", f"<dc:title>{escape(doc_title)}</dc:title>", c)
    open(core, "w", encoding="utf-8").write(c)

    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(ct_path, "[Content_Types].xml")
        for root, _, files in os.walk(work):
            for f in sorted(files):
                full = os.path.join(root, f)
                arc = os.path.relpath(full, work)
                if arc != "[Content_Types].xml":
                    z.write(full, arc)
    shutil.rmtree(work)
    return out_path
