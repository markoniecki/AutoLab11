#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generowanie etykiet .lbx  (Arial 9 pt Bold – wyjątki 5 pt)
---------------------------------------------------------
• ROD-11/12/14                    → bazolecROD.lbx  (pomiędzy Ramka1 a Ramka9)
• RESO-3F10 +  l≠0 , h==0        → bazolecSL.lbx
• RESO-3F10 +  l==0, h≠0        → bazolecSH.lbx
• RESO-3F10 +  l≠0 , h≠0        → bazolecSH_SL.lbx
• Pozostałe przypadki            → bazolec.lbx     (dokładamy na końcu)
"""

import shutil, textwrap, zipfile, sys, os, time
from pathlib import Path
import xml.etree.ElementTree as ET

# ---- logika projektu ----
from ROD_11_._24mm_Logic import (
    identyfikacjaROD_czy_RESO,
    daneTasma24mmLewyGornyRog,
    wartosciZabezpieczen,
    opisyObwodow24mm,
    h,  # wysokość (mm)
    l,  # długość  (mm)
    daneL123STVLD,
    rodzajZabezpieczeniaSTV_VLD_S30_,
    iloscObwodow,
    oblicz_szerokosc,
)

# ---------- ścieżki ----------
TEMP_DIR      = Path("temp_brother_unpack")
LABEL_XML     = TEMP_DIR / "label.xml"
PROP_XML      = TEMP_DIR / "prop.xml"
OUTPUT_LBX    = Path("wyginam_poprawka_z_tekstami.lbx")

def resource_path(relative_path):
    """Zwraca ścieżkę do pliku, zgodną z trybem exe i trybem deweloperskim."""
    try:
        base_path = sys._MEIPASS  # folder tymczasowy PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

TEMPLATE_STD    = Path(resource_path("tamuryn/bazolec.lbx"))
TEMPLATE_ROD    = Path(resource_path("tamuryn/bazolecROD.lbx"))
TEMPLATE_SL     = Path(resource_path("tamuryn/bazolecSL.lbx"))
TEMPLATE_SH     = Path(resource_path("tamuryn/bazolecSH.lbx"))
TEMPLATE_SL_SH  = Path(resource_path("tamuryn/bazolecSH_SL.lbx"))

# TEMPLATE_STD   = Path("tamuryn/bazolec.lbx")
# TEMPLATE_ROD   = Path("tamuryn/bazolecROD.lbx")
# TEMPLATE_SL    = Path("tamuryn/bazolecSL.lbx")       # RESO: l≠0, h==0
# TEMPLATE_SH    = Path("tamuryn/bazolecSH.lbx")       # RESO: l==0, h≠0
# TEMPLATE_SL_SH = Path("bazolecSH_SL.lbx")    # RESO: l≠0, h≠0

# ---------- XML NS ----------
NS_MAIN = "http://schemas.brother.info/ptouch/2007/lbx/main"
NS_DRAW = "http://schemas.brother.info/ptouch/2007/lbx/draw"
NS_TEXT = "http://schemas.brother.info/ptouch/2007/lbx/text"
NS = {"ns0": NS_MAIN, "ns2": NS_DRAW, "ns3": NS_TEXT}

# ---------- geometria ----------
pt = lambda v: f"{v:.1f}pt"
DEF_W    = 96.5
H_FRAME  = 64
H_STRIP  = 18
GAP      = 3.0
Y_TOP    = 2
START_ID = 300

# ---------- dane z modułu ----------
aa = daneTasma24mmLewyGornyRog()
bb = wartosciZabezpieczen()
cc = opisyObwodow24mm()
records = [[aa[i], f"{bb[i]}A", cc[i]] for i in range(len(cc))]

# wyliczenie szerokości ramek:
ramki_cfg = []
dane_lista     = daneL123STVLD()
rodzajBezpieki = rodzajZabezpieczeniaSTV_VLD_S30_()
for c in range(iloscObwodow):
    szer_mm = oblicz_szerokosc(dane_lista[c], rodzajBezpieki[c])
    ramki_cfg.append({
        89: {"width": 213.0},  # STV 3-p
        60: {"width": 140.0},  # STV 2-p
        27: {"width":  65.0},  # STV 1-p  (wąska → 5 pt)
        58: {"width": 142.0},  # VLD 3-p
        37: {"width":  93.0},  # VLD 2-p
        17: {"width":  45.0},  # VLD 1-p  (wąska → 5 pt)
    }[szer_mm])

# dopisz „puste” – gdyby było mniej konfiguracji niż rekordów
ramki_cfg += [{}] * 3

# ---------- pomocnicze ----------
def wrap_text(text: str, box_w_pt: float) -> str:
    """Zawijanie: stałe 25 znaków, chyba że ramka wąska (45/65 pt)."""
    if box_w_pt in (45.0, 65.0):          # wąskie 1-polowe – zawężamy
        max_chars = max(6, int((box_w_pt - 2) / 0.9))
    else:
        max_chars = 25                    # szerokie – stały limit
    return "\n".join(textwrap.wrap(text, max_chars)) or text

def font_size(box_w_pt: float) -> str:
    """9 pt wszędzie, 5 pt tylko dla ramek 45 pt i 65 pt."""
    return "5pt" if box_w_pt in (45.0, 65.0) else "9pt"

def make_text(x, y, w, h, content, name, obj_id, fsize):
    # fsize = font_size(w)
    wrapped = wrap_text(content, w)

    t  = ET.Element(f"{{{NS_TEXT}}}text")
    st = ET.SubElement(t, f"{{{NS_MAIN}}}objectStyle", {
        "x": pt(x), "y": pt(y), "width": pt(w), "height": pt(h),
        "backColor": "#FFFFFF", "backPrintColorNumber": "0",
        "ropMode": "COPYPEN", "angle": "0", "anchor": "TOPLEFT", "flip": "NONE"
    })
    ET.SubElement(st, f"{{{NS_MAIN}}}pen", {
        "style":"NULL","widthX":"0.5pt","widthY":"0.5pt",
        "color":"#000000","printColorNumber":"1"})
    ET.SubElement(st, f"{{{NS_MAIN}}}brush", {
        "style":"NULL","color":"#000000","printColorNumber":"1","id":"0"})
    ET.SubElement(st, f"{{{NS_MAIN}}}expanded", {
        "objectName": name, "ID": str(obj_id), "lock":"0"})

    def add_font(node):
        ET.SubElement(node, f"{{{NS_TEXT}}}logFont", {
            "name":"Arial","width":"0","italic":"false",
            "weight":"700","charSet":"238","pitchAndFamily":"34"})
        ET.SubElement(node, f"{{{NS_TEXT}}}fontExt", {
            "effect":"BOLD","underline":"0","strikeout":"0",
            "size":fsize,"orgSize":fsize,
            "textColor":"#000000","textPrintColorNumber":"1"})

    add_font(ET.SubElement(t, f"{{{NS_TEXT}}}ptFontInfo"))

    ET.SubElement(t, f"{{{NS_TEXT}}}textControl", {
        "control":"FREE","clipFrame":"false","aspectNormal":"false",
        "shrink":"false","autoLF":"true","avoidImage":"false"})
    ET.SubElement(t, f"{{{NS_TEXT}}}textAlign", {
        "horizontalAlignment":"CENTER","verticalAlignment":"CENTER",
        "inLineAlignment":"BASELINE"})
    ET.SubElement(t, f"{{{NS_TEXT}}}textStyle", {
        "vertical":"false","nullBlock":"false","charSpace":"0",
        "lineSpace":"2","orgPoint":fsize,"combinedChars":"false"})
    ET.SubElement(t, f"{{{NS_MAIN}}}data").text = wrapped

    add_font(ET.SubElement(
        ET.SubElement(t, f"{{{NS_TEXT}}}stringItem", {"charLen": str(len(content))}),
        f"{{{NS_TEXT}}}ptFontInfo"
    ))
    return t

def obj_style(el): return el.find("ns0:objectStyle", NS)
def x_of(el):     return float(obj_style(el).attrib["x"][:-2])
def get_excel_folder_from_temp():
    temp_dir = Path(os.environ['TEMP'])
    coton_path = temp_dir / "coton.txt"

    timeout = 5
    waited = 0
    while not coton_path.exists():
        time.sleep(0.2)
        waited += 0.2
        if waited >= timeout:
            raise FileNotFoundError(f"❌ Nie znaleziono pliku: {coton_path}")

    with open(coton_path, 'r', encoding='cp1250') as f:
        folder_path = f.read().strip()

    return folder_path


def odpalarka24bis():
    # ---------- 1. wybór szablonu ----------
    mode = identyfikacjaROD_czy_RESO()
    if mode in {"ROD-11","ROD-12","ROD-14"}:
        template = TEMPLATE_ROD
    elif mode == "RESO-3F10":
        if l != 0 and h == 0:
            template = TEMPLATE_SL
        elif l == 0 and h != 0:
            template = TEMPLATE_SH
        elif l != 0 and h != 0:
            template = TEMPLATE_SL_SH
        else:
            template = TEMPLATE_STD
    else:
        template = TEMPLATE_STD

    if not template.exists():
        raise FileNotFoundError(f"Brak szablonu: {template}")

    # ---------- 2. rozpakuj ----------
    if TEMP_DIR.exists(): shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir()

    with zipfile.ZipFile(template) as zf:
        zf.extractall(TEMP_DIR)

    tree   = ET.parse(LABEL_XML)
    root   = tree.getroot()
    objs   = root.find(".//ns0:objects", NS)
    frames = objs.findall("ns2:frame", NS)

    # ---------- 3. przygotuj ogon ----------
    tail = []
    if mode in {"ROD-11","ROD-12","ROD-14"}:
        ramka9 = next(fr for fr in frames
                      if fr.find("ns0:objectStyle/ns0:expanded", NS)
                        .attrib["objectName"] == "Ramka9")
        right9 = x_of(ramka9) + float(obj_style(ramka9).attrib["width"][:-2])
        for el in list(objs):
            if x_of(el) >= right9:
                tail.append(objs.remove(el) or el)
        insert_x = right9 + GAP
    else:
        insert_x = max(x_of(fr)+float(obj_style(fr).attrib["width"][:-2]) for fr in frames) + GAP

    # ---------- 4. wstaw nowe ramki ----------
    cur_id = START_ID
    for i, rec in enumerate(records):
        w = (ramki_cfg[i] if i < len(ramki_cfg) else {}).get("width", DEF_W)

        fr = ET.SubElement(objs, f"{{{NS_DRAW}}}frame")
        st = ET.SubElement(fr, f"{{{NS_MAIN}}}objectStyle", {
            "x": pt(insert_x), "y": pt(Y_TOP), "width": pt(w), "height": pt(H_FRAME),
            "backColor":"#FFFFFF","backPrintColorNumber":"0",
            "ropMode":"COPYPEN","angle":"0","anchor":"TOPLEFT","flip":"NONE"})
        ET.SubElement(st, f"{{{NS_MAIN}}}pen", {
            "style":"NULL","widthX":"0.5pt","widthY":"0.5pt",
            "color":"#000000","printColorNumber":"1"})
        ET.SubElement(st, f"{{{NS_MAIN}}}brush",{
            "style":"NULL","color":"#000000","printColorNumber":"1","id":"0"})
        ET.SubElement(st, f"{{{NS_MAIN}}}expanded", {
            "objectName":f"Ramka{i+10}","ID":str(cur_id),"lock":"2"})
        ET.SubElement(fr, f"{{{NS_DRAW}}}frameStyle",
                      {"category":"SIMPLE","style":"0","stretchCenter":"true"})
        cur_id += 1

        # ─── POLA TEKSTOWE ───────────────────────────────────────────────────
        mm2pt = lambda mm: mm / 0.35278  # 1 mm ≈ 2.8378 pt

        if w in (45.0, 65.0):  # Ramka11 / Ramka12 – wąskie
            top_w = mm2pt(6)  # ≈ 17 pt
            top_h = mm2pt(4)  # ≈ 11.3 pt
            gap_pt = 1.0

            fields = [
                # lewy górny narożnik (czcionka 5 pt)
                (insert_x + gap_pt, Y_TOP + gap_pt,
                 top_w, top_h, rec[0], "5pt"),

                # prawy górny narożnik (czcionka 5 pt)
                (insert_x + w - top_w - gap_pt, Y_TOP + gap_pt,
                 top_w, top_h, rec[1], "5pt"),

                # opis centralny (czcionka 5 pt)
                (insert_x + gap_pt * 4,
                 Y_TOP + top_h + gap_pt * 2,
                 w - gap_pt * 8,
                 H_FRAME - (top_h + gap_pt * 4),
                 rec[2], "5pt"),
            ]
        else:  # pozostałe przypadki (czcionka 9 pt)
            fields = [
                (insert_x + 2, Y_TOP, 28, H_STRIP, rec[0], "9pt"),
                (insert_x + w - 30, Y_TOP, 28, H_STRIP, rec[1], "9pt"),
                (insert_x + 13, Y_TOP + 25, w - 28, 28, rec[2], "9pt"),
            ]

        # wstaw pola
        for j, (tx, ty, tw, th, txt, fsize) in enumerate(fields):
            objs.append(make_text(tx, ty, tw, th, txt, f"T{i}_{j}", cur_id, fsize))
            cur_id += 1

        insert_x += w + GAP

    # ---------- 5. doklej ogon ----------
    shift = sum((ramki_cfg[i] if i < len(ramki_cfg) else {}).get("width", DEF_W) + GAP
                for i in range(len(records)))

    for el in tail:
        obj_style(el).attrib["x"] = pt(x_of(el) + shift)
        objs.append(el)

    # ---------- 6. zapis ----------
    tree.write(LABEL_XML, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(os.path.join(get_excel_folder_from_temp(), "tasma24mm.lbx"), "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(LABEL_XML, arcname="label.xml")
        if PROP_XML.exists():
            zf.write(PROP_XML, arcname="prop.xml")


