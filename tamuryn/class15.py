# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generowanie etykiet .lbx (Arial Bold 9pt lub 5pt dla wąskich ramek)
"""
import os, time, sys
import shutil, textwrap, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from ROD_11_._24mm_Logic import (
    identyfikacjaROD_czy_RESO,
    daneTasma24mmLewyGornyRog,
    wartosciZabezpieczen,
    opisyObwodow24mm,
    h, l,
    daneL123STVLD,
    rodzajZabezpieczeniaSTV_VLD_S30_,
    iloscObwodow,
    oblicz_szerokosc
)

TEMP_DIR = Path("temp_brother_unpack")
LABEL_XML = TEMP_DIR / "label.xml"
PROP_XML = TEMP_DIR / "prop.xml"
OUTPUT_LBX = Path("tasma24mm.lbx")

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

# TEMPLATE_STD = Path("tamuryn/bazolec.lbx")
# TEMPLATE_ROD = Path("tamuryn/bazolecROD.lbx")
# TEMPLATE_SL = Path("tamuryn/bazolecSL.lbx")
# TEMPLATE_SH = Path("tamuryn/bazolecSH.lbx")
# TEMPLATE_SL_SH = Path("tamuryn/bazolecSH_SL.lbx")

NS_MAIN = "http://schemas.brother.info/ptouch/2007/lbx/main"
NS_DRAW = "http://schemas.brother.info/ptouch/2007/lbx/draw"
NS_TEXT = "http://schemas.brother.info/ptouch/2007/lbx/text"
NS = {"ns0": NS_MAIN, "ns2": NS_DRAW, "ns3": NS_TEXT}

pt = lambda v: f"{v:.1f}pt"
DEF_W = 96.5
H_FRAME = 64
H_STRIP = 18
GAP = 7.0
Y_TOP = 2
START_ID = 300

aa = daneTasma24mmLewyGornyRog()
bb = wartosciZabezpieczen()
cc = opisyObwodow24mm()
records = [[aa[i], f"{bb[i]}A", cc[i]] for i in range(len(cc))]

ramki_cfg = []
dane_lista = daneL123STVLD()
rodzajBezpieki = rodzajZabezpieczeniaSTV_VLD_S30_()
for c in range(iloscObwodow):
    zabezpieczenie = rodzajBezpieki[c]
    prąd = dane_lista[c]
    szer_mm = oblicz_szerokosc(prąd, zabezpieczenie)
    if szer_mm == 89:
        ramki_cfg.append({"width": 225.0})
    elif szer_mm == 60:
        ramki_cfg.append({"width": 140.0})
    elif szer_mm == 27:
        ramki_cfg.append({"width": 72.0})
    elif szer_mm == 58:
        ramki_cfg.append({"width": 142.0})
    elif szer_mm == 37:
        ramki_cfg.append({"width": 93.0})
    elif szer_mm == 17:
        ramki_cfg.append({"width": 45.0})
ramki_cfg += [{}] * 3



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


def wrap_text(text, box_w_pt):
    est = 1.25
    max_char = max(8, int((box_w_pt - 2) / est))
    return "\n".join(textwrap.wrap(text, max_char)) or text


def make_text(x, y, w, h, content, name, obj_id, font_size):
    t = ET.Element(f"{{{NS_TEXT}}}text")
    st = ET.SubElement(t, f"{{{NS_MAIN}}}objectStyle", {
        "x": pt(x), "y": pt(y), "width": pt(w), "height": pt(h),
        "backColor": "#FFFFFF", "backPrintColorNumber": "0",
        "ropMode": "COPYPEN", "angle": "0", "anchor": "TOPLEFT", "flip": "NONE"
    })
    ET.SubElement(st, f"{{{NS_MAIN}}}pen", {
        "style": "NULL", "widthX": "0.5pt", "widthY": "0.5pt",
        "color": "#000000", "printColorNumber": "1"
    })
    ET.SubElement(st, f"{{{NS_MAIN}}}brush", {
        "style": "NULL", "color": "#000000", "printColorNumber": "1", "id": "0"
    })
    ET.SubElement(st, f"{{{NS_MAIN}}}expanded", {
        "objectName": name, "ID": str(obj_id), "lock": "0",
        "templateMergeTarget": "LABELLIST", "templateMergeType": "NONE",
        "templateMergeID": "0", "linkStatus": "NONE", "linkID": "0"
    })

    def add_font(node):
        ET.SubElement(node, f"{{{NS_TEXT}}}logFont", {
            "name": "Arial", "width": "0", "italic": "false",
            "weight": "700", "charSet": "238", "pitchAndFamily": "34"
        })
        ET.SubElement(node, f"{{{NS_TEXT}}}fontExt", {
            "effect": "BOLD", "underline": "0", "strikeout": "0",
            "size": font_size, "orgSize": font_size,
            "textColor": "#000000", "textPrintColorNumber": "1"
        })

    wrapped = wrap_text(content, w)
    finfo = ET.SubElement(t, f"{{{NS_TEXT}}}ptFontInfo")
    add_font(finfo)
    ET.SubElement(t, f"{{{NS_TEXT}}}textControl", {
        "control": "FREE", "clipFrame": "false", "aspectNormal": "false",
        "shrink": "false", "autoLF": "true", "avoidImage": "false"
    })
    ET.SubElement(t, f"{{{NS_TEXT}}}textAlign", {
        "horizontalAlignment": "CENTER", "verticalAlignment": "CENTER",
        "inLineAlignment": "BASELINE"
    })
    ET.SubElement(t, f"{{{NS_TEXT}}}textStyle", {
        "vertical": "false", "nullBlock": "false", "charSpace": "0",
        "lineSpace": "2", "orgPoint": font_size, "combinedChars": "false"
    })
    ET.SubElement(t, f"{{{NS_MAIN}}}data").text = wrapped
    s_item = ET.SubElement(t, f"{{{NS_TEXT}}}stringItem", {"charLen": str(len(content))})
    s_finfo = ET.SubElement(s_item, f"{{{NS_TEXT}}}ptFontInfo")
    add_font(s_finfo)
    return t


def obj_style(el): return el.find("ns0:objectStyle", NS)


def x_of(el): return float(obj_style(el).attrib["x"][:-2])


def odpalarka24bis():

    mode = identyfikacjaROD_czy_RESO()
    print(mode, l, h)
    if mode in {"ROD-11", "ROD-12", "ROD-14"}:
        template = TEMPLATE_ROD
    elif mode == "RESO-3F10" or mode == "RESO-3F" or mode == "RESO-3F10T":
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

    if TEMP_DIR.exists(): shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir()

    with zipfile.ZipFile(template) as zf:
        zf.extractall(TEMP_DIR)

    tree = ET.parse(LABEL_XML)
    root = tree.getroot()
    objs = root.find(".//ns0:objects", NS)
    frames = objs.findall("ns2:frame", NS)

    tail = []
    if mode in {"ROD-11", "ROD-12", "ROD-14"}:
        ramka9 = next(fr for fr in frames
                      if fr.find("ns0:objectStyle/ns0:expanded", NS).attrib["objectName"] == "Ramka1")
        right9 = x_of(ramka9) + float(obj_style(ramka9).attrib["width"][:-2])
        for el in list(objs):
            if x_of(el) >= right9:
                tail.append(objs.remove(el) or el)
        insert_x = right9 + GAP
    else:
        max_r = max(x_of(fr) + float(obj_style(fr).attrib["width"][:-2]) for fr in frames)
        insert_x = max_r + GAP

    cur_id = START_ID
    for i, rec in enumerate(records):
        w_frame = (ramki_cfg[i] if i < len(ramki_cfg) else {}).get("width", DEF_W)
        font_sz = "5pt" if w_frame in (65.0, 45.0) else "9pt"
        fr = ET.SubElement(objs, f"{{{NS_DRAW}}}frame")
        st = ET.SubElement(fr, f"{{{NS_MAIN}}}objectStyle", {
            "x": pt(insert_x), "y": pt(Y_TOP), "width": pt(w_frame), "height": pt(H_FRAME),
            "backColor": "#FFFFFF", "backPrintColorNumber": "0",
            "ropMode": "COPYPEN", "angle": "0", "anchor": "TOPLEFT", "flip": "NONE"
        })
        ET.SubElement(st, f"{{{NS_MAIN}}}pen", {
            "style": "NULL", "widthX": "0.5pt", "widthY": "0.5pt",
            "color": "#000000", "printColorNumber": "1"
        })
        ET.SubElement(st, f"{{{NS_MAIN}}}brush", {
            "style": "NULL", "color": "#000000", "printColorNumber": "1", "id": "0"
        })
        ET.SubElement(st, f"{{{NS_MAIN}}}expanded", {
            "objectName": f"Ramka{i + 10}", "ID": str(cur_id), "lock": "2"
        })
        ET.SubElement(fr, f"{{{NS_DRAW}}}frameStyle", {
            "category": "SIMPLE", "style": "0", "stretchCenter": "true"
        })
        cur_id += 1

        fields = [
            (insert_x + 2, Y_TOP, 28, H_STRIP, rec[0]),
            (insert_x + w_frame - 30, Y_TOP, 28, H_STRIP, rec[1]),
            (insert_x + 13, Y_TOP + 25, w_frame - 28, 28, rec[2])
        ]
        for j, (tx, ty, tw, th, txt) in enumerate(fields):
            objs.append(make_text(tx, ty, tw, th, txt, f"T{i}_{j}", cur_id, font_sz))
            cur_id += 1
        insert_x += w_frame + GAP

    shift = sum((ramki_cfg[i] if i < len(ramki_cfg) else {}).get("width", DEF_W) + GAP for i in range(len(records)))
    for el in tail:
        obj_style(el).attrib["x"] = pt(x_of(el) + shift)
        objs.append(el)

    tree.write(LABEL_XML, encoding="utf-8", xml_declaration=True)
    # with zipfile.ZipFile(OUTPUT_LBX, "w", zipfile.ZIP_DEFLATED) as zf:
    with zipfile.ZipFile(os.path.join(get_excel_folder_from_temp(), "tasma24mm.lbx"), "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(LABEL_XML, arcname="label.xml")
        if PROP_XML.exists():
            zf.write(PROP_XML, arcname="prop.xml")
