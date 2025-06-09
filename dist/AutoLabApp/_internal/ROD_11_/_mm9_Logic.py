from utils.odczyt_tabela_material_NewSheet import identyfikacjaROD_czy_RESO, generujeNapisyNaZlaczkiOdplywowe, \
    identyfikacja_czy_ARI, identyfikacja_czy_extra_Modem, jaka_szafa_wariacie, identyfikacja_extra_blok_rozdzielczy_B, \
    h, \
    l, identyfikacja_extra_blok_rozdzielczy_A, identyfikacja_czy_liczy_H_czy_przekladnia

import os
import zipfile
import shutil
import xml.etree.ElementTree as ET

napisy9mmZlaczkiOdp = []

def funkcja_do_opisow_9mm():  # działa dla ROD_11_ i ROD-14
    # napisy9mmZlaczkiOdp = []
    if identyfikacja_czy_liczy_H_czy_przekladnia() == 'półpośredni':
        napisy9mmZlaczkiOdp.append("LR1")
    napisy9mmZlaczkiOdp.append("L1")
    napisy9mmZlaczkiOdp.append("L2")
    napisy9mmZlaczkiOdp.append("L3")
    napisy9mmZlaczkiOdp.append("N")
    napisy9mmZlaczkiOdp.append("0X")
    napisy9mmZlaczkiOdp.append("XBH")
    napisy9mmZlaczkiOdp.append("XSH")
    napisy9mmZlaczkiOdp.append("E1")
    for a in generujeNapisyNaZlaczkiOdplywowe():
        napisy9mmZlaczkiOdp.append(a)
    if identyfikacja_czy_ARI() == "P4 (ARI)":
        napisy9mmZlaczkiOdp.append("XT")
        napisy9mmZlaczkiOdp.append("FT1")
    if identyfikacja_czy_extra_Modem() == "ODW-730-F1" or identyfikacja_czy_extra_Modem() == "ODW-730-F2" or identyfikacja_czy_extra_Modem() == "ODW-720-F1":
        napisy9mmZlaczkiOdp.append("A7")
    elif identyfikacja_czy_extra_Modem() == "DDW-120":
        napisy9mmZlaczkiOdp.append("XTD1")
        napisy9mmZlaczkiOdp.append("FT1")
    elif identyfikacja_czy_extra_Modem() == "DDW-142" or identyfikacja_czy_extra_Modem() == "DDW-142 485":
        napisy9mmZlaczkiOdp.append("XTD1")
        napisy9mmZlaczkiOdp.append("XTD2")
        napisy9mmZlaczkiOdp.append("FT1")
        napisy9mmZlaczkiOdp.append("FT2")
    elif identyfikacja_czy_extra_Modem() == "MATRIX 500":
        napisy9mmZlaczkiOdp.append("A2")
    if jaka_szafa_wariacie() == "Typ 0" or jaka_szafa_wariacie() == "Typ 1":
        napisy9mmZlaczkiOdp.append("EG1")
        napisy9mmZlaczkiOdp.append("SK1")

    elif jaka_szafa_wariacie() == "Typ 2" or jaka_szafa_wariacie() == "Typ 3":
        napisy9mmZlaczkiOdp.append("EG1")
        napisy9mmZlaczkiOdp.append("EG2")
        napisy9mmZlaczkiOdp.append("SK1")
    elif jaka_szafa_wariacie() == "Typ 2+0" or jaka_szafa_wariacie() == "Typ 2+1":
        napisy9mmZlaczkiOdp.append("EG11")
        napisy9mmZlaczkiOdp.append("EG12")
        napisy9mmZlaczkiOdp.append("EG21")
        napisy9mmZlaczkiOdp.append("SK1")
        napisy9mmZlaczkiOdp.append("SK2")
        napisy9mmZlaczkiOdp.append("E2")
    elif jaka_szafa_wariacie() == "Typ 2+2":
        napisy9mmZlaczkiOdp.append("EG11")
        napisy9mmZlaczkiOdp.append("EG12")
        napisy9mmZlaczkiOdp.append("EG21")
        napisy9mmZlaczkiOdp.append("EG22")
        napisy9mmZlaczkiOdp.append("SK1")
        napisy9mmZlaczkiOdp.append("SK2")
        napisy9mmZlaczkiOdp.append("E2")
    if identyfikacja_extra_blok_rozdzielczy_B()[0] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[0] != 0:
        napisy9mmZlaczkiOdp.append("0XH")
    elif identyfikacja_extra_blok_rozdzielczy_B()[1] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[1] != 0:
        napisy9mmZlaczkiOdp.append("0XL")
    elif identyfikacja_extra_blok_rozdzielczy_B()[2] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[2] != 0:
        napisy9mmZlaczkiOdp.append("0XN")
    if identyfikacjaROD_czy_RESO() == "ROD-11":
        napisy9mmZlaczkiOdp.remove("XBH")
        napisy9mmZlaczkiOdp.remove("XSH")

    return napisy9mmZlaczkiOdp

    # return napisy9mmZlaczkiOdp
if identyfikacjaROD_czy_RESO() == "ROD-11":
    funkcja_do_opisow_9mm()

elif identyfikacjaROD_czy_RESO() == "ROD-14" or identyfikacjaROD_czy_RESO() == "ROD-12":
    funkcja_do_opisow_9mm()

elif identyfikacjaROD_czy_RESO() == "RESO-3F10":

    funkcja_do_opisow_9mm()

    if h > 0:
        napisy9mmZlaczkiOdp.append("XSH")
        napisy9mmZlaczkiOdp.append("XBH")
    elif l > 0:
        napisy9mmZlaczkiOdp.append("XSL")

    if ((identyfikacja_extra_blok_rozdzielczy_B()[0] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[0] != 0) or
            (identyfikacja_extra_blok_rozdzielczy_B()[1] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[
                1] != 0) or (
                    identyfikacja_extra_blok_rozdzielczy_B()[2] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[
                2] != 0)):
        napisy9mmZlaczkiOdp.append("0XZ")

def generator9mm(napisy9mmZlaczkiOdp):
    import os
    import shutil
    import zipfile
    import time
    import xml.etree.ElementTree as ET
    from pathlib import Path
    from datetime import datetime

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

    base_dir = os.path.dirname(os.path.abspath(__file__))
    original_file = os.path.join(base_dir, "Szbx9mm.lbx")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = f"lbx_unpacked_{timestamp}"
    temp_zip = f"temp_lbx_{timestamp}.zip"
    temp_lbx = f"temp_lbx_{timestamp}.lbx"
    modified_file = os.path.join(get_excel_folder_from_temp(), "Szbx9mm_modified.lbx")

    os.makedirs(work_dir, exist_ok=True)

    try:
        with zipfile.ZipFile(original_file, 'r') as zip_ref:
            zip_ref.extractall(work_dir)

        label_path = os.path.join(work_dir, "label.xml")
        tree = ET.parse(label_path)
        root = tree.getroot()

        ns = {
            "pt": "http://schemas.brother.info/ptouch/2007/lbx/main",
            "text": "http://schemas.brother.info/ptouch/2007/lbx/text",
            "style": "http://schemas.brother.info/ptouch/2007/lbx/style"
        }

        paper_elem = root.find(".//style:paper", namespaces=ns)
        if paper_elem is not None:
            paper_elem.set("autoLength", "true")
            print("✨ Ustawiono autoLength=True")
        else:
            print("⚠️ Nie znaleziono <style:paper>")

        def stworz_replacement_map(tablicex):
            return {f"Tekst{i + 1}": val.strip() for i, val in enumerate(tablicex)}

        replacement_map = stworz_replacement_map(napisy9mmZlaczkiOdp)

        objects_elem = root.find(".//pt:objects", namespaces=ns)
        modified_count = 0
        to_remove = []

        for text_elem in objects_elem.findall("text:text", namespaces=ns):
            expanded = text_elem.find(".//pt:expanded", namespaces=ns)
            obj_name = expanded.attrib.get("objectName") if expanded is not None else None
            data_elem = text_elem.find("pt:data", namespaces=ns)

            if obj_name in replacement_map:
                if data_elem is not None:
                    old_text = data_elem.text
                    data_elem.text = replacement_map[obj_name]
                    print(f"✅ Zmieniono {obj_name}: '{old_text}' → '{data_elem.text}'")
                    modified_count += 1
            else:
                if data_elem is None or data_elem.text is None or data_elem.text.strip() == "":
                    to_remove.append(text_elem)

        for elem in to_remove:
            objects_elem.remove(elem)
            print("🗑️ Usunięto puste pole tekstowe.")

        if modified_count == 0:
            print("❌ Nie znaleziono danych do modyfikacji.")
        else:
            for text_elem in objects_elem.findall("text:text", namespaces=ns):
                for pt_font_info in text_elem.findall(".//text:ptFontInfo", namespaces=ns):
                    font_ext = pt_font_info.find("text:fontExt", namespaces=ns)
                    if font_ext is not None:
                        old_size = font_ext.get("size")
                        font_ext.set("size", "18pt")
                        print(f"🔠 Czcionka: {old_size} → 8pt")

            tree.write(label_path, encoding="utf-8", xml_declaration=True)

            if os.path.exists(modified_file):
                os.remove(modified_file)

            shutil.make_archive(temp_zip.replace(".zip", ""), 'zip', work_dir)
            os.rename(temp_zip, temp_lbx)
            shutil.move(temp_lbx, modified_file)

            print(f"✅ Plik zapisano w: {modified_file}")

    finally:
        if os.path.exists(work_dir):
            shutil.rmtree(work_dir)




