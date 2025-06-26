import math
from utils.odczyt_tabela_material_NewSheet import identyfikacja_czy_ARI, identyfikacja_zasilania_1_3_faz, \
    identyfikacja_czy_extra_Modem, daneTasma24mmLewyGornyRog, opisy6mmNaDali, h, l, n, RCDoznaczenie, \
    identyfikacjaROD_czy_RESO, identyfikacja_extra_blok_rozdzielczy_A, identyfikacja_extra_blok_rozdzielczy_B, \
    identyfikacja_czy_liczy_L_czy_przekladnia, identyfikacja_czy_liczy_H_czy_przekladnia, jaka_szafa_wariacie
import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
napisy6mmZlaczkiOdp_long = []


def opisy_stale():
    napisy6mmZlaczkiOdp_long.append("BT1")
    napisy6mmZlaczkiOdp_long.append("GN ")
    napisy6mmZlaczkiOdp_long.append(" KZ")
    napisy6mmZlaczkiOdp_long.append(" F5")
    napisy6mmZlaczkiOdp_long.append(" F6")
    napisy6mmZlaczkiOdp_long.append(" SL")
    # napisy6mmZlaczkiOdp.append("A51")
    napisy6mmZlaczkiOdp_long.append(" A4")
    # RCD
    for element in RCDoznaczenie():
        if not (isinstance(element, float) and math.isnan(element)):
            napisy6mmZlaczkiOdp_long.append(element)

    # def .....

    for n in daneTasma24mmLewyGornyRog():
        napisy6mmZlaczkiOdp_long.append(n)
    for m in opisy6mmNaDali():
        napisy6mmZlaczkiOdp_long.append(m)
    if identyfikacja_zasilania_1_3_faz() == "L1 | L2 | L3":
        napisy6mmZlaczkiOdp_long.append("F1÷3")
    else:
        napisy6mmZlaczkiOdp_long.append("F1")


if identyfikacjaROD_czy_RESO() == "ROD-11":  # ROD-11
    for p in range(l):
        napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KL1")
    opisy_stale()

elif identyfikacjaROD_czy_RESO() == "ROD-14" or identyfikacjaROD_czy_RESO() == "ROD-12":  # ROD_14

    opisy_stale()

    # stycznik pomocniczy
    if l <= 8:
        napisy6mmZlaczkiOdp_long.append("KL1")
    elif l >= 9 and l <= 16:
        napisy6mmZlaczkiOdp_long.append("KL2")

    for p in range(l):
        napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KL1")
        napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KLU")


# RESO-3F
elif identyfikacjaROD_czy_RESO() == "RESO-3F10":

    opisy_stale()
    if l > 0:  # jeśli z oświetleniem

        for p in range(l):
            napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KL1")
            napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KLU")
        if l <= 8:
            napisy6mmZlaczkiOdp_long.append("KL1")
        elif l >= 9 and l <= 16:
            napisy6mmZlaczkiOdp_long.append("KL2")
    if h > 0:
        for p in range(l):
            napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KH1")
            napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KHU")
        if l <= 8:
            napisy6mmZlaczkiOdp_long.append("KH1")
        elif l >= 9 and l <= 16:
            napisy6mmZlaczkiOdp_long.append("KH2")

    for p in range(h):
        napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KH1")
        napisy6mmZlaczkiOdp_long.append(str(p + 1) + "KHU")
    if jaka_szafa_wariacie() == "Typ 1" and h != 0 and l == 0 and n == 0:
        napisy6mmZlaczkiOdp_long.append("BR-H")
    if jaka_szafa_wariacie() == "Typ 1" and h != 0 and l != 0 and n == 0:
        napisy6mmZlaczkiOdp_long.append("BR-H")
        napisy6mmZlaczkiOdp_long.append("BR-L")
    if jaka_szafa_wariacie() == "Typ 1" and h != 0 and l == 0 and n != 0:
        napisy6mmZlaczkiOdp_long.append("BR-H")
        napisy6mmZlaczkiOdp_long.append("BR-N")
    if jaka_szafa_wariacie() == "Typ 1" and h == 0 and l != 0 and n != 0:
        napisy6mmZlaczkiOdp_long.append("BR-L")
        napisy6mmZlaczkiOdp_long.append("BR-N")
    if jaka_szafa_wariacie() == "Typ 2":
        if h != 0 and l == 0 and n == 0:
            napisy6mmZlaczkiOdp_long.append("BR-H")
        if h != 0 and l != 0 and n == 0:
            napisy6mmZlaczkiOdp_long.append("BR-H")
            napisy6mmZlaczkiOdp_long.append("BR-L")
        if h != 0 and l == 0 and n != 0:
            napisy6mmZlaczkiOdp_long.append("BR-H")
            napisy6mmZlaczkiOdp_long.append("BR-N")
        if h == 0 and l != 0 and n != 0:
            napisy6mmZlaczkiOdp_long.append("BR-L")
            napisy6mmZlaczkiOdp_long.append("BR-N")

    if jaka_szafa_wariacie() == "Typ 2+0" or jaka_szafa_wariacie() == "Typ 2+1" or jaka_szafa_wariacie() == "Typ 2+2":
        napisy6mmZlaczkiOdp_long.append("BT2")

        if identyfikacja_extra_blok_rozdzielczy_A()[0] != 'brak' and identyfikacja_extra_blok_rozdzielczy_A()[0] != 0:
            napisy6mmZlaczkiOdp_long.append("BR-H1")
        if identyfikacja_extra_blok_rozdzielczy_A()[1] != 'brak' and identyfikacja_extra_blok_rozdzielczy_A()[1] != 0:
            napisy6mmZlaczkiOdp_long.append("BR-L1")
        if identyfikacja_extra_blok_rozdzielczy_A()[2] != 'brak' and identyfikacja_extra_blok_rozdzielczy_A()[2] != 0:
            napisy6mmZlaczkiOdp_long.append("BR-N1")

        if identyfikacja_extra_blok_rozdzielczy_B()[0] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[0] != 0:
            napisy6mmZlaczkiOdp_long.append("BR-H2")
        if identyfikacja_extra_blok_rozdzielczy_B()[1] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[1] != 0:
            napisy6mmZlaczkiOdp_long.append("BR-L2")
        if identyfikacja_extra_blok_rozdzielczy_B()[2] != 'brak' and identyfikacja_extra_blok_rozdzielczy_B()[2] != 0:
            napisy6mmZlaczkiOdp_long.append("BR-N2")


        if identyfikacja_czy_liczy_H_czy_przekladnia()[0] != 'brak' and identyfikacja_czy_liczy_H_czy_przekladnia()[
            0] != 0:
            napisy6mmZlaczkiOdp_long.append("P1")
        if (identyfikacja_czy_liczy_H_czy_przekladnia()[1] != 'brak' and identyfikacja_czy_liczy_H_czy_przekladnia()[
            1] != 0 and identyfikacja_czy_liczy_L_czy_przekladnia()[1] != ".") or \
                identyfikacja_czy_liczy_H_czy_przekladnia()[0] == "półpośredni":
            napisy6mmZlaczkiOdp_long.append("T1")
            napisy6mmZlaczkiOdp_long.append("T2")
            napisy6mmZlaczkiOdp_long.append("T3")

        if identyfikacja_czy_liczy_L_czy_przekladnia()[0] != 'brak' and identyfikacja_czy_liczy_L_czy_przekladnia()[
            0] != 0 and h > 0:
            napisy6mmZlaczkiOdp_long.append("P2")
        if identyfikacja_czy_liczy_L_czy_przekladnia()[1] != 'brak' and identyfikacja_czy_liczy_L_czy_przekladnia()[
            1] != 0 and identyfikacja_czy_liczy_L_czy_przekladnia()[1] != "." or \
                identyfikacja_czy_liczy_L_czy_przekladnia()[0] == "półpośredni":
            napisy6mmZlaczkiOdp_long.append("T1")
            napisy6mmZlaczkiOdp_long.append("T2")
            napisy6mmZlaczkiOdp_long.append("T3")

        if identyfikacja_czy_liczy_L_czy_przekladnia()[0] != 'brak' and identyfikacja_czy_liczy_L_czy_przekladnia()[
            0] != 0 and h == 0:
            napisy6mmZlaczkiOdp_long.append("P1")

def generator6mm_long(napisy6mmZlaczkiOdp):
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
    original_file = os.path.join(base_dir, "Szbx6mm_long.lbx")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = f"lbx_unpacked_{timestamp}"
    temp_zip = f"temp_lbx_{timestamp}.zip"
    temp_lbx = f"temp_lbx_{timestamp}.lbx"
    modified_file = os.path.join(get_excel_folder_from_temp(), "tasma_6mm_dluga.lbx")

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

        replacement_map = stworz_replacement_map(napisy6mmZlaczkiOdp)

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
                        font_ext.set("size", "10pt")
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
