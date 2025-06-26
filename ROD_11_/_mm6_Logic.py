import math
from utils.odczyt_tabela_material_NewSheet import identyfikacja_czy_ARI, identyfikacja_zasilania_1_3_faz, \
    identyfikacja_czy_extra_Modem, daneTasma24mmLewyGornyRog, opisy6mmNaDali, h, l, n, RCDoznaczenie, \
    identyfikacjaROD_czy_RESO, identyfikacja_extra_blok_rozdzielczy_A, identyfikacja_extra_blok_rozdzielczy_B, \
    identyfikacja_czy_liczy_L_czy_przekladnia, identyfikacja_czy_liczy_H_czy_przekladnia, jaka_szafa_wariacie



napisy6mmZlaczkiOdp = []


def opisy_stale():


    # napisy6mmZlaczkiOdp.append("GN ")
    # napisy6mmZlaczkiOdp.append(" KZ")
    # napisy6mmZlaczkiOdp.append(" F5")
    # napisy6mmZlaczkiOdp.append(" F6")
    # napisy6mmZlaczkiOdp.append(" SL")
    # # napisy6mmZlaczkiOdp.append("A51")
    # napisy6mmZlaczkiOdp.append(" A4")

    if identyfikacja_czy_ARI() == "P4 (ARI)":
        # napisy6mmZlaczkiOdp.append("G1")
        napisy6mmZlaczkiOdp.append("A52")
        napisy6mmZlaczkiOdp.append("A52")
        napisy6mmZlaczkiOdp.append("A52")
        napisy6mmZlaczkiOdp.append("A52")
        napisy6mmZlaczkiOdp.append(" COM1")
    elif identyfikacja_czy_ARI() == "GSM":
        napisy6mmZlaczkiOdp.append(" A3")

    if identyfikacja_czy_extra_Modem() == "ODW-730-F1" or identyfikacja_czy_extra_Modem() == "ODW-730-F2" or identyfikacja_czy_extra_Modem() == "ODW-720-F1":
        napisy6mmZlaczkiOdp.append("A7")
    elif identyfikacja_czy_extra_Modem() == "DDW-120":
        napisy6mmZlaczkiOdp.append("A7")
    elif identyfikacja_czy_extra_Modem() == "DDW-142" or identyfikacja_czy_extra_Modem() == "DDW-142 485":
        napisy6mmZlaczkiOdp.append("A7")
    elif identyfikacja_czy_extra_Modem() == "MATRIX 500":
        napisy6mmZlaczkiOdp.append(" A3")

def modulu_A11():
    if l <= 4:
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
    if l >= 4 and l <= 8:
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")

        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")
    if l >= 8 and l <= 12:
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")

        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")

        napisy6mmZlaczkiOdp.append("A13")
        napisy6mmZlaczkiOdp.append("A13")
        napisy6mmZlaczkiOdp.append("A13")
        napisy6mmZlaczkiOdp.append("A13")
    if l >= 12 and l <= 16:
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")
        napisy6mmZlaczkiOdp.append("A11")

        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")
        napisy6mmZlaczkiOdp.append("A12")

        napisy6mmZlaczkiOdp.append("A13")
        napisy6mmZlaczkiOdp.append("A13")
        napisy6mmZlaczkiOdp.append("A13")
        napisy6mmZlaczkiOdp.append("A13")

        napisy6mmZlaczkiOdp.append("A14")
        napisy6mmZlaczkiOdp.append("A14")
        napisy6mmZlaczkiOdp.append("A14")
        napisy6mmZlaczkiOdp.append("A14")

def modulu_A31():
    napisy6mmZlaczkiOdp.append("A30")
    napisy6mmZlaczkiOdp.append("A30")
    napisy6mmZlaczkiOdp.append("A30")
    napisy6mmZlaczkiOdp.append("A30")
    if h <= 4:
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
    if h >= 4 and h <= 8:
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")

        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")

        napisy6mmZlaczkiOdp.append("A41")
        napisy6mmZlaczkiOdp.append("A41")
        napisy6mmZlaczkiOdp.append("A41")
        napisy6mmZlaczkiOdp.append("A41")
    if h >= 9 and h <= 18:
        napisy6mmZlaczkiOdp.append("A41")
        napisy6mmZlaczkiOdp.append("A41")
        napisy6mmZlaczkiOdp.append("A41")
        napisy6mmZlaczkiOdp.append("A41")

        napisy6mmZlaczkiOdp.append("A42")
        napisy6mmZlaczkiOdp.append("A42")
        napisy6mmZlaczkiOdp.append("A42")
        napisy6mmZlaczkiOdp.append("A42")
    if h >= 9 and h <= 12:
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")

        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")

        napisy6mmZlaczkiOdp.append("A33")
        napisy6mmZlaczkiOdp.append("A33")
        napisy6mmZlaczkiOdp.append("A33")
        napisy6mmZlaczkiOdp.append("A33")
    if h >= 13 and h <= 16:
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")
        napisy6mmZlaczkiOdp.append("A31")

        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")
        napisy6mmZlaczkiOdp.append("A32")

        napisy6mmZlaczkiOdp.append("A33")
        napisy6mmZlaczkiOdp.append("A33")
        napisy6mmZlaczkiOdp.append("A33")
        napisy6mmZlaczkiOdp.append("A33")

        napisy6mmZlaczkiOdp.append("A34")
        napisy6mmZlaczkiOdp.append("A34")
        napisy6mmZlaczkiOdp.append("A34")
        napisy6mmZlaczkiOdp.append("A34")


if identyfikacjaROD_czy_RESO() == "ROD-11":  # ROD-11

    opisy_stale()

    # napisy6mmZlaczkiOdp.append("BR")
    # napisy6mmZlaczkiOdp.append(" COM1")


elif identyfikacjaROD_czy_RESO() == "ROD-14" or identyfikacjaROD_czy_RESO() == "ROD-12":  # ROD_14

    opisy_stale()
    napisy6mmZlaczkiOdp.append("COM2")
    napisy6mmZlaczkiOdp.append(" BR")

    # moduly A ...
    modulu_A11()

# RESO-3F
elif identyfikacjaROD_czy_RESO() == "RESO-3F10":

    napisy6mmZlaczkiOdp.append(" A1")
    napisy6mmZlaczkiOdp.append(" A2")
    napisy6mmZlaczkiOdp.append(" X1")
    napisy6mmZlaczkiOdp.append(" X3")
    napisy6mmZlaczkiOdp.append(" X4")
    napisy6mmZlaczkiOdp.append(" X5")
    napisy6mmZlaczkiOdp.append(" X6")
    napisy6mmZlaczkiOdp.append(" G2")
    napisy6mmZlaczkiOdp.append(" SH")
    opisy_stale()
    if l == 0:



        napisy6mmZlaczkiOdp.remove(" COM1")
    if h != 0:
        napisy6mmZlaczkiOdp.append("A51")
        napisy6mmZlaczkiOdp.append("A51")
        napisy6mmZlaczkiOdp.append("A51")
        napisy6mmZlaczkiOdp.append("A51")

    if h > 0:
        modulu_A31()

    # if identyfikacja_czy_liczy_H_czy_przekladnia()[0] != 'brak' and identyfikacja_czy_liczy_H_czy_przekladnia()[0] != 0:
    #     napisy6mmZlaczkiOdp.append("P1")
    # if (identyfikacja_czy_liczy_H_czy_przekladnia()[1] != 'brak' and identyfikacja_czy_liczy_H_czy_przekladnia()[
    #     1] != 0 and identyfikacja_czy_liczy_L_czy_przekladnia()[1] != ".") or \
    #         identyfikacja_czy_liczy_H_czy_przekladnia()[0] == "półpośredni":
    #     napisy6mmZlaczkiOdp.append("T1")
    #     napisy6mmZlaczkiOdp.append("T2")
    #     napisy6mmZlaczkiOdp.append("T3")
    #
    # if identyfikacja_czy_liczy_L_czy_przekladnia()[0] != 'brak' and identyfikacja_czy_liczy_L_czy_przekladnia()[
    #     0] != 0 and h > 0:
    #     napisy6mmZlaczkiOdp.append("P2")
    # if identyfikacja_czy_liczy_L_czy_przekladnia()[1] != 'brak' and identyfikacja_czy_liczy_L_czy_przekladnia()[
    #     1] != 0 and identyfikacja_czy_liczy_L_czy_przekladnia()[1] != "." or \
    #         identyfikacja_czy_liczy_L_czy_przekladnia()[0] == "półpośredni":
    #     napisy6mmZlaczkiOdp.append("T1")
    #     napisy6mmZlaczkiOdp.append("T2")
    #     napisy6mmZlaczkiOdp.append("T3")
    #
    # if identyfikacja_czy_liczy_L_czy_przekladnia()[0] != 'brak' and identyfikacja_czy_liczy_L_czy_przekladnia()[
    #     0] != 0 and h == 0:
    #     napisy6mmZlaczkiOdp.append("P1")

def generator6mm_normal(napisy6mmZlaczkiOdp):
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

    base_dir = os.path.dirname(os.path.abspath(__file__))  # <- Jesteś już w ROD_11_
    original_file = os.path.join(base_dir, "Szbx6mm.lbx")  # <- bez podwójnego ROD_11_

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = f"lbx_unpacked_{timestamp}"
    temp_zip = f"temp_lbx_{timestamp}.zip"
    temp_lbx = f"temp_lbx_{timestamp}.lbx"
    modified_file = os.path.join(get_excel_folder_from_temp(), "tasma_6mm.lbx")

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
            print("⚠️ Nie znaleziono <style:paper> – nie można ustawić autoLength.")

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
            print("🗑️ Usunięto puste lub niewykorzystane pole tekstowe.")

        if modified_count == 0:
            print("❌ Nie znaleziono żadnych <pt:data> do zmiany.")
        else:
            for text_elem in objects_elem.findall("text:text", namespaces=ns):
                for pt_font_info in text_elem.findall(".//text:ptFontInfo", namespaces=ns):
                    font_ext = pt_font_info.find("text:fontExt", namespaces=ns)
                    if font_ext is not None:
                        old_size = font_ext.get("size")
                        font_ext.set("size", "8pt")
                        print(f"🔠 Zmieniono rozmiar czcionki z {old_size} na 8pt")

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
