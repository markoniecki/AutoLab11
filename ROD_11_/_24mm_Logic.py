from ROD_11_.ROD_11_24_Class import RectangleImage, split_text
from utils.odczyt_tabela_material_NewSheet import iloscObwodow, rodzajZabezpieczeniaSTV_VLD_S30_, \
    oblicz_szerokosc, opisyObwodow24mm, daneTasma24mmLewyGornyRog, daneL123STVLD, wartosciZabezpieczen, RCDcurrent, \
    RCDoznaczenie, identyfikacjaRCD, h, l, n, identyfikacjaROD_czy_RESO
import time
from pathlib import Path
from datetime import datetime
import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
import base64

base_path = os.path.dirname(os.path.abspath(__file__))
from pathlib import Path
import os
import shutil
import zipfile
import base64
import xml.etree.ElementTree as ET
from datetime import datetime


# def update_obraz_lbx(rectangle):
#     # Ścieżki i pliki
#     temp_dir = Path(os.environ['TEMP'])
#     timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#     bmp_file = temp_dir / f"24mmPt1_{timestamp}.bmp"
#
#     rectangle.save_image(str(bmp_file))  # Zapisz obraz BMP
#
#     original_file = os.path.join(os.getcwd(), "ROD_11_", "Obraz.lbx")
#     work_dir = "lbx_unpacked"
#     temp_zip = "temp_lbx.zip"
#     temp_lbx = "temp_lbx.lbx"
#     modified_file = os.path.join(get_excel_folder_from_temp(), "24mm.lbx")
#
#     # Rozpakuj .lbx jako ZIP
#     with zipfile.ZipFile(original_file, 'r') as zip_ref:
#         zip_ref.extractall(work_dir)
#
#     # Nadpisz plik BMP
#     target_image_path = os.path.join(work_dir, "Object0.bmp")
#     shutil.copyfile(bmp_file, target_image_path)
#     print(f"🖼️ Nadpisano plik obrazka: {target_image_path}")
#
#     # Parsuj XML
#     label_path = os.path.join(work_dir, "label.xml")
#     tree = ET.parse(label_path)
#     root = tree.getroot()
#
#     ns = {
#         "pt": "http://schemas.brother.info/ptouch/2007/lbx/main",
#         "image": "http://schemas.brother.info/ptouch/2007/lbx/image",
#         "style": "http://schemas.brother.info/ptouch/2007/lbx/style"
#     }
#
#     # Ustaw autoLength=True
#     paper_elem = root.find(".//style:paper", namespaces=ns)
#     if paper_elem is not None:
#         paper_elem.set("autoLength", "true")
#         print("✨ Ustawiono autoLength=True")
#     else:
#         print("⚠️ Nie znaleziono <style:paper>")
#
#     # Zakoduj BMP do base64
#     with open(bmp_file, "rb") as f:
#         b64_image = base64.b64encode(f.read()).decode("utf-8")
#
#     # Znajdź i podmień <pt:data> w elemencie image:image[objectName="Obraz3"]
#     found = False
#     for image_elem in root.findall(".//image:image", namespaces=ns):
#         expanded = image_elem.find(".//pt:expanded", namespaces=ns)
#         if expanded is not None and expanded.attrib.get("objectName") == "Obraz3":
#             data_elem = image_elem.find("pt:data", namespaces=ns)
#             if data_elem is None:
#                 image_style = image_elem.find("image:imageStyle", namespaces=ns)
#                 data_elem = ET.Element(f"{{{ns['pt']}}}data")
#                 data_elem.text = b64_image
#                 image_elem.insert(list(image_elem).index(image_style), data_elem)
#                 print("➕ Dodano brakujący <pt:data>")
#             else:
#                 data_elem.text = b64_image
#                 print("✏️ Nadpisano istniejący <pt:data>")
#             found = True
#             break
#
#     if not found:
#         print("❌ Nie znaleziono elementu Obraz3.")
#     else:
#         # Zapisz XML
#         tree.write(label_path, encoding="utf-8", xml_declaration=True)
#
#         # Usuń stary plik wynikowy
#         if os.path.exists(modified_file):
#             os.remove(modified_file)
#
#         # Spakuj jako ZIP
#         shutil.make_archive("temp_lbx", 'zip', work_dir)
#         os.rename(temp_zip, temp_lbx)
#
#         # Przenieś do folderu docelowego
#         shutil.move(temp_lbx, modified_file)
#
#         # Sprzątanie
#         shutil.rmtree(work_dir)
#
#         print(f"✅ Plik zapisano w: {modified_file}")

def update_obraz_lbx(rectangle):
    # --- Zapis obrazu BMP ---
    temp_dir = Path(os.environ['TEMP'])
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bmp_file = temp_dir / f"24mmPt1_{timestamp}.bmp"
    rectangle.save_image(str(bmp_file))

    # --- Praca na .lbx ---
    def get_excel_folder_from_temp():
        coton_path = temp_dir / "coton.txt"
        waited = 0
        while not coton_path.exists():
            time.sleep(0.2)
            waited += 0.2
            if waited >= 5:
                raise FileNotFoundError("❌ Brak coton.txt")
        return Path(coton_path.read_text(encoding='cp1250').strip())

    base_dir = os.path.dirname(os.path.abspath(__file__))
    original_file = os.path.join(base_dir, "Obraz.lbx")
    work_dir = f"lbx_unpacked_{timestamp}"
    temp_zip = f"temp_lbx_{timestamp}.zip"
    temp_lbx = f"temp_lbx_{timestamp}.lbx"
    modified_file = get_excel_folder_from_temp() / "24mm.lbx"

    os.makedirs(work_dir, exist_ok=True)

    try:
        with zipfile.ZipFile(original_file, 'r') as zip_ref:
            zip_ref.extractall(work_dir)

        shutil.copyfile(bmp_file, os.path.join(work_dir, "Object0.bmp"))
        print(f"🖼️ Nadpisano Object0.bmp")

        label_path = os.path.join(work_dir, "label.xml")
        tree = ET.parse(label_path)
        root = tree.getroot()

        ns = {
            "pt": "http://schemas.brother.info/ptouch/2007/lbx/main",
            "image": "http://schemas.brother.info/ptouch/2007/lbx/image",
            "style": "http://schemas.brother.info/ptouch/2007/lbx/style"
        }

        paper_elem = root.find(".//style:paper", namespaces=ns)
        if paper_elem is not None:
            paper_elem.set("autoLength", "true")
            print("✨ Ustawiono autoLength=True")

        with open(bmp_file, "rb") as f:
            b64_image = base64.b64encode(f.read()).decode("utf-8")

        found = False
        for image_elem in root.findall(".//image:image", namespaces=ns):
            expanded = image_elem.find(".//pt:expanded", namespaces=ns)
            if expanded is not None and expanded.attrib.get("objectName") == "Obraz3":
                data_elem = image_elem.find("pt:data", namespaces=ns)
                if data_elem is None:
                    image_style = image_elem.find("image:imageStyle", namespaces=ns)
                    data_elem = ET.Element(f"{{{ns['pt']}}}data")
                    data_elem.text = b64_image
                    image_elem.insert(list(image_elem).index(image_style), data_elem)
                    print("➕ Dodano <pt:data>")
                else:
                    data_elem.text = b64_image
                    print("✏️ Nadpisano <pt:data>")
                found = True
                break

        if found:
            tree.write(label_path, encoding="utf-8", xml_declaration=True)
            if modified_file.exists():
                modified_file.unlink()

            shutil.make_archive(temp_zip.replace(".zip", ""), 'zip', work_dir)
            os.rename(temp_zip, temp_lbx)
            shutil.move(temp_lbx, modified_file)
            print(f"✅ Plik zapisano w: {modified_file}")
        else:
            print("❌ Nie znaleziono elementu Obraz3")

    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
        if bmp_file.exists():
            try:
                bmp_file.unlink()
                print(f"🧹 Usunięto tymczasowy plik: {bmp_file}")
            except Exception as e:
                print(f"⚠️ Nie udało się usunąć pliku BMP: {e}")


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


def odpalarka24mm():
    if identyfikacjaROD_czy_RESO() == "ROD-11" or identyfikacjaROD_czy_RESO() == "ROD-12" or identyfikacjaROD_czy_RESO() == "ROD-14":

        # Ustawienia prostokąta w milimetrach
        width_mm, height_mm = 43, 25  # szerokosc wysokosc
        # -----------------------2424242424242424-----------------------------------------
        # Tworzenie instancji dla prostokąta 24 mm
        rectangle = RectangleImage(width_mm, height_mm, font_name="arialbd.ttf")
        # Dodawanie stałych elementów (GNIAZDO SERWIS.)
        font_size_top = 35
        rectangle.add_rectangle(argon="GN", freon=" ", gold="GNIAZDO SERWISOWE", rect_width_mm=50, font_size_top=35,
                                font_size_center=40)

        # Pętla do dodawania ruchomych elementów (Opisy obwodów, złączki odpływowe, styczniki, etc)
        dane_lista = daneL123STVLD()
        rodzajBezpieki = rodzajZabezpieczeniaSTV_VLD_S30_()
        for c in range(iloscObwodow):
            zabezpieczenie = rodzajBezpieki[c]
            prąd = dane_lista[c]
            opis = str(opisyObwodow24mm()[c])
            etykieta_lewa = str(daneTasma24mmLewyGornyRog()[c])
            wartosc_bezpiecznika = str(wartosciZabezpieczen()[c]) + "A"

            # Oblicz szerokość prostokąta i dobierz czcionkę
            szer_mm = oblicz_szerokosc(prąd, zabezpieczenie)

            if szer_mm == 17:  # VLD 1p
                font_size_center = 24
                font_size_top = 35
                max_znaki_na_linie = 12
                szer_mm = 20
            elif szer_mm == 27:  # STV 1p
                font_size_center = 40
                max_znaki_na_linie = 12
            elif szer_mm == 37:  # VLD 2p
                font_size_center = 35
                font_size_top = 35
                max_znaki_na_linie = 20
            # elif szer_mm == 54:  # STV 2p
            elif szer_mm == 60:  # STV 2p
                font_size_center = 35
                font_size_top = 35
                max_znaki_na_linie = 20
            elif szer_mm == 58:  # VLD 3p
                font_size_center = 40
                max_znaki_na_linie = 30
            elif szer_mm == 89:  # STV 3p
                font_size_center = 40
                max_znaki_na_linie = 30
            else:
                font_size_center = 40
                max_znaki_na_linie = 25

            # Podziel tekst tylko raz
            fragmenty = opis.split("^")  # jeśli tekst za długo w 1p to ^ go łamie
            linie = []

            for i, fragment in enumerate(fragmenty):
                if i > 0:
                    # Dodaj "-" tam, gdzie było "^" (czyli na początku kolejnego fragmentu)
                    fragment = "- " + fragment.lstrip()
                linie.extend(split_text(fragment, max_znaki_na_linie))

            gold_text = "\n".join(linie)

            # Dodaj prostokąt
            rectangle.add_rectangle(
                argon=etykieta_lewa,
                freon=wartosc_bezpiecznika,
                gold=gold_text,
                rect_width_mm=szer_mm,
                font_size_top=font_size_top,
                font_size_center=font_size_center
            )
            # dodaje RDC jeśli jest
            if identyfikacjaRCD()[c - h] == "tak" and (szer_mm == 58 or szer_mm == 89):
                rectangle.add_rectangle(argon=RCDoznaczenie()[c], freon=str(RCDcurrent()[c]) + "A", gold=gold_text,
                                        rect_width_mm=76,
                                        font_size_top=35,
                                        font_size_center=40)

            elif identyfikacjaRCD()[c - h] == "tak" and (szer_mm == 20 or szer_mm == 27):
                rectangle.add_rectangle(argon=RCDoznaczenie()[c], freon=str(RCDcurrent()[c]) + "A", gold=gold_text,
                                        rect_width_mm=41,
                                        font_size_top=35,
                                        font_size_center=40)

        rectangle.add_rectangle(argon="F5", freon="16A", gold="OŚWIET.ROZDZ.\n GRZEJNIK \n GNIAZDO", rect_width_mm=41,
                                font_size_top=35,
                                font_size_center=40)
        rectangle.add_rectangle(argon="F6", freon="6A", gold="ZASILANIE \n STEROWANIA\n 230V;50Hz", rect_width_mm=20,
                                font_size_top=35,
                                font_size_center=25)
        rectangle.add_rectangle(argon="KZ", freon="", gold="PRZEKAŹNIK\n ZMIERZCHOWY", rect_width_mm=41,
                                font_size_top=35,
                                font_size_center=40)
        rectangle.add_przelacznikSterowania("SL", (
            " PZEŁĄCZNIK SL\n"
            "0-BLOKADA\n"
            "1-STEROWANIE\n"
            "2-STER.RĘCZNE ZMIERZCHÓWKA\n"
            "3-STER.RĘCZNE ZAŁĄCZENIE"
        ))
        # Zapis obrazu do pliku

        update_obraz_lbx(rectangle)

    # RESO
    if identyfikacjaROD_czy_RESO() == "RESO-3F10":
        # Ustawienia prostokąta w milimetrach
        width_mm, height_mm = 43, 25  # szerokosc wysokosc
        # -----------------------2424242424242424-----------------------------------------
        # Tworzenie instancji dla prostokąta 24 mm
        rectangle = RectangleImage(width_mm, height_mm, font_name="arialbd.ttf")
        # Dodawanie stałych elementów (GNIAZDO SERWIS.)
        font_size_top = 35
        rectangle.add_rectangle(argon="GN", freon=" ", gold="GNIAZDO SERWISOWE", rect_width_mm=50, font_size_top=35,
                                font_size_center=40)
        rectangle.add_rectangle(argon="F5", freon="16A", gold="OŚWIET.ROZDZ.\n GRZEJNIK \n GNIAZDO", rect_width_mm=41,
                                font_size_top=35,
                                font_size_center=40)
        # rectangle.add_rectangle(argon="F6", freon="6A", gold="ZASILANIE \n STEROWANIA\n 230V;50Hz", rect_width_mm=54,
        rectangle.add_rectangle(argon="F6", freon="6A", gold="ZASILANIE \n STEROWANIA\n 230V;50Hz", rect_width_mm=60,
                                font_size_top=35,
                                font_size_center=40)
        # rectangle.add_rectangle(argon="KZ", freon="", gold="PRZEKAŹNIK\n ZMIERZCHOWY", rect_width_mm=41, font_size_top=35,
        #                         font_size_center=40)
        if l != 0 and h == 0:
            rectangle.add_przelacznikSterowania("SL", (
                " PZEŁĄCZNIK SL\n"
                "0-BLOKADA\n"
                "1-STEROWANIE\n"
                "2-STER.RĘCZNE ZMIERZCHÓWKA\n"
                "3-STER.RĘCZNE ZALACZENIE"
            ))
        elif l == 0 and h != 0:
            rectangle.add_przelacznikSterowania("SH", (
                " PZEŁĄCZNIK SH\n"
                "               0-BLOKADA\n"
                "               1-STEROWANIE\n"
                "               2-AWARYJNE"
            ))
        elif h != 0 and l != 0:

            rectangle.add_przelacznikSterowania("SL", (
                " PRZEŁĄCZNIK SL\n"
                "0-BLOKADA\n"
                "1-STEROWANIE\n"
                "2-STER.RĘCZNE ZMIERZCHÓWKA\n"
                "3-STER.RĘCZNE ZALACZENIE"
            ))
            rectangle.add_przelacznikSterowania("SH", (
                " PZEŁĄCZNIK SH\n"
                "           0-BLOKADA\n"
                "           1-STEROWANIE\n"
                "           2-AWARYJNE"
            ), line_spacing=15)

        # Pętla do dodawania ruchomych elementów (Opisy obwodów, złączki odpływowe, styczniki, etc)
        dane_lista = daneL123STVLD()

        rodzajBezpieki = rodzajZabezpieczeniaSTV_VLD_S30_()
        for c in range(iloscObwodow):
            zabezpieczenie = rodzajBezpieki[c]
            prad = dane_lista[c]
            opis = str(opisyObwodow24mm()[c])
            etykieta_lewa = str(daneTasma24mmLewyGornyRog()[c])
            wartosc_bezpiecznika = str(wartosciZabezpieczen()[c]) + "A"

            # Oblicz szerokość prostokąta i dobierz czcionkę
            szer_mm = oblicz_szerokosc(prad, zabezpieczenie)
            # print(szer_mm)
            if szer_mm == 17:  # VLD 1p
                # font_size_center = 27
                # font_size_top = 25
                font_size_center = 24
                font_size_top = 35
                max_znaki_na_linie = 12
                szer_mm == 20
            elif szer_mm == 27:  # STV 1p
                font_size_center = 40
                max_znaki_na_linie = 12
            elif szer_mm == 37:  # VLD 2p
                font_size_center = 35
                font_size_top = 35
                max_znaki_na_linie = 20
            elif szer_mm == 60:  # STV 2p
                # elif szer_mm == 'STV 2p':  # STV 2p
                font_size_center = 35
                font_size_top = 35
                max_znaki_na_linie = 20
            elif szer_mm == 58:  # VLD 3p
                font_size_center = 35
                max_znaki_na_linie = 30
            elif szer_mm == 89:  # STV 3p
                font_size_center = 40
                max_znaki_na_linie = 30
            else:
                font_size_center = 40
                max_znaki_na_linie = 25

            # Podziel tekst tylko raz
            fragmenty = opis.split("^")  # jeśli tekst za długo w 1p to ^ go łamie
            linie = []

            for i, fragment in enumerate(fragmenty):
                if i > 0:
                    # Dodaj "-" tam, gdzie było "^" (czyli na początku kolejnego fragmentu)
                    fragment = "- " + fragment.lstrip()
                linie.extend(split_text(fragment, max_znaki_na_linie))

            gold_text = "\n".join(linie)

            # Dodaj prostokąt
            rectangle.add_rectangle(
                argon=etykieta_lewa,
                freon=wartosc_bezpiecznika,
                gold=gold_text,
                rect_width_mm=szer_mm,
                font_size_top=font_size_top,
                font_size_center=font_size_center
            )
            # dodaje RDC jeśli jest

            try:
                if identyfikacjaRCD()[c - h] == "tak" and (szer_mm == 58 or szer_mm == 89):
                    rectangle.add_rectangle(argon=RCDoznaczenie()[c], freon=str(RCDcurrent()[c]) + "A", gold=gold_text,
                                            rect_width_mm=76,
                                            font_size_top=35,
                                            font_size_center=40)
                elif identyfikacjaRCD()[c - h] == "tak" and (szer_mm == 20 or szer_mm == 27):
                    rectangle.add_rectangle(argon=RCDoznaczenie()[c], freon=str(RCDcurrent()[c]) + "A", gold=gold_text,
                                            rect_width_mm=41,
                                            font_size_top=35,
                                            font_size_center=40)
            except:
                IndexError

            # Zapis obrazu do pliku

            # filename = r"\\fs1\Oddział Energetyki kolejowej\EK\8_Produkcja\4_Naklejki do rozdzielnic\testowy\AutoLabel\24mmPt1.bmp"
            # rectangle.save_image(filename)
            # temp_dir = Path(os.environ['TEMP'])
            # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            #
            # filename = temp_dir / f"24mmPt1_{timestamp}.bmp"
            # rectangle.save_image(str(filename))  # zapis pliku BMP
            update_obraz_lbx(rectangle)
