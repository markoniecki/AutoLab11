import pandas as pd
import numpy as np
import os, time
from pathlib import Path
# from ROD_11_.exporterPlikuTxT import dawajZawartoscPlikuTekstowego

# def kloppe():
#     # Ścieżka do folderu z plikiem
#     folder_path = dawajZawartoscPlikuTekstowego()
#
#     # Pełna ścieżka do pliku
#     plik = 'tabela_material.xlsm'
#     full_path = os.path.join(folder_path, plik)
#
#     # Sprawdzenie, czy plik istnieje
#     if not os.path.exists(full_path):
#         raise FileNotFoundError(f"❌ Plik {full_path} nie został znaleziony!")
#
#     return full_path

# def slope():
#     # Zmiana katalogu roboczego na katalog 'util'
#     os.chdir(os.path.join(os.path.dirname(__file__)))
#
#     plik = 'tabela_material.xlsm'
#     # print("Sprawdzam, czy plik istnieje:", plik)
#     # print(os.path.exists(plik))  # Sprawdzamy, czy plik istnieje
#
#     if not os.path.exists(plik):
#         raise FileNotFoundError(f"❌ Plik {plik} nie został znaleziony!")
#
#     # Reszta kodu do obsługi pliku Excel
#     # print(f"✅ Otwieram plik: {plik}")
#     return plik

plik = 'tabela_material.xlsm'
# folder_path = dawajZawartoscPlikuTekstowego()
# full_path = os.path.join(folder_path, plik)
# weźmij ścieżkę do zapisu z TEMP cotton.txt
# def get_excel_folder_from_temp():
#     temp_dir = Path(os.environ['TEMP'])  # lub %TMP%
#     coton_path = temp_dir / "coton.txt"
#
#
#     if not coton_path.exists():
#         raise FileNotFoundError(f"Nie znaleziono pliku: {coton_path}")
#
#     with open(coton_path, 'r', encoding='cp1250') as f:
#         folder_path = f.read().strip()
#
#     return folder_path
def get_excel_folder_from_temp():
    temp_dir = Path(os.environ['TEMP'])
    coton_path = temp_dir / "coton.txt"

    # Czekamy maksymalnie 5 sekund na pojawienie się pliku
    timeout = 5
    waited = 0
    while not coton_path.exists():
        time.sleep(0.2)
        waited += 0.2
        if waited >= timeout:
            raise FileNotFoundError(f"❌ Nie znaleziono pliku: {coton_path}")

    # Teraz otwieramy plik
    with open(coton_path, 'r', encoding='cp1250') as f:
        folder_path = f.read().strip()

    return folder_path

# def get_full_excel_path():
#     # Ścieżka do coton.txt (na sztywno)
#     coton_path = r"\\fs1\\Oddział Energetyki kolejowej\\EK\\8_Produkcja\\4_Naklejki do rozdzielnic\\testowy\\AutoLabel\\coton.txt"
#
#     # Odczytaj ścieżkę folderu z coton.txt (obsługuje BOM)
#     with open(coton_path, "r", encoding="cp1250") as f:
#     # with open(coton_path, "r", encoding="utf-8-sig") as f:
#         folder_path = f.readline().strip()
#
#     # Zbuduj pełną ścieżkę do pliku Excela
#     full_path = os.path.join(folder_path, "tabela_material.xlsm")
#     # full_path = os.path.join(folder_path)
#     return full_path
full_path = os.path.join(get_excel_folder_from_temp(), "tabela_material.xlsm")


if not os.path.exists(full_path):
    raise FileNotFoundError(f"❌ Plik {full_path} nie został znaleziony!")

# odczytaj ilość obwodów L H N w ogóle - do iteracji
dfOgolneH = pd.read_excel(full_path, sheet_name='AutoLabel', usecols='D', header=None, skiprows=5, nrows=1)
dfOgolneL = pd.read_excel(full_path, sheet_name='AutoLabel', usecols='D', header=None, skiprows=6, nrows=1)
dfOgolneN = pd.read_excel(full_path, sheet_name='AutoLabel', usecols='D', header=None, skiprows=7, nrows=1)
h = dfOgolneH.iloc[0, 0]
l = dfOgolneL.iloc[0, 0]
n = dfOgolneN.iloc[0, 0]
# Sprawdzanie i przypisywanie 0 jeśli wartość to NaN
h = 0 if np.isnan(h) else h
l = 0 if np.isnan(l) else l
n = 0 if np.isnan(n) else n

iloscObwodow = h + l + n  # wartość do iteracji wszystkich obwodów

# odczyt oznaczeń zabezpieczeń | Lewy górny róg rectanglea

df = pd.read_excel(full_path, sheet_name='AutoLabel', header=None)


def identyfikacjaROD_czy_RESO():
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    # Odczyt wartości z komórki B9 (8 wiersz, 1 kolumna w indeksowaniu 0)
    komorka_d5 = df.iloc[3, 3]
    return komorka_d5


def identyfikacja_zasilania_1_3_faz():
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    # Odczyt wartości z komórki B9 (8 wiersz, 1 kolumna w indeksowaniu 0)
    komorka_dx = df.iloc[7, 3]
    return komorka_dx


def identyfikacja_czy_ARI():
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    komorka_d12 = df.iloc[10, 3]
    return komorka_d12

def identyfikacja_czy_extra_Modem():
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    komorka_f13 = df.iloc[11, 5]
    return komorka_f13

def jaka_szafa_wariacie():
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    komorka_d10 = df.iloc[8, 3]
    return komorka_d10

def identyfikacja_czy_liczy_H_czy_przekladnia():
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    komorka_d39 = df.iloc[37, 3]
    komorka_d40 = df.iloc[38, 3]
    return [komorka_d39, komorka_d40]

def identyfikacja_czy_liczy_L_czy_przekladnia():
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    komorka_d42 = df.iloc[40, 3]
    komorka_d43 = df.iloc[41, 3]
    return [komorka_d42, komorka_d43]

def identyfikacja_extra_blok_rozdzielczy_A():
    output = []
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    extra_blok_H_A = df.iloc[15, 3]  # H
    extra_blok_L_A = df.iloc[16, 3]  # L
    extra_blok_N_A = df.iloc[17, 3]  # N
    output.append(extra_blok_H_A)
    output.append(extra_blok_L_A)
    output.append(extra_blok_N_A)
    return output


def identyfikacja_extra_blok_rozdzielczy_B():
    output = []
    df = pd.read_excel(full_path, sheet_name="AutoLabel", engine="openpyxl")
    extra_blok_H_B = df.iloc[15, 4]  # H
    extra_blok_L_B = df.iloc[16, 4]  # L
    extra_blok_N_B = df.iloc[17, 4]  # N
    output.append(extra_blok_H_B)
    output.append(extra_blok_L_B)
    output.append(extra_blok_N_B)
    return output

def daneTasma24mmLewyGornyRog():  # ['1FH1', '2FH1', '3FH1', '4FH1', '1FL1', '2FL1', ...]
    # Wiersz 28 to index 27, kolumna E to index 4
    wiersz_index = 29
    start_kolumna_index = 4
    tasma24mmLewyGornyRog = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        tasma24mmLewyGornyRog.append(wartosc)
    return tasma24mmLewyGornyRog  # ['1FH1', '2FH1', '3FH1', '4FH1', '1FL1', '2FL1', ...]

# def opisyObwodow24mm():
#     # Wiersz 21 to index 20, kolumna E to index 4
#     wiersz_index = 22
#     start_kolumna_index = 4
#     opis24mmCentrumTasmy = []
#     for i in range(iloscObwodow):
#         kolumna_index = start_kolumna_index + i
#         wartosc = df.iloc[wiersz_index, kolumna_index]
#         if isinstance(wartosc, str) and "opornic" in wartosc:
#             wartosc = wartosc.replace("opornic", "OPORNIC ROZJAZD")
#         opis24mmCentrumTasmy.append(wartosc)
#     return opis24mmCentrumTasmy
def opisyObwodow24mm():
    # Wiersz 21 to index 20, kolumna E to index 4
    wiersz_index = 22
    start_kolumna_index = 4
    opis24mmCentrumTasmy = []
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]

        if isinstance(wartosc, str):
            wartosc = wartosc.upper()  # Zamień na wielkie litery
            wartosc = wartosc.replace("OPORNIC", "OPORNIC ROZJAZD")

        opis24mmCentrumTasmy.append(wartosc)

    return opis24mmCentrumTasmy

# def opisyObwodow24mm():  # ['Ogrzewanie opornic 1', 'Ogrzewanie opornic 2', 'Ogrzewanie opornic 3', 'Ogrzewanie zamknięć 1,2,3', 'Oświetlenie peron 1', ...]
#     # Wiersz 21 to index 20, kolumna E to index 4
#     wiersz_index = 22
#     start_kolumna_index = 4
#     opis24mmCentrumTasmy = []
#     # Iteracja po odpowiednich kolumnach
#     for i in range(iloscObwodow):
#         kolumna_index = start_kolumna_index + i
#         wartosc = df.iloc[wiersz_index, kolumna_index]
#         print(wartosc)
#         opis24mmCentrumTasmy.append(wartosc)
#     return opis24mmCentrumTasmy


def daneL123STVLD():  # ['D1', 'D2', 'D3', 'L1 | L2 | L3', 'L1 | L2 | L3', 'L1', 'L2', 'L1 | L2', 'L1 | L2 | L3', 'L1 | L2 | L3',...]
    # Wiersz 30 to index 31, kolumna E to index 4
    wiersz_index = 31
    start_kolumna_index = 4
    daneSzerokosciAparatu = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        daneSzerokosciAparatu.append(wartosc)
    return daneSzerokosciAparatu


def rodzajZabezpieczeniaSTV_VLD_S30_():  # ['STV', 'STV', 'STV', 'VLD', 'VLD', 'VLD', 'VLD', 'S_300B
    # Wiersz 24 to index 23, kolumna E to index 4
    wiersz_index = 23
    start_kolumna_index = 4
    daneRodzajuAparatu = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        daneRodzajuAparatu.append(wartosc)
    return daneRodzajuAparatu


# def generujeNapisyNaZlaczkiOdplywowe():  # ['1XH', '2XH', '3XH', '4XH', '1XL', '2XL', '3XL', '4XL', '5XL' ...]
#     # Wiersz 27 to index 26, kolumna E to index 4
#     wiersz_index = 26
#     start_kolumna_index = 4
#     daneOpis24mm = []
#     # Iteracja po odpowiednich kolumnach
#     for i in range(iloscObwodow):
#         kolumna_index = start_kolumna_index + i
#         wartosc = df.iloc[wiersz_index, kolumna_index]
#         daneOpis24mm.append(wartosc)
#     return daneOpis24mm
def generujeNapisyNaZlaczkiOdplywowe():  # ['1XH', '2XH', '3XH', '4XH', '1XL', '2XL', '3XL', '4XL', '5XL' ...]
    # Wiersz 27 to index 26, kolumna E to index 4
    wiersz_index = 26
    start_kolumna_index = 4
    daneOpis24mm = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        if 'ZAMKNIĘĆ' not in opisyObwodow24mm()[i]:
            daneOpis24mm.append(wartosc)
    return daneOpis24mm

def wartosciZabezpieczen():  # [6, 6, 16, 16, 16, 16, 16, 16, 20 ...]
    # Wiersz 25 to index 24, kolumna E to index 4
    wiersz_index = 24
    start_kolumna_index = 4
    daneAmper = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        daneAmper.append(wartosc)
    return daneAmper


def identyfikacjaDali():
    df = pd.read_excel(full_path, sheet_name='Dane', header=None)
    wiersz_index = 33
    start_kolumna_index = 3
    daneOpis24mm = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        daneOpis24mm.append(wartosc)
    return daneOpis24mm


def opisy6mmNaDali():
    mops = identyfikacjaDali()
    hops = []
    x = []
    for g in range(h):
        mops.pop(0)
    for i in range(l):
        if mops[i] == "Dali":
            a = "A7" + str(i)
            b = "KD" + str(i + 1)
            hops.append(a)
            hops.append(b)
    return hops


def identyfikacjaRCD():  # [' ', ' ', 'nie', 'nie', 'tak', ' ', ' ', 'tak', ' ', ' ']
    # Wiersz 34 to index 33, kolumna E to index 4
    wiersz_index = 33
    start_kolumna_index = 4
    czyRCD = []
    takRCD = []
    RCD = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        czyRCD.append(wartosc)
    for j in czyRCD[h:]:
        takRCD.append(j)
    return takRCD
    # for k in range(l):
    #     if takRCD[k]=="tak"
    #         RCD.append()

def RCDcurrent():  # [nan, nan, nan, nan, 20, nan, 20, nan, nan, nan]
    wiersz_index = 34
    start_kolumna_index = 4
    daneAmperRCD = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        daneAmperRCD.append(wartosc)
    return daneAmperRCD


def RCDoznaczenie():  # [nan, nan, 2FL2, nan, 1FN2, nan ...]
    wiersz_index = 35
    start_kolumna_index = 4
    RCDoznakowanie = []
    # Iteracja po odpowiednich kolumnach
    for i in range(iloscObwodow):
        kolumna_index = start_kolumna_index + i
        wartosc = df.iloc[wiersz_index, kolumna_index]
        RCDoznakowanie.append(wartosc)
    return RCDoznakowanie


def oblicz_szerokosc(fazy, rozlacznik_bezp):
    # print(fazy)
    """
    Funkcja oblicza szerokość na podstawie liczby faz i typu rozłącznika.

    :param fazy: lista reprezentująca fazy
    :param rozlacznik_bezp: typ rozłącznika, np. "STV" lub "VLD"
    :return: szerokość (int) lub None, jeśli brak dopasowania
    """
    if len(fazy) == 12 and rozlacznik_bezp == "STV": # STV 3p
        return 89  # OK
    elif len(fazy) == 7 and rozlacznik_bezp == "STV": # STV 2p
        # return 54
        return 60
    elif (len(fazy) == 12 and rozlacznik_bezp == "VLD" or len(fazy) == 12 and rozlacznik_bezp == "S_300B" or len(
            fazy) == 12 and rozlacznik_bezp == "S_300C" or len(fazy) == 12 and rozlacznik_bezp == "FR_300"):# VLD 3p
        return 58
    elif (len(fazy) == 7 and rozlacznik_bezp == "VLD" or len(fazy) == 7 and rozlacznik_bezp == "S_300B" or len(
            fazy) == 7 and rozlacznik_bezp == "S_300C" or len(fazy) == 7 and rozlacznik_bezp == "FR_300"): #VLD 2p
        return 37
    # elif (len(fazy) == 2 and rozlacznik_bezp == "STV"): # STV 1p
    #     return 20
    elif (len(fazy) == 2 and rozlacznik_bezp == "VLD" or len(fazy) == 2 and rozlacznik_bezp == "S_300B" or len(
            fazy) == 2 and rozlacznik_bezp == "S_300C" or len(fazy) == 2 and rozlacznik_bezp == "FR_300"): # VLD 1p
        return 17
    elif (len(fazy) == 2 and rozlacznik_bezp == "STV" and fazy == "D1" or
          len(fazy) == 2 and rozlacznik_bezp == "STV" and fazy == "D2" or
          len(fazy) == 2 and rozlacznik_bezp == "STV" and fazy == "D3"):
        return 89
    elif (len(fazy) == 2 and rozlacznik_bezp == "STV" and fazy == "L1" or # STV 1p
          len(fazy) == 2 and rozlacznik_bezp == "STV" and fazy == "L2" or
          len(fazy) == 2 and rozlacznik_bezp == "STV" and fazy == "L3"):
        return 27
    else:
        return None
