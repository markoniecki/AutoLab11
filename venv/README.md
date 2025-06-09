# AutoLab11

**AutoLab11** to narzędzie automatyzujące generowanie etykiet `.lbx` dla rozdzielnic energetycznych, w tym:

- tekstowe etykiety 6mm, 9mm, 24mm (Brother P-Touch)
- graficzne oznaczenia obwodów (BMP → LBX)
- integrację z arkuszem Excela przez `coton.txt`

## Funkcje

- 🔄 Modyfikacja plików `.lbx` na podstawie danych wejściowych
- 🖨️ Generacja graficznych etykiet i ich osadzenie w `.lbx`
- 📁 Automatyczny zapis w katalogu Excela
- 🧼 Obsługa czyszczenia tymczasowych plików

## Struktura katalogów

```
AutoLab11/
├── main.py
├── ROD_11_/
│   ├── _mm6_Logic.py
│   ├── _mm6_Logic_long.py
│   ├── _mm9_Logic.py
│   ├── _24mm_Logic.py
│   ├── Szbx6mm.lbx
│ 	├── Szbx6mm_long.lbx
│ 	├── Szbx9mm.lbx
│ 	└── Szbx24mm.lbx
├── utils/
├── requirements.txt
├── .gitignore
└── README.md
```

## Uruchomienie

1. Uruchom `main.py` lub wersję `.exe`
2. Gotowe pliki `.lbx` zapiszą się obok pliku Excela

## Wymagania

- Python 3.8+
- Pillow, openpyxl, shutil, zipfile, xml.etree
- (opcjonalnie) PyInstaller do kompilacji do `.exe`

## Autor

Projekt tworzony i rozwijany wewnętrznie do zastosowań przemysłowych.