from PIL import Image, ImageDraw, ImageFont


class RectangleImage96:
    def __init__(self, width_mm, height_mm, dpi=300, font_name="arialbd.ttf"):
        self.dpi = dpi
        self.width_mm = width_mm
        self.height_mm = height_mm

        # Przeliczanie milimetrów na piksele
        self.width_px = int(self.width_mm * self.dpi / 25.4)
        self.height_px = int(self.height_mm * self.dpi / 25.4)

        # Tworzenie obrazu o przeliczonej szerokości i wysokości
        self.total_width = 0
        self.image = Image.new("RGB", (self.total_width, self.height_px), "white")
        self.draw = ImageDraw.Draw(self.image)
        self.font_name = font_name
        self.rectangle_count = 0  # 👈 **Inicjalizacja licznika prostokątów**

    def _load_font(self, size):
        try:
            return ImageFont.truetype(self.font_name, size)
        except IOError:
            return ImageFont.load_default()

    def add_rectangle(self, text="Gold", rect_width_mm=None, font_size_center=20):
        # Przeliczenie szerokości prostokąta z milimetrów na piksele
        if rect_width_mm is None:
            rect_width_mm = self.width_mm
        rect_width_px = int(rect_width_mm * self.dpi / 25.4)

        # Ładowanie czcionki z podanego rozmiaru
        font_center = self._load_font(font_size_center)

        # Powiększanie obrazu o miejsce na nowy prostokąt
        self.expand_image(rect_width_px)
        offset_x = self.total_width - rect_width_px

        # Obliczanie wymiarów tekstu
        text_bbox = self.draw.textbbox((0, 0), text, font=font_center)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]

        # Wyliczenie środka prostokąta
        center_x = offset_x + rect_width_px // 2
        center_y = self.height_px // 2
        tajemnaZmienna = (text_height % 2) + 10  # do regulacji położenia tekstu w pionie
        # Przesunięcie tekstu o jedną linijkę do góry
        center_y -= tajemnaZmienna  # Zmniejsz `center_y` o wysokość tekstu

        # Rysowanie tekstu wyśrodkowanego w poziomie i przesuniętego w górę
        self.draw.text((center_x - text_width // 2, center_y - text_height // 2), text, fill="black", font=font_center)
        # Zwiększenie licznika obiektów
        self.rectangle_count += 1

    #
    def expand_image(self, rect_width_px):
        # Powiększanie obrazu o kolejny prostokąt
        new_image = Image.new("RGB", (self.total_width + rect_width_px, self.height_px), "white")
        new_image.paste(self.image, (0, 0))
        self.image = new_image
        self.total_width += rect_width_px
        self.draw = ImageDraw.Draw(self.image)

    def get_rectangle_count(self):
        """Zwraca liczbę dodanych prostokątów."""
        return self.rectangle_count

    def save_image(self, filename):
        self.image.save(filename)


def split_text(text, max_length):
    """
    Dzieli tekst na linie, wstawiając nową linię w miejscu ostatniej spacji w przypadku,
    gdy długość tekstu przekroczy max_length. Rekurencyjnie dzieli tekst na mniejsze linie.
    """
    # Jeśli długość tekstu jest mniejsza niż max_length, zwróć go bez zmian
    if len(text) < max_length:
        return [text]

    # Szukamy ostatniej spacji w obrębie max_length
    split_pos = text.rfind(' ', 0, max_length)

    # Jeśli nie znaleziono spacji, po prostu tniemy tekst na dwie części
    if split_pos == -1:
        split_pos = max_length

    # Dzielimy tekst na dwie części
    first_part = text[:split_pos]
    second_part = text[split_pos:].strip()  # Usuwamy spacje z początku drugiej części

    # Rekurencyjnie dzielimy drugą część
    return [first_part] + split_text(second_part, max_length)