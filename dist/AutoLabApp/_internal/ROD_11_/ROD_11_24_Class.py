from PIL import Image, ImageDraw, ImageFont

class RectangleImage:
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
        # self.rectangle_count = 0  # 👈 **Inicjalizacja licznika prostokątów**

    def _load_font(self, size):
        try:
            return ImageFont.truetype(self.font_name, size)
        except IOError:
            return ImageFont.load_default()

    def _split_text(self, text, font, max_width):
        # Dzielenie tekstu na linie, które mieszczą się w max_width
        if not isinstance(text, str):
            text = str(text)  # Upewnij się, że text jest ciągiem znaków

        words = text.split(" ")
        lines = []
        current_line = words[0]

        for word in words[1:]:
            if self.draw.textbbox((0, 0), current_line + " " + word, font=font)[2] <= max_width:
                current_line += " " + word
            else:
                lines.append(current_line)
                current_line = word

        lines.append(current_line)
        return lines

    def add_rectangle(self, argon="Argon", freon="Freon", gold="Gold", rect_width_mm=None, font_size_top=20,
                      font_size_center=20):
        # Przeliczenie szerokości prostokąta z milimetrów na piksele
        if rect_width_mm is None:
            rect_width_mm = self.width_mm
        rect_width_px = int(rect_width_mm * self.dpi / 25.4)

        # Ładowanie czcionek z podanych rozmiarów
        font_top = self._load_font(font_size_top)
        font_center = self._load_font(font_size_center)

        # Powiększanie obrazu o miejsce na nowy prostokąt
        self.expand_image(rect_width_px)
        rectOdstep = 19  # parametr regulacyjny odstępu boxów
        offset_x = self.total_width - rect_width_px  # odstęp prostokątów

        # Rysowanie prostokąta
        gruboscLiniiRect = 3
        self.draw.rectangle((offset_x + rectOdstep, 5, offset_x + rect_width_px - rectOdstep, self.height_px - 5),
                            outline="black",
                            width=gruboscLiniiRect)

        # Rysowanie tekstu (górna sekcja)
        self.draw.text((offset_x + rectOdstep + 15, rectOdstep - 5), argon, fill="black",
                       font=font_top)  # wspolrzedne fontu 1

        # Rysowanie tekstu (prawy górny róg)
        text_bbox = self.draw.textbbox((0, 0), freon, font=font_top)
        text_width = text_bbox[2] - text_bbox[0]
        self.draw.text((offset_x + rect_width_px - text_width - rectOdstep - 15, rectOdstep - 5), freon,
                       fill="black",
                       font=font_top)  # wspolrzedne fontu 2

        # Rysowanie tekstu (środek, wyjustowany centralnie z uwzględnieniem nowej linii)
        lines = gold.split("\n")  # Podziel tekst na linie
        line_heights = []
        for line in lines:
            line_bbox = self.draw.textbbox((0, 0), line, font=font_center)
            line_heights.append(line_bbox[3] - line_bbox[1])  # Wysokość każdej linii

        total_text_height = sum(line_heights) + (len(lines) - 1) * 5  # Dodaj odstępy między liniami
        current_y = (self.height_px - total_text_height) // 2  # Pozycja startowa Y

        for line, line_height in zip(lines, line_heights):
            line_bbox = self.draw.textbbox((0, 0), line, font=font_center)
            line_width = line_bbox[2] - line_bbox[0]
            center_x = offset_x + rect_width_px // 2
            self.draw.text((center_x - line_width // 2, current_y), line, fill="black", font=font_center)
            current_y += line_height + 5  # Przejdź do następnej linii (z uwzględnieniem odstępu)

    def add_przelacznikSterowania(self, label, gold_text, font_size_title=45, font_size_sub=28, center_offset_y=20,
                                  title_spacing=30, line_spacing=7):
        # rect_width_mm = 50
        rect_width_mm = 54
        font_size_top = 35

        rect_width_px = int(rect_width_mm * self.dpi / 25.4)
        self.expand_image(rect_width_px)
        # rectOdstep = 20
        rectOdstep = 19
        offset_x = self.total_width - rect_width_px

        gruboscLiniiRect = 4
        self.draw.rectangle(
            (offset_x + rectOdstep, 5, offset_x + rect_width_px - rectOdstep, self.height_px - 5),
            outline="black",
            width=gruboscLiniiRect
        )

        font_top = self._load_font(font_size_top)
        font_title = self._load_font(font_size_title)
        font_sub = self._load_font(font_size_sub)

        self.draw.text((offset_x + rectOdstep + 15, rectOdstep - 5), label, fill="black", font=font_top)

        lines = gold_text.split("\n")
        title_line = lines[0]
        sub_lines = lines[1:]

        total_height = self.height_px - center_offset_y

        title_bbox = self.draw.textbbox((0, 0), title_line, font=font_title)
        title_width = title_bbox[2] - title_bbox[0]
        title_y = total_height // 2 - title_bbox[3] - title_spacing
        self.draw.text(
            (offset_x + rect_width_px // 2 - title_width // 2, title_y),
            title_line,
            fill="black",
            font=font_title
        )

        max_line_height = max(
            self.draw.textbbox((0, 0), line, font=font_sub)[3] -
            self.draw.textbbox((0, 0), line, font=font_sub)[1]
            for line in sub_lines
        )

        current_y = title_y + title_bbox[3] + title_spacing
        for sub_line in sub_lines:
            self.draw.text((offset_x + rectOdstep + 10, current_y), sub_line, fill="black", font=font_sub)
            current_y += max_line_height + line_spacing

    def expand_image(self, rect_width_px):
        # Powiększanie obrazu o kolejny prostokąt
        new_image = Image.new("RGB", (self.total_width + rect_width_px, self.height_px), "white")
        new_image.paste(self.image, (0, 0))
        self.image = new_image
        self.total_width += rect_width_px
        self.draw = ImageDraw.Draw(self.image)

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

