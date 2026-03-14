import re

from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt



class ComputedStyle:
    """Класс для хранения вычисленных стилей элемента с учетом всей цепочки наследования"""

    def __init__(self, element, parent_computed_style=None):
        self.element = element
        self.styles = {}

        # Начинаем с наследуемых стилей от родителя (которые уже включают всю цепочку)
        if parent_computed_style:
            # Копируем только наследуемые свойства от родителя
            self._inherit_from_parent(parent_computed_style)

        # Применяем собственные стили элемента (переопределяют родительские)
        self._parse_inline_styles()
        # Парсим атрибуты HTML в стили
        self._parsse_attribute_styles()

        # Применяем стили от тега (например, h1, h2 и т.д.)
        self._apply_tag_defaults()

    def _inherit_from_parent(self, parent_style):
        """Наследует соответствующие свойства от родителя"""
        # Список свойств, которые наследуются в CSS
        inheritable_properties = [
            "font-size",
            "font-family",
            "font-weight",
            "font-style",
            "color",
            "text-align",
            "text-indent",
            "text-decoration",
            "line-height",
            "letter-spacing",
            "word-spacing",
            "visibility",
            "white-space",
        ]

        for prop in inheritable_properties:
            if prop in parent_style.styles:
                self.styles[prop] = parent_style.styles[prop]

    def _parse_inline_styles(self):
        """Парсит inline-стили элемента"""
        style_attr = self.element.get("style", "")

        if not style_attr:
            return

        for item in style_attr.split(";"):
            item = item.strip()
            if not item:
                continue

            if ":" in item:
                name, value = [part.strip() for part in item.split(":", 1)]
                self.styles[name] = value

                # Логируем для отладки


    def _parsse_attribute_styles(self):
        for attr, value in self.element.attrs.items():
            if attr == "border":
                self.styles["border"] = value
            elif attr == "cellpadding":
                self.styles["padding"] = f"{value}px"
            elif attr == "cellspacing":
                self.styles["border-spacing"] = f"{value}px"
            elif attr == "bgcolor":
                self.styles["background-color"] = value
            elif attr == "width":
                if str(value).endswith("%"):
                    self.styles["width"] = value
                else:
                    self.styles["width"] = f"{value}px"
            elif attr == "height":
                if str(value).endswith("%"):
                    self.styles["height"] = value
                else:
                    self.styles["height"] = f"{value}px"
            elif attr == "align":
                self.styles["align"] = value
            elif attr == "valign":
                self.styles["vertical-align"] = value

    def _apply_tag_defaults(self):
        """Применяет стили по умолчанию для тегов"""
        tag_defaults = {
            "h1": {"font-size": "24pt", "font-weight": "bold", "margin-bottom": "12pt"},
            "h2": {"font-size": "18pt", "font-weight": "bold", "margin-bottom": "10pt"},
            "h3": {"font-size": "16pt", "font-weight": "bold", "margin-bottom": "8pt"},
            "h4": {"font-size": "14pt", "font-weight": "bold", "margin-bottom": "6pt"},
            "h5": {"font-size": "12pt", "font-weight": "bold", "margin-bottom": "4pt"},
            "h6": {"font-size": "10pt", "font-weight": "bold", "margin-bottom": "2pt"},
            "strong": {"font-weight": "bold"},
            "b": {"font-weight": "bold"},
            "em": {"font-style": "italic"},
            "i": {"font-style": "italic"},
            "u": {"text-decoration": "underline"},
            "p": {"margin-bottom": "10pt"},
            "td": {"vertical-align": "top"},
            "th": {"vertical-align": "middle", "font-weight": "bold"},
        }

        my_name = self.element.name
        if self.element.name in tag_defaults:
            for prop, value in tag_defaults[self.element.name].items():
                if prop not in self.styles:  # Не переопределяем уже заданные стили
                    self.styles[prop] = value

    def get(self, property_name, default=None):
        """Получает значение свойства стиля"""
        return self.styles.get(property_name, default)

    def get_numeric_value(self, property_name, default_value=None, default_unit="pt"):
        """Получает числовое значение свойства с единицами измерения"""
        value = self.get(property_name)
        if not value:
            return default_value, default_unit

        # Извлекаем число и единицы измерения
        match = re.search(r"([\d.]+)\s*([a-z%]*)", value)
        if not match:
            return default_value, default_unit

        num = float(match.group(1))
        unit = match.group(2) or default_unit

        return num, unit

    def get_font_size(self, parent_style=None):
        """Получает размер шрифта в пунктах с учетом наследования"""
        value = self.get("font-size")

        # Если нет своего размера, используем родительский (уже должен быть в стилях)
        if not value:
            # Пробуем найти в унаследованных стилях
            if "font-size" in self.styles:
                value = self.styles["font-size"]
            else:
                return Pt(11)  # Размер по умолчанию

        match = re.search(r"([\d.]+)\s*([a-z%]*)", value)
        if not match:
            return Pt(11)

        num = float(match.group(1))
        unit = match.group(2)

        if unit == "px":
            return Pt(num * 0.75)  # Примерное преобразование
        elif unit == "pt":
            return Pt(num)
        elif unit == "%":
            # Проценты от родительского размера
            # Нужно получить родительский размер
            if parent_style and "font-size" in parent_style.styles:
                parent_value = parent_style.styles["font-size"]
                parent_match = re.search(r"([\d.]+)\s*([a-z%]*)", parent_value)
                if parent_match:
                    parent_num = float(parent_match.group(1))
                    parent_unit = parent_match.group(2)

                    if parent_unit == "pt":
                        return Pt(parent_num * num / 100)
                    elif parent_unit == "px":
                        return Pt(parent_num * 0.75 * num / 100)

            return Pt(11 * num / 100)
        elif unit == "em":
            # em относительно родительского размера
            if parent_style and "font-size" in parent_style.styles:
                parent_value = parent_style.styles["font-size"]
                parent_match = re.search(r"([\d.]+)\s*([a-z%]*)", parent_value)
                if parent_match:
                    parent_num = float(parent_match.group(1))
                    parent_unit = parent_match.group(2)

                    if parent_unit == "pt":
                        return Pt(parent_num * num)
                    elif parent_unit == "px":
                        return Pt(parent_num * 0.75 * num)

            return Pt(11 * num)
        else:
            return Pt(num)

    def get_text_align(self):
        """Получает выравнивание текста"""
        align_map = {
            "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        }
        value = self.get("text-align")
        return align_map.get(value, WD_PARAGRAPH_ALIGNMENT.LEFT)

    def get_vertical_align(self):
        """Получает вертикальное выравнивание"""
        value = self.get("vertical-align")
        return value if value else "top"

    def get_background_color(self):
        """Получает цвет фона"""
        return self.get("background-color")

    def get_color(self):
        """Получает цвет текста"""
        return self.get("color")

    def get_font_family(self):
        """Получает семейство шрифта"""
        return self.get("font-family", "Times New Roman")

    def get_border(self):
        """Получает границы ячейки"""
        border = self.get("border", "0")
        if "px" in border:
            border = border.replace("px", "")
        return int(border)

    def is_bold(self):
        """Проверяет, должен ли текст быть жирным"""
        weight = self.get("font-weight")
        return weight in ("bold", "bolder", "700", "800", "900") or (
            weight and weight.isdigit() and int(weight) >= 700
        )

    def is_italic(self):
        """Проверяет, должен ли текст быть курсивным"""
        style = self.get("font-style")
        return style == "italic"

    def is_underlined(self):
        """Проверяет, должен ли текст быть подчеркнутым"""
        decor = self.get("text-decoration")
        return decor == "underline"