from docx.shared import Pt, Cm

# Пути (заполняются автоматически в main.py при сканировании папки)
INPUT_MD = None
OUTPUT_DOCX = None
TEMPLATE_DOCX = None
IMAGES_DIR = "images"

# Оформление текста
FONT = "Times New Roman"
FONT_SIZE = Pt(14)
LINE_SPACING = 1.5
FIRST_LINE_INDENT = Cm(1.25)

# Поля страницы
MARGINS = {
    "left": Cm(3),
    "right": Cm(1.5),
    "top": Cm(2),
    "bottom": Cm(2),
}

# Структурные элементы (печатаются ПРОПИСНЫМИ, по центру, с новой страницы)
STRUCTURAL_HEADINGS = {
    "введение",
    "заключение",
    "список литературы",
    "содержание",
}

# Флаги
STRICT_GOST_LEVEL1 = False      # True: главы слева с отступом, False: по центру
AUTO_NUMBER_FORMULAS = True     # Автонумерация выключных формул (1), (2)...

# Ширина текста на A4: 21 - 3 - 1.5 = 16.5 см
TEXT_WIDTH_CM = 16.5