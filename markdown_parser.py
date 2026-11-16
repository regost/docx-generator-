# markdown_parser.py

import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import STRUCTURAL_HEADINGS

INLINE_MATH_RE = re.compile(r'(?<!\\)\$([^$\n]+?)\$')
IMAGE_RE = re.compile(r'!\[(.*?)\]\((.*?)\)')
DISPLAY_MATH_RE = re.compile(r'^\s*\$\$(.+?)\$\$\s*(?:\(([-\d.]+)\))?\s*$', re.DOTALL)


def strip_front_matter(text):
    if text.startswith("---"):
        parts = text.split("---", 2)
        if len(parts) >= 3:
            return parts[2]
    return text


def normalize_latex_delimiters(text: str) -> str:
    """Заменяет стандартные LaTeX-делимитеры \[...\] и \(...\) на Markdown-формат."""
    # 1. Блочные: \[ ... \]  ->  $$ ... $$
    text = re.sub(r'(?<!\\)\\\[(.*?)\\\]', r'$$\1$$', text, flags=re.DOTALL)
    # 2. Строчные: \( ... \)  ->  $ ... $
    text = re.sub(r'(?<!\\)\\\((.*?)\\\)', r'$\1$', text)
    return text


def collapse_display_math(text):
    """Склеивает многострочные $$...$$ в одну строку."""
    return re.sub(r'\$\$(.+?)\$\$', lambda m: '$$' + re.sub(r'\s+', ' ', m.group(1)).strip() + '$$', text, flags=re.DOTALL)


def is_table_separator(line: str) -> bool:
    return bool(re.match(r'^\s*\|?[\s:\-]+\|[\s:\-|]*$', line.strip()))


def split_md_row(row: str):
    row = row.strip()
    if row.startswith("|"): row = row[1:]
    if row.endswith("|"): row = row[:-1]
    
    cells = []
    current_cell = []
    in_math = False
    
    i = 0
    while i < len(row):
        char = row[i]
        
        # Отслеживаем вход/выход из математики ($...$)
        if char == '$' and (i == 0 or row[i-1] != '\\'):
            in_math = not in_math
            current_cell.append(char)
            
        # Если встретили \| (экранированный пайп), не разрываем строку
        elif char == '\\' and i + 1 < len(row) and row[i+1] == '|':
            current_cell.append('\\|')
            i += 1
            
        # Если встретили |, но мы НЕ внутри формулы — это разделитель колонок
        elif char == '|' and not in_math:
            cells.append("".join(current_cell).strip())
            current_cell = []
            
        else:
            current_cell.append(char)
            
        i += 1
        
    cells.append("".join(current_cell).strip())
    return cells


def parse_alignments(sep_row: str):
    cells = split_md_row(sep_row)
    out = []
    for c in cells:
        c = c.strip()
        if c.startswith(':') and c.endswith(':'): out.append(WD_ALIGN_PARAGRAPH.CENTER)
        elif c.endswith(':'): out.append(WD_ALIGN_PARAGRAPH.RIGHT)
        else: out.append(WD_ALIGN_PARAGRAPH.LEFT)
    return out


def get_heading(line):
    line = line.strip()
    m = re.match(r"^(#{1,3})\s+(.+)$", line)
    if m:
        level = len(m.group(1))
        text = m.group(2).strip()
        if text.lower() in STRUCTURAL_HEADINGS:
            text = text.upper()
        return level, text

    if re.match(r"^(Введение|Заключение|Список литературы|Содержание)\s*$", line, re.I):
        return 1, line.upper()

    if re.match(r"^Глава\s+\d+", line, re.I):
        return 1, line

    if re.match(r"^\d+\.\d+\.\s+", line):
        return 2, line

    return None