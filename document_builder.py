import re
import os
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.image.image import Image

import config
from config import (
    LINE_SPACING, FIRST_LINE_INDENT, STRICT_GOST_LEVEL1, 
    AUTO_NUMBER_FORMULAS, TEXT_WIDTH_CM, STRUCTURAL_HEADINGS
)
from math_converter import latex_to_omml_element
import docx_utils as utils
from markdown_parser import (
    INLINE_MATH_RE, DISPLAY_MATH_RE, IMAGE_RE, get_heading, 
    is_table_separator, split_md_row, parse_alignments
)

def _set_table_borders(table):
    tblPr = table._tbl.tblPr
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
        
    borders = OxmlElement('w:tblBorders')
    for name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        borders.append(b)
    
    tblPr.insert_element_before(borders, 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook')

def _make_table_autofit(table):
    table.autofit = True
    tblPr = table._tbl.tblPr
    
    existing_margins = tblPr.find(qn('w:tblCellMar'))
    if existing_margins is not None:
        tblPr.remove(existing_margins)
        
    tblCellMar = OxmlElement('w:tblCellMar')
    for side in ('top', 'left', 'bottom', 'right'):
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), '0')
        m.set(qn('w:type'), 'dxa')
        tblCellMar.append(m)
    
    tblPr.insert_element_before(tblCellMar, 'w:tblLook', 'w:tblCaption', 'w:tblDescription')


def add_inline_math(p, latex_str, size=None):
    omml = latex_to_omml_element(latex_str)
    if omml is not None:
        p._p.append(omml)
    else:
        run = p.add_run(latex_str.replace(r'\,', ' '))
        utils.set_run_font(run, size=size, italic=True) if size else utils.set_run_font(run, italic=True)


def _add_text_with_math(p, text, bold=False):
    pos = 0
    for m in INLINE_MATH_RE.finditer(text):
        if m.start() > pos:
            run = p.add_run(text[pos:m.start()])
            utils.set_run_font(run, bold=bold)
        add_inline_math(p, m.group(1))
        pos = m.end()
    if pos < len(text):
        run = p.add_run(text[pos:])
        utils.set_run_font(run, bold=bold)

def _add_paragraph_with_inline_math(doc, text):
    p = doc.add_paragraph()
    utils.set_paragraph_format(p, first_indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    _add_text_with_math(p, text, bold=False)


def add_display_formula(doc, latex_str, number=None):
    omml = latex_to_omml_element(latex_str)

    if omml is None:
        p = doc.add_paragraph()
        utils.set_paragraph_format(p, first_indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        run = p.add_run(latex_str.replace(r'\,', ' '))
        utils.set_run_font(run, italic=True)
        if number is not None:
            run2 = p.add_run(f"    ({number})")
            utils.set_run_font(run2)
        return

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = LINE_SPACING
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    if number is None:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p._p.append(omml)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_WIDTH_CM / 2), WD_TAB_ALIGNMENT.CENTER)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT)

        run_tab1 = p.add_run()
        run_tab1.add_tab()
        p._p.append(omml)
        run_tab2 = p.add_run()
        run_tab2.add_tab()
        run_num = p.add_run(f"({number})")
        utils.set_run_font(run_num)


def _fill_cell(cell, text, align, bold=False):
    for r in list(cell.paragraphs[0].runs): r._r.getparent().remove(r._r)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.right_indent = Cm(0)

    text = text.replace('<br>', '\n').replace('<br/>', '\n')
    _add_text_with_math(p, text, bold=bold)
    if not text:
        utils.set_run_font(p.add_run(""))


def add_markdown_table(doc, table_lines, caption=None):
    header = split_md_row(table_lines[0])
    alignments = parse_alignments(table_lines[1])
    while len(alignments) < len(header):
        alignments.append(WD_ALIGN_PARAGRAPH.LEFT)

    rows = [split_md_row(r) for r in table_lines[2:]]

    if caption:
        p = doc.add_paragraph()
        utils.set_paragraph_format(p, first_indent=False, align=WD_ALIGN_PARAGRAPH.LEFT)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(4)
        _add_text_with_math(p, caption)

    table = doc.add_table(rows=1, cols=len(header))
    _set_table_borders(table)
    _make_table_autofit(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    trPr = table.rows[0]._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), 'true')
    trPr.insert_element_before(tblHeader, 'w:tblCellSpacing', 'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange')

    hdr_cells = table.rows[0].cells
    for i, text in enumerate(header):
        cell = hdr_cells[i]
        _fill_cell(cell, text, WD_ALIGN_PARAGRAPH.CENTER, bold=True)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for row_data in rows:
        row_cells = table.add_row().cells
        trPr = table.rows[-1]._tr.get_or_add_trPr()
        cantSplit = OxmlElement('w:cantSplit')
        trPr.insert_element_before(cantSplit, 'w:trHeight', 'w:tblHeader', 'w:tblCellSpacing', 'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange')

        for i, text in enumerate(row_data):
            if i >= len(row_cells): break
            cell = row_cells[i]
            align = alignments[i] if i < len(alignments) else WD_ALIGN_PARAGRAPH.LEFT
            _fill_cell(cell, text, align, bold=False)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    doc.add_paragraph()


def add_image(doc, caption, rel_path):
    base_dir = os.path.dirname(os.path.abspath(config.INPUT_MD))
    abs_path = os.path.join(base_dir, rel_path)
    
    if not os.path.exists(abs_path) and not os.path.isabs(rel_path):
        abs_path = os.path.join(base_dir, config.IMAGES_DIR, rel_path)

    if not os.path.exists(abs_path):
        print(f"  [img error] Не найден файл: {abs_path}")
        p = doc.add_paragraph()
        utils.set_paragraph_format(p, first_indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        run = p.add_run(f"[Отсутствует рисунок: {caption} ({rel_path})]")
        utils.set_run_font(run, italic=True)
        return

    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.first_line_indent = Cm(0)
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(4)
    run = p_img.add_run()
    
    try:
        img = Image.from_file(abs_path)
        width_cm = img.width.cm
        if width_cm > TEXT_WIDTH_CM:
            run.add_picture(abs_path, width=Cm(TEXT_WIDTH_CM))
        else:
            run.add_picture(abs_path)
    except Exception as e:
        print(f"  [img error] Ошибка вставки {abs_path}: {e}")
        return

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.first_line_indent = Cm(0)
    p_cap.paragraph_format.line_spacing = LINE_SPACING
    p_cap.paragraph_format.space_before = Pt(0)
    p_cap.paragraph_format.space_after = Pt(12)
    _add_text_with_math(p_cap, caption)



def add_code_block(doc, code_text):
    """Вставляет блок кода в рамке с серым фоном и моноширинным шрифтом."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 1. Настраиваем рамку (светло-серая)
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), 'D3D3D3') # Светло-серый цвет рамки
        borders.append(el)
    tblPr.append(borders)

    # 2. Настраиваем серый фон (заливку)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F5F5F5') # Очень светлый серый
    tblPr.append(shd)

    # 3. Настраиваем текст внутри ячейки
    cell = table.cell(0, 0)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.0 # Одинарный интервал
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Cm(0) # Без красной строки

    # 4. Добавляем сам код
    # Заменяем табуляции на пробелы для корректного отображения
    code_text = code_text.expandtabs(4)
    
    run = p.add_run(code_text)
    # ГОСТ для кода: Courier New, 12 кегль
    run.font.name = "Courier New"
    run.font.size = Pt(12)
    
    # Задаем шрифт для всех алфавитов
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert_element_before(rfonts, 'w:b', 'w:i', 'w:color', 'w:sz', 'w:lang')
    rfonts.set(qn("w:eastAsia"), "Courier New")
    rfonts.set(qn("w:hAnsi"), "Courier New")
    rfonts.set(qn("w:cs"), "Courier New")

    doc.add_paragraph() # Отступ после кода
    
    
def add_body(doc, body):
    lines = body.splitlines()
    i = 0
    in_bibliography = False
    formula_counter = 0
    img_counter = 0

    while i < len(lines):
        line = lines[i].rstrip()

        if not line.strip():
            i += 1
            continue

        # --- Блок кода ---
        if line.strip().startswith("```"):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1 # Пропускаем закрывающие ```
            
            add_code_block(doc, "\n".join(code_lines))
            continue

        # --- Картинка ---
        m_img = IMAGE_RE.match(line.strip())
        if m_img:
            caption = m_img.group(1).strip()
            rel_path = m_img.group(2).strip()
            img_counter += 1
            
            if not caption.lower().startswith("рисунок"):
                if caption: caption = f"Рисунок {img_counter} — {caption}"
                else: caption = f"Рисунок {img_counter}"
            
            add_image(doc, caption, rel_path)
            i += 1
            continue

        # --- Подпись + таблица ---
        if re.match(r'^Таблица\s+\d+', line.strip()):
            caption = line.strip()
            j = i + 1
            while j < len(lines) and not lines[j].strip(): j += 1
            if (j + 1 < len(lines) and lines[j].strip().startswith("|") and is_table_separator(lines[j + 1])):
                table_lines = [lines[j], lines[j + 1]]
                k = j + 2
                while k < len(lines) and lines[k].strip().startswith("|"):
                    table_lines.append(lines[k])
                    k += 1
                add_markdown_table(doc, table_lines, caption=caption)
                i = k
                continue

        # --- Выключная формула ---
        m_disp = DISPLAY_MATH_RE.match(line)
        if m_disp:
            latex_src = m_disp.group(1).strip()
            explicit_num = m_disp.group(2)
            
            # Проверяем, является ли формула продолжением
            is_continuation = latex_src.startswith(('+', '-', '=', '\\quad'))
            
            if explicit_num is not None: 
                number = explicit_num
            elif AUTO_NUMBER_FORMULAS and not is_continuation:
                formula_counter += 1
                number = str(formula_counter)
            else: 
                number = None
                
            add_display_formula(doc, latex_src, number=number)
            i += 1
            continue

        # --- Markdown-таблица без подписи ---
        if (line.strip().startswith("|") and i + 1 < len(lines) and is_table_separator(lines[i + 1])):
            table_lines = [line, lines[i + 1]]
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i])
                i += 1
            add_markdown_table(doc, table_lines)
            continue

        # --- Заголовок ---
        heading = get_heading(line)
        if heading:
            level, text = heading
            p = doc.add_heading(text, level=level)
            for run in p.runs: utils.set_run_font(run, bold=True)

            if level == 1:
                p.paragraph_format.page_break_before = True
                if STRICT_GOST_LEVEL1 and text.lower() not in STRUCTURAL_HEADINGS:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.paragraph_format.first_line_indent = FIRST_LINE_INDENT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.first_line_indent = Cm(0)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = FIRST_LINE_INDENT

            p.paragraph_format.line_spacing = LINE_SPACING
            p.paragraph_format.space_after = Pt(12)

            in_bibliography = level == 1 and "список литературы" in text.lower()
            i += 1
            continue

        # --- Список литературы ---
        if in_bibliography:
            p = doc.add_paragraph()
            utils.set_paragraph_format(p, first_indent=False, align=WD_ALIGN_PARAGRAPH.LEFT)
            _add_text_with_math(p, line.strip())
            i += 1
            continue

        # --- Обычный абзац ---
        para_lines = [line.strip()]
        i += 1
        while (i < len(lines) and lines[i].strip()
               and not get_heading(lines[i])
               and not DISPLAY_MATH_RE.match(lines[i])
               and not IMAGE_RE.match(lines[i].strip())
               and not re.match(r'^Таблица\s+\d+', lines[i].strip())
               and not lines[i].strip().startswith("|")
               and not lines[i].strip().startswith("```")):
            para_lines.append(lines[i].strip())
            i += 1

        text = " ".join(para_lines)
        _add_paragraph_with_inline_math(doc, text)