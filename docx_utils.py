from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from config import FONT, FONT_SIZE, LINE_SPACING, FIRST_LINE_INDENT


def set_run_font(run, size=FONT_SIZE, bold=None, italic=None):
    run.font.name = FONT
    run.font.size = size
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    
    rpr = run._element.get_or_add_rPr()
    
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert_element_before(rfonts, 'w:b', 'w:i', 'w:color', 'w:sz', 'w:lang')
    rfonts.set(qn("w:eastAsia"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)
    rfonts.set(qn("w:cs"), FONT)

    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.insert_element_before(lang, 'w:eastAsianLayout', 'w:specVanish', 'w:oMath')
    lang.set(qn("w:val"), "ru-RU")
    lang.set(qn("w:eastAsia"), "ru-RU")
    lang.set(qn("w:bidi"), "ar-SA")


def set_style_language(style, lang_code="ru-RU"):
    rpr = style.element.get_or_add_rPr()
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.insert_element_before(lang, 'w:eastAsianLayout', 'w:specVanish', 'w:oMath')
    lang.set(qn("w:val"), lang_code)
    lang.set(qn("w:eastAsia"), lang_code)
    lang.set(qn("w:bidi"), "ar-SA")


def set_default_language(doc, lang_code="ru-RU"):
    styles_el = doc.styles.element
    doc_defaults = styles_el.find(qn("w:docDefaults"))
    if doc_defaults is None:
        doc_defaults = OxmlElement("w:docDefaults")
        styles_el.insert(0, doc_defaults)
    rpr_default = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = OxmlElement("w:rPrDefault")
        doc_defaults.append(rpr_default)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rpr_default.append(rpr)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.insert_element_before(lang, 'w:eastAsianLayout', 'w:specVanish', 'w:oMath')
    lang.set(qn("w:val"), lang_code)
    lang.set(qn("w:eastAsia"), lang_code)
    lang.set(qn("w:bidi"), "ar-SA")


def force_russian_language_on_all_runs(doc):
    def fix_run(run):
        rpr = run._element.get_or_add_rPr()
        lang = rpr.find(qn("w:lang"))
        if lang is None:
            lang = OxmlElement("w:lang")
            rpr.insert_element_before(lang, 'w:eastAsianLayout', 'w:specVanish', 'w:oMath')
        lang.set(qn("w:val"), "ru-RU")
        lang.set(qn("w:eastAsia"), "ru-RU")
        lang.set(qn("w:bidi"), "ar-SA")

    def fix_container(container):
        for p in container.paragraphs:
            for run in p.runs: fix_run(run)
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells: fix_container(cell)

    fix_container(doc)
    for section in doc.sections:
        for hf in (section.header, section.first_page_header, section.even_page_header, section.footer, section.first_page_footer, section.even_page_footer):
            try: fix_container(hf)
            except Exception: pass


def ensure_heading_styles(doc):
    for level in (1, 2, 3):
        name = f"Heading {level}"
        try:
            doc.styles[name]
            continue
        except KeyError:
            pass
        try:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
        except Exception:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        try: style.base_style = doc.styles["Normal"]
        except KeyError: pass
        ppr = style.element.get_or_add_pPr()
        if ppr.find(qn("w:outlineLvl")) is None:
            outline = OxmlElement("w:outlineLvl")
            outline.set(qn("w:val"), str(level - 1))
            ppr.append(outline)


def configure_heading_styles(doc):
    for name in ["Heading 1", "Heading 2", "Heading 3"]:
        st = doc.styles[name]
        st.font.name = FONT
        st.font.size = FONT_SIZE
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.line_spacing = LINE_SPACING
        st.paragraph_format.space_after = Pt(12)
        st.paragraph_format.first_line_indent = Cm(0)
        
        # ЖЕСТКО ПРОПИСЫВАЕМ ШРИФТ В XML, чтобы Word Online не сбрасывал его
        rpr = st.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), FONT)
        rfonts.set(qn("w:hAnsi"), FONT)
        rfonts.set(qn("w:eastAsia"), FONT)
        rfonts.set(qn("w:cs"), FONT)
        
        set_style_language(st)


def configure_toc_styles(doc):
    for name in ["TOC 1", "TOC 2", "TOC 3", "TOC Heading"]:
        try: st = doc.styles[name]
        except KeyError: continue
        st.font.name = FONT
        st.font.size = FONT_SIZE
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.line_spacing = LINE_SPACING
        st.paragraph_format.space_after = Pt(0)
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.first_line_indent = Cm(0)
        
        # ЖЕСТКИЙ ФИКС ШРИФТА ДЛЯ TOC
        rpr = st.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), FONT)
        rfonts.set(qn("w:hAnsi"), FONT)
        rfonts.set(qn("w:eastAsia"), FONT)
        rfonts.set(qn("w:cs"), FONT)
        
        set_style_language(st)


def set_paragraph_format(p, first_indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    pf = p.paragraph_format
    pf.line_spacing = LINE_SPACING
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.first_line_indent = FIRST_LINE_INDENT if first_indent else Cm(0)
    pf.alignment = align


def add_page_number_footer(section):
    footer = section.footer
    for p in list(footer.paragraphs): p.text = ""
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    run1 = p.add_run()
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    run1._r.append(fld1)
    
    run2 = p.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    run2._r.append(instr)
    
    run3 = p.add_run()
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run3._r.append(fld2)
    
    for r in (run1, run2, run3): set_run_font(r)


def add_toc(paragraph):
    run1 = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    run1._r.append(fld1)

    run2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-2" \\h \\z \\u'
    run2._r.append(instr)

    run3 = paragraph.add_run()
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "separate")
    run3._r.append(fld2)

    run4 = paragraph.add_run("Обновите поле: выделите и нажмите F9")

    run5 = paragraph.add_run()
    fld3 = OxmlElement("w:fldChar")
    fld3.set(qn("w:fldCharType"), "end")
    run5._r.append(fld3)
    
    for r in (run1, run2, run3, run4, run5): set_run_font(r)


def set_update_fields(doc):
    settings = doc.settings.element
    # Удаляем, если уже есть, чтобы не было дублей
    for el in settings.findall(qn('w:updateFields')):
        settings.remove(el)
        
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.insert_element_before(update_fields, 
        'w:compat', 'w:rsids', 'm:mathPr', 'w:themeFontLang', 'w:clrSchemeMapping', 
        'w:doNotIncludeSubdocsInStats', 'w:doNotAutoCompressPictures', 'w:forceUpgrade', 
        'w:captions', 'w:readModeInkLockDown', 'w:smartTagType', 'w:shapeDefaults', 
        'w:doNotEmbedSmartTags', 'w:decimalSymbol', 'w:listSeparator'
    )


def add_toc_page(doc):
    p = doc.add_paragraph()
    set_paragraph_format(p, first_indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.page_break_before = True
    run = p.add_run("СОДЕРЖАНИЕ")
    set_run_font(run, bold=True)
    p.paragraph_format.space_after = Pt(12)

    p2 = doc.add_paragraph()
    set_paragraph_format(p2, first_indent=False, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_toc(p2)