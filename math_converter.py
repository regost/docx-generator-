import os
import re
import shutil
import subprocess
import tempfile
import zipfile

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

PANDOC = shutil.which("pandoc")
SKIP_PANDOC = bool(os.environ.get("NO_PANDOC"))

_LATEX_TO_MATHML = None
try:
    from latex2mathml.converter import convert as _latex_to_mathml_fn
    _LATEX_TO_MATHML = _latex_to_mathml_fn
except ImportError:
    pass

_MATHML_TO_OMML = None
try:
    from mathml2omml import convert as _mathml_to_omml_fn
    _MATHML_TO_OMML = _mathml_to_omml_fn
except ImportError:
    pass

if not SKIP_PANDOC and PANDOC is None:
    print("ВНИМАНИЕ: pandoc не найден в PATH.")

_MATH_STATS = {"pandoc": 0, "m2o": 0, "fail": 0}
_PANDOC_CACHE = {}


def _clean_latex(s: str) -> str:
    s = s.strip()
    s = re.sub(r'\\text\s*\{\s*([A-Za-z]+)\s*\}', lambda m: r'\mathrm{' + m.group(1) + '}', s)
    s = re.sub(r'\\text\s*\{\s*([^}]*[А-Яа-я][^}]*)\s*\}', lambda m: r'\mathrm{' + re.sub(r'\s+', '', m.group(1)) + '}', s)
    s = re.sub(r'\\[,;!:]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _extract_omath_fragment(xml: str):
    m = re.search(r'<m:oMath\b[^>]*>.*?</m:oMath>', xml, re.DOTALL)
    if not m:
        return None
    frag = m.group(0)
    if 'xmlns:m' not in frag:
        frag = frag.replace('<m:oMath', f'<m:oMath xmlns:m="{OMML_NS}"', 1)
    return frag


def _patch_omml(omml):
    """Добавляет Cambria Math в <m:r> в строгом соответствии со схемой."""
    for m_r in omml.iter(qn('m:r')):
        w_rPr = m_r.find(qn('w:rPr'))
        if w_rPr is None:
            w_rPr = OxmlElement('w:rPr')
            # СТРОГИЙ ПОРЯДОК: w:rPr должен идти ПОСЛЕ m:rPr, но ПЕРЕД m:t
            m_t = m_r.find(qn('m:t'))
            if m_t is not None:
                m_t.addprevious(w_rPr)
            else:
                m_r.append(w_rPr)
        
        w_rFonts = w_rPr.find(qn('w:rFonts'))
        if w_rFonts is None:
            w_rFonts = OxmlElement('w:rFonts')
            w_rPr.insert(0, w_rFonts)
            
        w_rFonts.set(qn('w:ascii'), 'Cambria Math')
        w_rFonts.set(qn('w:hAnsi'), 'Cambria Math')
        w_rFonts.set(qn('w:cs'), 'Cambria Math')
    return omml


# ==========================================
# НАСТРОЙКИ ПРЕПРОЦЕССИНГА
# ==========================================

# Список операторов, которые нельзя оставлять "голыми" в начале выражения
OPERATORS = ['<', '>', '=', '+', '-', '*', '~', '<=', '>=', '~=', '\\le', '\\ge', '\\approx', '\\neq']

# Невидимый символ нулевой ширины (Zero-Width Space) в текстовом блоке.
# Он нужен, чтобы LibreOffice видел левый/правый операнд там, где его нет.
INVISIBLE_OPERAND = '\\text{\u200b}'


def _fix_leading_operators(latex_str: str) -> str:
    """Добавляет невидимый операнд в начало формулы или степени, если там стоит оператор."""
    # Сортируем по длине, чтобы <= срабатывало раньше <
    ops_sorted = sorted(OPERATORS, key=len, reverse=True)
    ops_regex = "|".join(re.escape(op) for op in ops_sorted)
    
    # 1. В начале всей формулы: <10 -> \text{\u200b} <10
    latex_stripped = latex_str.lstrip()
    if re.match(r'^(' + ops_regex + r')', latex_stripped):
        latex_str = INVISIBLE_OPERAND + ' ' + latex_stripped
        
    # 2. В начале степени: x^<10 -> x^{\text{\u200b}<10}
    # Сначала оборачиваем в скобки, если их нет: x^< -> x^{<}
    pattern_unwrapped = r'(\^)\s*(' + ops_regex + r')'
    latex_str = re.sub(pattern_unwrapped, r'\1{\2}', latex_str)
    
    # Затем добавляем невидимый операнд внутри скобок: x^{<10} -> x^{\text{\u200b}<10}
    pattern_wrapped = r'(\^)\{\s*(' + ops_regex + r')'
    latex_str = re.sub(pattern_wrapped, r'\1{' + INVISIBLE_OPERAND + r'\2', latex_str)
    
    return latex_str


def _fix_trailing_operators(latex_str: str) -> str:
    """Добавляет невидимый операнд в конец формулы, если там стоит оператор."""
    latex_stripped = latex_str.rstrip()
    
    # 1. Если формула заканчивается на оператор
    if latex_stripped.endswith(('+', '-', '=')):
        latex_str = latex_stripped + ' ' + INVISIBLE_OPERAND
        
    # 2. Если формула заканчивается на \quad (перенос строки)
    elif latex_stripped.endswith(r'\quad'):
        base = latex_stripped[:-5].rstrip()
        # Проверяем, есть ли оператор перед \quad
        if base.endswith(('+', '-', '=')):
            latex_str = base + ' ' + INVISIBLE_OPERAND + ' ' + r'\quad'
            
    return latex_str


def _preprocess_latex(latex_str: str) -> str:
    """Главная функция, которая применяет все фиксы к строке LaTeX."""
    # 1. Фиксы выравнивания (LibreOffice не понимает aligned/array)
    latex_str = latex_str.replace(r'\begin{aligned}', '').replace(r'\end{aligned}', '')
    latex_str = latex_str.replace(r'\begin{array}{l}', '').replace(r'\end{array}', '')
    latex_str = latex_str.replace('&', '') 
    latex_str = latex_str.replace(r'\\', r'\quad') 
    
    # 2. Юникод-степени (⁻¹⁰ -> ^{-10})
    superscripts = {
        '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
        '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9',
        '⁻': '-', '⁺': '+'
    }
    def replace_unicode_sup(match):
        chars = match.group(1)
        converted = ''.join(superscripts.get(c, c) for c in chars)
        return f'^{{{converted}}}'
    latex_str = re.sub(r'([⁰¹²³⁴⁵⁶⁷⁸⁹⁻⁺]+)', replace_unicode_sup, latex_str)
    
    # 3. Оборачивание многозначных чисел в степени (256^2 -> {256}^{2})
    latex_str = re.sub(r'(?<![\w\{])(\d+)\^(\{[^\}]+\}|[^\s^])', r'{\1}^{\2}', latex_str)
    
    # 4. Звездочка в степени (^* -> ^{\text{*}})
    latex_str = re.sub(r'\^\s*\{\s*\*\s*\}', r'^{\\text{*}}', latex_str)
    latex_str = re.sub(r'\^\s*\*', r'^{\\text{*}}', latex_str)
    
    # 5. Фиксы операторов (вызов вспомогательных функций)
    latex_str = _fix_leading_operators(latex_str)
    latex_str = _fix_trailing_operators(latex_str)
            
    return latex_str


def latex_to_omml_via_pandoc(latex_str):
    if SKIP_PANDOC or not PANDOC:
        return None

    if latex_str in _PANDOC_CACHE:
        cached = _PANDOC_CACHE[latex_str]
        if cached is None: return None
        try: return _patch_omml(parse_xml(cached))
        except Exception: return None

    latex_fixed = _preprocess_latex(latex_str)

    md_path = docx_path = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8") as f:
            f.write(f"${latex_fixed}$")
            md_path = f.name
        docx_path = md_path[:-3] + ".docx"

        subprocess.run([PANDOC, md_path, "-o", docx_path], check=True, capture_output=True, timeout=30)

        with zipfile.ZipFile(docx_path) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")

        frag = _extract_omath_fragment(doc_xml)
        if frag is None:
            _PANDOC_CACHE[latex_str] = None
            return None

        _PANDOC_CACHE[latex_str] = frag
        return _patch_omml(parse_xml(frag))
        
    except Exception as e:
        _PANDOC_CACHE[latex_str] = None
        return None
    finally:
        for p in (md_path, docx_path):
            if p and os.path.exists(p):
                try: os.unlink(p)
                except OSError: pass


def latex_to_omml_via_mathml2omml(latex_str):
    if not (_LATEX_TO_MATHML and _MATHML_TO_OMML):
        return None
    try:
        mathml = _LATEX_TO_MATHML(_clean_latex(latex_str))
        omml = _MATHML_TO_OMML(mathml)
        if 'xmlns:m' not in omml:
            if omml.startswith('<m:oMath'):
                omml = omml.replace('<m:oMath', f'<m:oMath xmlns:m="{OMML_NS}"', 1)
            elif omml.startswith('<oMath'):
                omml = omml.replace('<oMath', f'<m:oMath xmlns:m="{OMML_NS}"', 1).replace('</oMath>', '</m:oMath>')
            else:
                omml = f'<m:oMath xmlns:m="{OMML_NS}">{omml}</m:oMath>'
        return _patch_omml(parse_xml(omml))
    except Exception as e:
        return None


def latex_to_omml_element(latex_str, verbose=False):
    el = latex_to_omml_via_pandoc(latex_str)
    if el is not None:
        _MATH_STATS["pandoc"] += 1
        return el

    el = latex_to_omml_via_mathml2omml(latex_str)
    if el is not None:
        _MATH_STATS["m2o"] += 1
        return el

    _MATH_STATS["fail"] += 1
    return None