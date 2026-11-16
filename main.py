import os
import sys
import argparse
from docx import Document

import config
import math_converter
import docx_utils as utils
from document_builder import add_body
from markdown_parser import strip_front_matter, collapse_display_math, normalize_latex_delimiters


def process_project_folder(project_dir):
    """Сканирует папку проекта и настраивает пути в config."""
    if not os.path.isdir(project_dir):
        print(f"ОШИБКА: Папка '{project_dir}' не найдена.")
        sys.exit(1)

    print(f"Сканирую папку проекта: {project_dir}...")
    files = os.listdir(project_dir)

    # 1. Ищем файл .md или .txt (текст работы)
    md_files = [f for f in files if f.endswith('.md')]
    txt_files = [f for f in files if f.endswith('.txt') and f != 'requirements.txt']
    
    if len(md_files) == 1:
        config.INPUT_MD = os.path.join(project_dir, md_files[0])
    elif len(md_files) > 1:
        print("ОШИБКА: В папке найдено несколько .md файлов. Оставьте только один.")
        sys.exit(1)
    elif len(txt_files) == 1:
        config.INPUT_MD = os.path.join(project_dir, txt_files[0])
    elif len(txt_files) > 1:
        print("ОШИБКА: В папке найдено несколько .txt файлов. Оставьте только один.")
        sys.exit(1)
    else:
        print("ОШИБКА: В папке не найдено ни одного .md или .txt файла.")
        sys.exit(1)

    # 2. Ищем титульный лист (.docx)
    output_filename = os.path.basename(config.OUTPUT_DOCX) if config.OUTPUT_DOCX else "referat.docx"
    all_docx = [f for f in files if f.endswith('.docx') and f != output_filename]
    
    # Ищем те, что начинаются с _
    underscore_docx = [f for f in all_docx if f.startswith('_')]
    
    if len(underscore_docx) == 1:
        config.TEMPLATE_DOCX = os.path.join(project_dir, underscore_docx[0])
        print(f"  Найден титульный лист: {underscore_docx[0]}")
    elif len(underscore_docx) > 1:
        print("\nНайдено несколько файлов-титульников (начинаются с '_'):")
        for idx, f in enumerate(underscore_docx):
            print(f"  {idx + 1}. {f}")
        try:
            choice = int(input("Выберите номер шаблона (по умолчанию 1): ") or "1")
            config.TEMPLATE_DOCX = os.path.join(project_dir, underscore_docx[choice - 1])
        except (ValueError, IndexError):
            config.TEMPLATE_DOCX = os.path.join(project_dir, underscore_docx[0])
    elif len(all_docx) == 1:
        # Если титульников с _ нет, но есть один обычный docx — берем его
        config.TEMPLATE_DOCX = os.path.join(project_dir, all_docx[0])
    elif len(all_docx) > 1:
        # Если несколько обычных docx, просим выбрать
        print("\nНайдено несколько файлов .docx:")
        for idx, f in enumerate(all_docx):
            print(f"  {idx + 1}. {f}")
        try:
            choice = int(input("Выберите номер шаблона (по умолчанию 1): ") or "1")
            config.TEMPLATE_DOCX = os.path.join(project_dir, all_docx[choice - 1])
        except (ValueError, IndexError):
            config.TEMPLATE_DOCX = os.path.join(project_dir, all_docx[0])
    else:
        print("ОШИБКА: В папке не найден файл титульного листа (.docx).")
        sys.exit(1)

    # 3. Настраиваем папку с картинками и выходной файл
    config.IMAGES_DIR = os.path.join(project_dir, "images")
    config.OUTPUT_DOCX = os.path.join(project_dir, "output.docx")

    print(f"  - Входной текст:   {os.path.basename(config.INPUT_MD)}")
    print(f"  - Папка картинок:  {config.IMAGES_DIR}")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Генератор реферата в .docx по ГОСТ.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("-p", "--project", default=None,
                        help="Имя папки с проектом (где лежат .md, _Титульник.docx и images)")
    parser.add_argument("-m", "--md", default=None, 
                        help="Путь к конкретному .md/.txt файлу (перекрывает -p)")
    parser.add_argument("-o", "--output", default=None, 
                        help="Имя выходного .docx файла (перекрывает -p)")
    parser.add_argument("-i", "--images", default=None, 
                        help="Путь к папке с картинками (перекрывает -p)")
    parser.add_argument("-t", "--template", default=None, 
                        help="Путь к титульному листу (перекрывает -p)")
    return parser.parse_args()


def main():
    args = parse_args()
    
    # Если передан флаг проекта, сканируем папку
    if args.project:
        process_project_folder(args.project)
    else:
        # Старое поведение: ищем файл в текущей папке
        config.INPUT_MD = config.INPUT_MD
        config.IMAGES_DIR = config.IMAGES_DIR
        config.TEMPLATE_DOCX = config.TEMPLATE_DOCX
        config.OUTPUT_DOCX = config.OUTPUT_DOCX

    # Если переданы конкретные флаги, они имеют наивысший приоритет
    if args.md: config.INPUT_MD = args.md
    if args.images: config.IMAGES_DIR = args.images
    if args.template: config.TEMPLATE_DOCX = args.template
    if args.output: config.OUTPUT_DOCX = args.output

    if not os.path.exists(config.INPUT_MD):
        print(f"ОШИБКА: не найден файл '{config.INPUT_MD}'.")
        sys.exit(1)
    if not os.path.exists(config.TEMPLATE_DOCX):
        print(f"ОШИБКА: не найден '{config.TEMPLATE_DOCX}'.")
        sys.exit(1)

    print("=" * 60)
    print(f"pandoc:            {'найден' if math_converter.PANDOC else 'НЕ найден'}")
    print(f"Шаблон:            {config.TEMPLATE_DOCX}")
    print(f"Вход:              {config.INPUT_MD}")
    print(f"Выход:             {config.OUTPUT_DOCX}")
    print("=" * 60)

    with open(config.INPUT_MD, "r", encoding="utf-8") as f:
        text = f.read()

    body = strip_front_matter(text)
    body = normalize_latex_delimiters(body)
    body = collapse_display_math(body)

    doc = Document(config.TEMPLATE_DOCX)

    for section in doc.sections:
        section.different_first_page_header_footer = True
        utils.add_page_number_footer(section)

    utils.set_default_language(doc)
    utils.ensure_heading_styles(doc)
    utils.configure_heading_styles(doc)
    utils.configure_toc_styles(doc)

    utils.add_toc_page(doc)
    add_body(doc, body)

    utils.force_russian_language_on_all_runs(doc)
    utils.set_update_fields(doc)
    doc.save(config.OUTPUT_DOCX)

    print("=" * 60)
    print(f"Формул сконвертировано: {math_converter._MATH_STATS['pandoc']}")
    print(f"Формул НЕ сконвертировано: {math_converter._MATH_STATS['fail']}")
    print("=" * 60)
    print(f"Готово: {config.OUTPUT_DOCX}")


if __name__ == "__main__":
    main()