import argparse
import os
from formatter.document_format import format_document_stp

def main():
    parser = argparse.ArgumentParser(
        description="Форматирует отчёт по ГОСТ/СПТ для учебной лабораторной работы."
    )
    
    parser.add_argument(
        "input_file",
        nargs="?",
        default="input.docx",
        help="Путь к входному .docx файлу (по умолчанию: input.docx)"
    )
    
    parser.add_argument(
        "output_file",
        nargs="?",
        default="formatted_file.docx",
        help="Путь к выходному .docx файлу (по умолчанию: formatted_file.docx)"
    )
    
    parser.add_argument(
        "--fio", "-f",
        default="Петров А. С.",
        help="ФИО студента (по умолчанию: Петров А. С.)"
    )
    
    parser.add_argument(
        "--group", "-g",
        default="П32",
        help="Номер группы (по умолчанию: П32)"
    )
    
    parser.add_argument(
        "--lab", "-l",
        type=int,
        default=1,
        help="Номер лабораторной работы (по умолчанию: 1)"
    )

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"Файл '{args.input_file}' не найден.")
        print("Поместите файл в папку со скриптом или укажите полный путь.")
        return

    format_document_stp(
        input_path=args.input_file,
        output_path=args.output_file,
        fio=args.fio,
        group=args.group,
        lab_number=args.lab
    )

if __name__ == "__main__":
    main()