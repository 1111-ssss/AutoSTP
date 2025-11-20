import argparse
import os
import json
from formatter.document_format import format_document_stp

CONFIG_FILE = "config.json"
DEFAULT_FIO = "Иванов И. И."
DEFAULT_GROUP = "П52"

def load_config():
    """Загружает ФИО и группу из config.json, если файл существует"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                return config.get("fio", DEFAULT_FIO), config.get("group", DEFAULT_GROUP)
        except (json.JSONDecodeError, FileNotFoundError):
            pass
    return DEFAULT_FIO, "П32"

def save_config(fio, group):
    """Сохраняет ФИО и группу в config.json"""
    config = {"fio": fio, "group": group}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def main():
    default_fio, default_group = load_config()

    parser = argparse.ArgumentParser(
        description="Форматирует документ по СТП."
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
        default=default_fio,
        help=f"ФИО студента (по умолчанию: {default_fio})"
    )
    
    parser.add_argument(
        "--group", "-g",
        default=default_group,
        help="Номер группы (по умолчанию: {default_group})"
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
    
    save_config(args.fio, args.group)

    format_document_stp(
        input_path=args.input_file,
        output_path=args.output_file,
        fio=args.fio,
        group=args.group,
        lab_number=args.lab
    )

if __name__ == "__main__":
    main()