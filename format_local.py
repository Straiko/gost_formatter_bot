import argparse
import sys
from pathlib import Path

from docx import Document

from bot import apply_gost_styles


def validate_paths(input_path: Path) -> Path:
    """Проверяет существование входного файла и формирует путь для выходного документа.

    Args:
        input_path (Path): Путь к исходному файлу .docx.

    Returns:
        Path: Сгенерированный путь для сохранения отформатированного файла.

    Raises:
        FileNotFoundError: Если файл не найден или не является регулярным файлом.
        ValueError: Если расширение файла отличается от .docx.
    """
    # Проверяем физическое существование файла на диске
    if not input_path.exists() or not input_path.is_file():
        raise FileNotFoundError(f"Файл '{input_path}' не найден.")

    # Проверяем корректность формата (только .docx)
    if input_path.suffix.lower() != ".docx":
        raise ValueError(f"Файл '{input_path}' не является документом .docx.")

    return input_path.with_name(f"{input_path.stem}_ГОСТ{input_path.suffix}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Локальное форматирование документов .docx по ГОСТ.")
    parser.add_argument("input_path", type=Path, help="Путь к исходному файлу .docx")
    args = parser.parse_args()

    try:
        output_path = validate_paths(args.input_path)
    except (FileNotFoundError, ValueError) as err:
        print(f"Ошибка валидации: {err}", file=sys.stderr)
        sys.exit(1)

    try:
        doc = Document(input_path)
        apply_gost_styles(doc, "ALL")
        doc.save(output_path)
        print(f"Файл успешно отформатирован и сохранен как:\n{output_path}")
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
