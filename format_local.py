import argparse
import sys
from pathlib import Path

from docx import Document

from docx_formatter import apply_gost_styles


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


def process_document(input_path: Path, output_path: Path, styles: str = "ALL") -> None:
    """Выполняет загрузку, применение стилей ГОСТ и сохранение документа.

    Args:
        input_path (Path): Путь к исходному файлу .docx.
        output_path (Path): Путь для сохранения отформатированного файла.
        styles (str): Подмножество стилей для применения (по умолчанию "ALL").
    """
    # Загружаем документ; python-docx может выбросить ValueError при повреждении структуры ZIP
    doc = Document(input_path)
    
    # Применяем правила ГОСТ-оформления через внешний модуль bot
    apply_gost_styles(doc, styles)
    
    # Записываем результат на диск (возможна OSError при отсутствии прав записи)
    doc.save(output_path)


def main() -> None:
    """Точка входа CLI: парсинг аргументов, валидация путей и запуск обработки."""
    parser = argparse.ArgumentParser(description="Локальное форматирование документов .docx по ГОСТ.")
    parser.add_argument("input_path", type=Path, help="Путь к исходному файлу .docx")
    parser.add_argument("--styles", type=str, default="ALL", help="Категории применяемых стилей (по умолчанию: ALL)")
    args = parser.parse_args()

    try:
        output_path = validate_paths(args.input_path)
    except (FileNotFoundError, ValueError) as err:
        print(f"Ошибка валидации: {err}", file=sys.stderr)
        sys.exit(1)

    try:
        process_document(args.input_path, output_path, args.styles)
        print(f"Файл успешно отформатирован и сохранен как:\n{output_path}")
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
