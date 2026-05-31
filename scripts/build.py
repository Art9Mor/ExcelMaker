import os
import sys
from pathlib import Path
import shutil
import subprocess


PROJECT_ROOT = Path(__file__).parent.parent

DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
OUTPUT_DIR = PROJECT_ROOT / "output"


def clean() -> None:
    """
    Удаление предыдущих артефактов сборки.
    """

    for path in (DIST_DIR, BUILD_DIR):
        if path.exists():
            shutil.rmtree(path)

    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)


def build() -> None:
    """
    Сборка приложения через PyInstaller.
    """

    main_py = PROJECT_ROOT / "main.py"
    if not main_py.exists():
        print(f"Ошибка: {main_py} не найден!")
        sys.exit(1)

    icon_path = PROJECT_ROOT / "assets" / "icon.ico"

    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name", "ExcelMaker",
        "--add-data", f"src{os.pathsep}src",
        "--hidden-import", "openpyxl",
        "--hidden-import", "loguru",
        "--hidden-import", "typer",
        "--hidden-import", "PyQt6",
        "--hidden-import", "PyQt6.QtCore",
        "--hidden-import", "PyQt6.QtGui",
        "--hidden-import", "PyQt6.QtWidgets",
        "--collect-all", "openpyxl",
        "--collect-all", "loguru",
        str(main_py),
    ]

    if icon_path.exists():
        cmd.extend(["--icon", str(icon_path)])
        print(f"✓ Иконка найдена: {icon_path}")
    else:
        print("⚠ Иконка не найдена, сборка без иконки")

    cmd.append(str(main_py))

    print(f"\n🚀 Запуск сборки...")
    print(f"Команда: {' '.join(cmd)}\n")

    result = subprocess.run(cmd, cwd=str(PROJECT_ROOT))

    if result.returncode != 0:
        print(f"\n❌ Ошибка сборки! Код: {result.returncode}")
        sys.exit(result.returncode)

    print("\n✅ Сборка завершена успешно!")


def copy_result() -> None:
    """
    Копирование результата в каталог output.
    """

    exe_source = DIST_DIR / "ExcelMaker.exe"

    if not exe_source.exists():
        print(f"❌ Файл {exe_source} не найден!")
        if DIST_DIR.exists():
            print(f"Содержимое {DIST_DIR}: {list(DIST_DIR.iterdir())}")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    shutil.copy2(
        DIST_DIR / "ExcelMaker.exe",
        OUTPUT_DIR / "ExcelMaker.exe",
    )
    print(f"✓ Файл скопирован: {OUTPUT_DIR / 'ExcelMaker.exe'}")


if __name__ == "__main__":
    print("=" * 50)
    print("🔨 СБОРКА ExcelMaker")
    print("=" * 50)

    clean()
    build()
    copy_result()

    print("\n" + "=" * 50)
    print("✨ Сборка завершена!")
    print(f"📁 Результат: {OUTPUT_DIR / 'ExcelMaker.exe'}")
    print("=" * 50)
