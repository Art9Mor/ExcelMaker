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

    for spec_file in PROJECT_ROOT.glob("*.spec"):
        spec_file.unlink()
        print(f"✓ Удалён spec файл: {spec_file}")


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

    exe_names = ["ExcelMaker", "ExcelMaker.exe"]
    exe_source = None

    for name in exe_names:
        candidate = DIST_DIR / name
        if candidate.exists():
            exe_source = candidate
            break

    if not exe_source:
        print(f"❌ Исполняемый файл не найден в {DIST_DIR}!")
        if DIST_DIR.exists():
            print(f"Содержимое {DIST_DIR}: {list(DIST_DIR.iterdir())}")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    output_file = OUTPUT_DIR / exe_source.name
    shutil.copy2(exe_source, output_file)

    size_mb = output_file.stat().st_size / (1024 * 1024)
    print(f"✓ Файл скопирован: {output_file} ({size_mb:.2f} MB)")


if __name__ == "__main__":
    print("=" * 50)
    print("🔨 СБОРКА ExcelMaker")
    print("=" * 50)

    clean()
    build()
    copy_result()

    print("\n" + "=" * 50)
    print("✨ Сборка завершена!")

    output_dir = OUTPUT_DIR
    if output_dir.exists():
        files = list(output_dir.iterdir())
        if files:
            print(f"📁 Результат: {files[0]}")

    print("=" * 50)
