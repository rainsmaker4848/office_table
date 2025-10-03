import subprocess
import sys
import os

# === Список необходимых внешних библиотек ===
required_packages = [
    "pandas",
    "openpyxl",
    "python-docx"
]

def install_packages():
    """Устанавливает все недостающие библиотеки"""
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Устанавливается: {package}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        else:
            print(f"Уже установлено: {package}")

def main():
    install_packages()
    print("✅ Все зависимости установлены.")

if __name__ == "__main__":
    main()
