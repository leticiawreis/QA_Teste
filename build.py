"""
Gera o executavel (.exe) do QA Control Center.

Como usar:
    1. Abra o terminal na pasta do projeto (onde está o main.py).
    2. Rode:  python build.py
"""
import subprocess
import sys

def run(cmd):
    print(f"\n>>> {' '.join(cmd)}\n")
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print("\nAlgo deu errado no comando acima. Veja a mensagem de erro logo acima.")
        sys.exit(1)


def main():
    print("Instalando/atualizando dependências (PyQt6, openpyxl, pyinstaller)...")
    run([sys.executable, "-m", "pip", "install", "--upgrade",
         "pyinstaller", "PyQt6", "openpyxl"])

    print("\nGerando o executável...")
    run([
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--onefile", "--windowed",
        "--name", "QA Control Center",
        "main.py",
    ])

    print("\n" + "=" * 60)
    print(" Pronto! O executável está em:")
    print(r"  dist\QA Control Center.exe")
    print("=" * 60)


if __name__ == "__main__":
    main()