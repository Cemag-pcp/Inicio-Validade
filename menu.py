from pathlib import Path
import subprocess
import sys


BASE_DIR = Path(__file__).resolve().parent


def escolher_script() -> str:
    opcoes = {
        "1": "inicio.py",
        "2": "fim.py",
        "inicio": "inicio.py",
        "fim": "fim.py",
    }

    while True:
        print("Escolha qual script deseja executar:")
        print("1 - inicio.py")
        print("2 - fim.py")
        escolha = input("Digite 1, 2, inicio ou fim: ").strip().lower()

        if escolha in opcoes:
            return opcoes[escolha]

        print("Opcao invalida. Tente novamente.\n")


def main() -> None:
    script = escolher_script()
    script_path = BASE_DIR / script
    subprocess.run([sys.executable, str(script_path)], check=True, cwd=BASE_DIR)


if __name__ == "__main__":
    main()
