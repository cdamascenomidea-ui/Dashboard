from pathlib import Path
import argparse
import sys

from dashboard import build_dashboard


def main(argv=None):
    parser = argparse.ArgumentParser(description="Gera dashboard a partir de um arquivo de chamados Jira")
    parser.add_argument("input", type=Path, help="Caminho para o arquivo Excel de entrada")
    parser.add_argument("output", type=Path, nargs="?", default=Path("Dashboard_Chamados_Jira.xlsx"), help="Arquivo Excel de saída")
    parser.add_argument("--locale", choices=["en", "pt"], default="en", help="Locale para fórmulas do Excel (en/pt)")
    args = parser.parse_args(argv)

    try:
        out = build_dashboard(args.input, args.output, args.locale)
        print(f"Dashboard salvo em: {out}")
    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
