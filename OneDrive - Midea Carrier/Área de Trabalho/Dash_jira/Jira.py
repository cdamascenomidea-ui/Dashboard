"""Compatibilidade: wrapper que delega para `scripts.jira` (nova organização).

Mantive um pequeno wrapper para compatibilidade com chamadas antigas como:

    python Jira.py input.xlsx output.xlsx --locale pt

"""

from scripts.jira import main

if __name__ == "__main__":
    main()

