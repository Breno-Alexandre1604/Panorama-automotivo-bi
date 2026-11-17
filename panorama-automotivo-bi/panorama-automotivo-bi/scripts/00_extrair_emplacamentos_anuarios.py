"""Extrai os emplacamentos mensais de automóveis dos Anuários FENABRAVE (PDF).

Uso: salve os PDFs em data/raw/anuarios_fenabrave/ e rode `python scripts/00_extrair_emplacamentos_anuarios.py`.
Cada gráfico mensal traz 3 anos por mês; os valores ficam na coluna esquerda (x < 100), em ordem vertical.
Páginas: Anuário 2023 p.12 (2021-2023), Anuário 2024 p.12 (2022-2024), Anuário 2025 p.14 (2023-2025).
"""
import json, pathlib, pdfplumber
BASE = pathlib.Path(__file__).resolve().parent.parent
PDFS = BASE / "data" / "raw" / "anuarios_fenabrave"
FONTES = [("Anuario2023.pdf", 12, [2021, 2022, 2023]),
          ("Anuário2024.pdf", 12, [2022, 2023, 2024]),
          ("Anuario2025.pdf", 14, [2023, 2024, 2025])]

def mensal(arquivo, pagina, anos):
    with pdfplumber.open(PDFS / arquivo) as pdf:
        p = pdf.pages[pagina - 1]
        w = [x for x in p.extract_words()
             if "." in x["text"] and x["text"].replace(".", "").isdigit() and x["x0"] < 100 and len(x["text"]) >= 6]
    v = [int(x["text"].replace(".", "")) for x in sorted(w, key=lambda x: x["top"])]
    return {a: [v[m * 3 + i] for m in range(12)] for i, a in enumerate(anos)}

if __name__ == "__main__":
    saida = {n: mensal(n, pg, anos) for n, pg, anos in FONTES}
    for n, d in saida.items():
        for ano, lista in d.items():
            print(n, ano, sum(lista))
    (BASE / "data" / "interim" / "emplacamentos_extraidos.json").write_text(json.dumps(saida, ensure_ascii=False, indent=1))
