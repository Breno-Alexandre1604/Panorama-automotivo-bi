# Power BI

1. Abra o Power BI Desktop e importe `data/powerbi/PowerBI_Modelo_de_Dados.xlsx` (13 tabelas).
2. Crie os relacionamentos e as tabelas de parâmetro conforme `docs/guia_dashboard_powerbi.md`.
3. Cole as medidas de `medidas_dax.txt` (troque `,` por `;` nos argumentos se o seu Power BI estiver em português).

As medidas DAX foram validadas em Python, mas ainda não foram executadas no Power BI.
Salve o `.pbix` nesta pasta. O `.gitignore` o ignora por padrão; use Git LFS se quiser versioná-lo.
