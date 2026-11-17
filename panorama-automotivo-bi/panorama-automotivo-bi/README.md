# Panorama Automotivo no Brasil: BI e decisão de compra por personas

Projeto da disciplina **Análise de Dados e Business Intelligence** (UNIPÊ, Prof. Daniel Brandão).
Analisa o mercado de automóveis no Brasil (2021-2025, com dados parciais de 2026): vendas, novos × usados,
transição para eletrificados e uma ferramenta de decisão que responde: **qual é a melhor opção de carro
para a sua renda, entrada e prazo?**

**Equipe:** Breno Alexandre, Luis Eduardo, Oliveiros Neto e Polibio Alves.
**Status:** planilhas das Etapas 1 a 3 concluídas; dashboard em Power BI com modelo de dados e medidas prontos, falta montar o `.pbix`.

## Principais resultados

| Pergunta | Resultado |
| --- | --- |
| Volume de automóveis novos (FENABRAVE) | 1,56 milhão em 2021 e 2,00 milhões em 2025 (+28%) |
| Novos × usados (2025) | 16,4% novos e 83,6% usados (5,1 usados por novo) |
| Eletrificados nos emplacamentos de automóveis (2025) | 14,1% no anuário, que inclui micro-híbridos; 223.912 unidades no critério ABVE, sem eles |
| BEV × híbridos (1º sem/2026, ABVE) | BEV 42,1%, PHEV 35,5%, híbridos sem recarga 22,3% |
| Geografia (1º sem/2026) | Sudeste 44,7%; São Paulo 28,7%; 5 maiores estados 57,6% |
| Capacidade de compra (regra dos 20%, 48 meses, 2,04% a.m., 20% de entrada) | Seu Zé R$ 23 mil; Ana Paula R$ 42 mil; Lucas & Mariana R$ 114 mil; Roberto R$ 266 mil |

Os tetos de orçamento do escopo para Ana Paula (R$ 65 mil) e Lucas & Mariana (R$ 130 mil) só valem com entrada de cerca de 49% e 30%.

## Estrutura do repositório

```
data/
  raw/equipe/            planilhas recebidas, sem alteração
  raw/anuarios_fenabrave/ PDFs (NÃO versionados; veja o LEIA-ME da pasta)
  interim/               dados extraídos dos anuários
  processed/etapa1_mercado/          P1-P3: volume, novos × usados, eletrificados (anuário)
  processed/etapa2_eletrificados/    P4-P8: tecnologia, modelos, marcas, geografia, frota
  processed/etapa3_renda_personas/   P9-P11: renda (POF), preços (FIPE/zero-km), personas
  powerbi/               PowerBI_Modelo_de_Dados.xlsx (13 tabelas)
docs/                    escopo, entrega parcial, auditoria das fontes, lista de fontes
powerbi/                 medidas DAX e instruções
scripts/                 scripts Python que geraram as planilhas
```

O arquivo mais completo é `data/processed/etapa3_renda_personas/Projeto_BI_Automotivo_Etapa3_Renda_Personas.xlsx`
(abas LEIA-ME, RESUMO, P1 a P11, Auditoria_Fontes e Fontes).

## Como reproduzir

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
# 1. baixe os anuários (data/raw/anuarios_fenabrave/LEIA-ME.md) e extraia os mensais:
python scripts/00_extrair_emplacamentos_anuarios.py
```

Os scripts `01` a `04` foram escritos no ambiente de trabalho do projeto e têm caminhos fixos (`/mnt/user-data/...`).
Ajuste as constantes `SRC`, `OUT` e os caminhos de leitura antes de rodar. Ordem: `01` → `02a` → `02b` → `03` → `04`.

## Fontes e limitações

- **FENABRAVE** (anuários 2021-2025): emplacamentos, razão de usados por novo, elétricos e híbridos por marca.
- **ABVE** (via comunicados e imprensa): tecnologia, modelos e geografia de 2025 e 2026.
- **IBGE POF 2017-2018**: distribuição de famílias por classe de renda (edição mais recente com resultados que encontrei).
- **Tabela FIPE** e preços de zero-km: reportagens e sites que reproduzem a FIPE; os valores mudam todo mês.
- **Banco Central**: juros de financiamento de veículos (2,04% ao mês, citados em jun/2026).

Pontos de atenção:
- A ABVE (sem micro-híbridos) e o anuário FENABRAVE (com micro-híbridos) usam critérios diferentes de "eletrificado". Nunca misture os dois no mesmo gráfico.
- Os usados são estimados a partir da razão publicada, com margem de cerca de ±1%.
- ABVE Data (painéis), SENATRAN e dados.gov.br não foram abertos diretamente.
- O catálogo de preços tem 34 veículos. Faltam SUV premium a gasolina, Volvo EX30, BMW 330e e elétricos usados.
- O perfil de uso diário (custo por km) ainda não está no modelo.
- A parcela não inclui IPVA, seguro, combustível nem manutenção.

Detalhes em [`docs/auditoria_fontes.md`](docs/auditoria_fontes.md) e [`docs/fontes.md`](docs/fontes.md).

## Dashboard (Power BI)

Veja [`powerbi/LEIA-ME.md`](powerbi/LEIA-ME.md). O guia completo (modelo de dados, 6 páginas e a pergunta final interativa)
fica em `docs/guia_dashboard_powerbi.md`.

## Licença

Código sob licença MIT (arquivo `LICENSE`). Os dados pertencem às suas fontes originais (FENABRAVE, ABVE, IBGE, FIPE, BCB);
confira os termos de cada uma antes de reutilizar.
