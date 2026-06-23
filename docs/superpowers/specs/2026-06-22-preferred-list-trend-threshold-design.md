# Tibia Hunt Analyzer — Preferred List, Thresholds & Historical Trend

## Resumo

Três melhorias no gerador de planilhas Excel do Hunt Analyser:
1. Nova aba "Preferred List" com top 5 criaturas por kills/h
2. Marcação visual (verde/vermelho) em sessões com XP/h desviando >30% da média do grupo
3. Seção de tendência histórica nas abas de detalhe de grupo
4. Limpeza e padronização das abas existentes

---

## 1. Aba "Preferred List"

Nova aba no workbook, posicionada após o Resumo.

**Colunas:**
| # | Criatura | Kills/h | ±% Média | Imagem |
|---|----------|---------|----------|--------|

- **Kills/h** = soma de kills da criatura em TODAS as sessões / soma de horas de TODAS as sessões
- **±% Média** = (kills_h_da_criatura - media_kills_h_do_top5) / media_kills_h_do_top5 * 100
- **Imagem** = PNG 64×64 baixada da TibiaWiki (mesmo processo das abas de detalhe)
- Ordenada decrescente por Kills/h, top 5 apenas
- Sem coluna "Total Kills" (irrelevante)

---

## 2. Thresholds — Marcação Visual

Aplicado em duas tabelas:

### a) Tabela "Detalhes das Sessoes neste Grupo" (aba de detalhe)
- Nova coluna "% Média" após "Profit/h"
- Fórmula: `(session_calc_raw_xp_h - group_avg_raw_xp_h) / group_avg_raw_xp_h * 100`
- Se **< -30%**: linha toda com fundo vermelho (`FFCCCC`)
- Se **> +30%**: linha toda com fundo verde (`CCFFCC`)

### b) Seção "Detalhamento por Sessao" (aba Resumo)
- Mesma lógica, mesma coluna extra, mesma pintura condicional

A média de referência é sempre a **média aritmética do grupo** ao qual a sessão pertence.

---

## 3. Tendência Histórica

Nova seção dentro de cada aba de detalhe de grupo (apenas grupos com 2+ sessões).

Posicionada após a tabela "Criaturas Mortas" e antes de "Detalhes das Sessoes".

**Elementos:**
- **Cabeçalho**: "Tendência Histórica"
- **Indicador de tendência**: compara a média da primeira metade das sessões (ordenadas cronologicamente pela data) com a média da segunda metade:
  - ↑ Melhorando: segunda metade >= 105% da primeira
  - ↓ Piorando: segunda metade <= 95% da primeira
  - → Estável: entre 95% e 105%
- **Melhor sessão** e **Pior sessão** destacadas visualmente na tabela de detalhes (já incluso no threshold section 2)
- **Gráfico**: substituir o gráfico de barras "XP/h por Sessao" por um gráfico de **linha** com a média do grupo plotada como linha horizontal de referência

---

## 4. Limpeza e Padronização

### Aba Resumo
- Renomear "TOP 5 LUCRO" → "TOP 5 Profit/h"
- Remover seção "Detalhamento por Sessao" (redundante com as abas de detalhe)

### Abas de detalhe de grupo
- Remover gráfico "Dano/h vs XP/h vs Lucro/h por Sessao" (poluído)
- Manter apenas gráfico de linha "XP/h por Sessao" (com média como referência)
- Padronizar largura de colunas entre todas as abas
- Prefixo nos nomes de sheet: `"Detalhe - {criatura}"`

### Geral
- Coluna "Sessao" com nome de arquivo de origem legível
- Números com separador de milhar consistente

---

## Arquitetura da Implementação

### Novas funções
- `compute_global_kills_h(sessions)` → retorna dict {criatura: kills/h} para todas as sessões
- `build_preferred_list_sheet(wb, global_data)` → cria a nova aba
- `add_threshold_formatting(ws, group_avg, start_row, end_row, col_xp)` → pinta linhas
- `build_trend_section(ws, group)` → escreve indicador de tendência e modifica gráfico

### Funções modificadas
- `build_summary_sheet()` → adicionar coluna % Média + pintura; remover Detalhamento por Sessao
- `build_hunt_detail_sheet()` → adicionar seção de tendência, modificar gráfico, pintura de sessões
- `main()` → chamar `build_preferred_list_sheet()` entre `build_summary_sheet` e `build_hunt_detail_sheet`

### Dependências
- Nenhuma nova biblioteca necessária (openpyxl já tem PatternFill, Chart já suporta line)
- Dados já disponíveis: `group["avg_raw_xp_h"]`, `s["calc_raw_xp_h"]`, `s["date"]`
