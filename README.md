# Tibia Hunt Analyzer - Excel Generator

Gera uma planilha Excel comparativa a partir de N sessoes do Hunt Analyser do Tibia.
Agrupa sessoes com as mesmas criaturas, calcula medias e **rankeia por Raw XP/h**.

## Instalacao

```bash
pip install -r requirements.txt
```

## Uso

### 1. Modo exemplo (2 sessoes inclusas)
```bash
python hunt_analyzer_excel.py
```

### 2. Modo paste (cole quantas sessoes quiser)
```bash
python hunt_analyzer_excel.py --paste
```
Cole as sessoes separadas por linhas em branco, Ctrl+Z + Enter.

### 3. Um ou mais arquivos
```bash
python hunt_analyzer_excel.py -f sessao1.txt -f sessao2.txt -f sessao3.txt -o resultado.xlsx
```

### 4. Diretorio com varias sessoes
Salve cada sessao em um arquivo `.txt` ou `.log` dentro de uma pasta:
```bash
python hunt_analyzer_excel.py --dir minhas_sessoes -o resultado.xlsx
```

## Como alimentar o rankeamento

1. Copie o output do Hunt Analyser de cada hunt para um arquivo `.txt`
2. Coloque todos os arquivos em uma pasta
3. Execute: `python hunt_analyzer_excel.py --dir pasta_das_sessoes -o ranking.xlsx`

A planilha vai:
- Agrupar sessoes com as mesmas criaturas
- Calcular a media de Raw XP/h e Profit/h por grupo
- **Ordenar do maior XP/h para o menor** (ranking)
- Gerar uma aba de detalhes por hunt com imagens das criaturas

## Funcionalidades

- Aceita **N sessoes** (arquivos, diretorio ou paste)
- Calcula Raw XP/h = Raw XP Gain / horas
- Calcula Profit/h = Balance / horas
- Agrupa sessoes com as mesmas criaturas automaticamente
- Rankeia grupos por Raw XP/h (maior primeiro)
- Busca imagens das criaturas no TibiaWiki e incorpora na planilha
- Aba "Resumo" com comparativo, medias e ranking
- Aba "Hunt ..." por grupo com detalhamento por criatura + imagens PNG
