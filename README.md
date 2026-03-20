# 💰 Pipeline ETL — Passivo UNIMED para ADMEX
> Motor FIFO de abatimento progressivo de passivos financeiros — transforma planilha consolidada de dívidas em layout CSV pronto para importação no sistema ADMEX.

![Python](https://img.shields.io/badge/Python-3.10+-8b5cf6?style=flat-square&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-2.x-8b5cf6?style=flat-square&logo=pandas&logoColor=white)
![Status](https://img.shields.io/badge/Status-Em%20Produção-22c55e?style=flat-square)
![Versão](https://img.shields.io/badge/Versão-v9.1-8b5cf620?style=flat-square)

---

## 📌 Sobre o Projeto

Pipeline ETL especializado para transformação de dados de passivos do plano UNIMED para o layout de importação do sistema ADMEX.

O script lê uma planilha Excel consolidada com histórico de débitos e acertos de múltiplos beneficiários, aplica um **motor FIFO** que abate progressivamente a dívida mais antiga com cada acerto em folha ou avulso, e gera dois arquivos CSV: `Entradas_ADMEX.csv` (débitos) e `Acertos_ADMEX.csv` (abatimentos).

**Problema resolvido:** a transformação manual desses dados exigia conhecimento profundo das regras do ADMEX e era propensa a erros de sequência de abatimento. O motor FIFO garante a ordem cronológica correta em todos os casos.

---

## 🚀 Funcionalidades

- ✅ **Motor FIFO** — abate a dívida mais antiga primeiro, progressivamente por valor disponível
- ✅ **Parsing inteligente de datas** — aceita `jan/24`, `01/2024`, `jan 24`, `01-24` e qualquer variação
- ✅ **Mapeamento automático de colunas** — identifica colunas de débito, folha e avulso por nome, sem configuração manual
- ✅ **ID Movimento** — classifica automaticamente entre Débito/Acerto (ID=1) e Informativo (ID=5)
- ✅ **Seletor de arquivo via UI** — abre janela gráfica para selecionar o Excel se executado sem argumentos
- ✅ **Formato de data ADMEX nativo** — converte para `mmm/aa` (ex: `jan/24`) automaticamente
- ✅ **Consolidação por CPF + Referência** — agrupa e soma valores duplicados antes de exportar

---

## 🛠️ Stack

| Tecnologia | Uso |
|---|---|
| Python 3.10+ | Lógica principal |
| Pandas | Leitura e transformação do Excel |
| Regex | Parsing flexível de datas nas colunas |
| argparse | Execução via linha de comando ou UI |
| tkinter | Seletor de arquivo gráfico (opcional) |
| unicodedata | Normalização de nomes de colunas |

---

## 📁 Estrutura

```
pipeline-passivo-admex/
├── transformar_passivo_unimed.py   # Script principal (v9.1)
├── requirements.txt
├── .gitignore
└── README.md

# Input esperado:
└── Consolidado_UNIMED.xlsx    # Planilha com abas por entidade (não versionada)

# Output gerado automaticamente em ./saida/:
├── Entradas_ADMEX.csv         # Débitos/obrigações
└── Acertos_ADMEX.csv          # Abatimentos (folha, avulso, informativo)
```

---

## ⚙️ Como Executar

```bash
pip install -r requirements.txt
```

**Com interface gráfica (recomendado):**
```bash
python transformar_passivo_unimed.py
# Abre janela para selecionar o arquivo Excel
```

**Via linha de comando:**
```bash
python transformar_passivo_unimed.py --xlsx "Consolidado_UNIMED.xlsx" --outdir "./saida"
```

---

## 🔧 Lógica do Motor FIFO

```
Beneficiário X tem débitos em: jan/24 (R$150), mar/24 (R$200), jun/24 (R$180)

Acerto em folha jun/24: R$280
  → Abate jan/24 completamente:  -R$150  (saldo: R$0)
  → Abate mar/24 parcialmente:   -R$130  (saldo: R$70)
  → mar/24 permanece com R$70 de saldo devedor
  → jun/24 intocado
```

---

## 📈 Exemplo de Output

```
[START] Lendo arquivo 'Consolidado_UNIMED.xlsx'...
  -> Processando aba: UNIMED_HOB
  -> Processando aba: UNIMED_SAMU
  -> Processando aba: UNIMED_URBEL

[SUCESSO] saida/Entradas_ADMEX.csv gerado com 1.243 linhas.
[SUCESSO] saida/Acertos_ADMEX.csv gerado com 987 linhas.

Pressione ENTER para fechar a janela...
```

---

## 👤 Autor

**Gustavo** — Dev & Founder · Inside.co

[![LinkedIn](https://img.shields.io/badge/LinkedIn-8b5cf6?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/gustavo-henriquesp/)
[![Portfolio](https://img.shields.io/badge/Portfolio-8b5cf6?style=flat-square&logo=netlify&logoColor=white)](https://seusite.netlify.app)
[![Email](https://img.shields.io/badge/Email-8b5cf6?style=flat-square&logo=gmail&logoColor=white)](mailto:ghspdm@gmail.com)

---
> *"Construo soluções que outros apenas descrevem em planilhas."*
