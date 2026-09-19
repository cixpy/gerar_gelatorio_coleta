# 📊 Gerador de Relatório de Atendimento - Fila COLETA

Uma aplicação em Python desenvolvida para automatizar a leitura, filtragem e consolidação de relatórios brutos de atendimento (em formatos CSV ou Excel) e gerar uma planilha Excel final formatada e pronta para impressão/exportação.

---

## 🚀 Funcionalidades

- **Interface Gráfica Integrada:** Seleção simples de arquivos de entrada e saída por meio de caixas de diálogo nativas (`Tkinter`).
- **Suporte Multi-formato:** Aceita arquivos de entrada em `.csv`, `.xlsx` e `.xlsm`.
- **Detecção Inteligente:**
  - **Múltiplas Codificações e Delimitadores:** Trata automaticamente arquivos CSV em `UTF-8`, `UTF-8-SIG`, `CP1252` e `Latin1`, identificando separadores como `;`, `,`, `\t` e `|`.
  - **Localização de Cabeçalhos Flexível:** Identifica colunas dinamicamente, ignorando acentuações e variações de maiúsculas/minúsculas.
  - **Formatos de Horário Flexíveis:** Processa durações de atendimento fornecidas em segundos, minutos/segundos, horas/minutos/segundos ou frações numéricas nativas do Excel.
- **Filtragem Específica:** Processa automaticamente apenas os registros pertencentes à fila **COLETA DE IMAGEM**.
- **Métricas Consolidadas:**
  - Nome do funcionário.
  - Média de tempo de atendimento (formatada nativamente como `[h]:mm:ss` no Excel).
  - Total de atendimentos por funcionário e geral.
  - Contagem de avaliações neutras/negativas (`REGULAR`, `RUIM`, `NÃO OPINOU`).
- **Formatação de Planilha Profissional:**
  - Layout pronto para impressão em folha A4 (Orientação Paisagem).
  - Congelamento de painéis (`Freeze Panes`) e inclusão automática de Autofiltro.
  - Cabeçalhos, totais e bordas estilizadas via `openpyxl`.

---

## 🛠️ Pré-requisitos

Para executar o script, você precisará ter o **Python 3.8+** instalado na sua máquina, além da biblioteca `openpyxl`.

### Dependências necessárias:

```bash
pip install openpyxl
