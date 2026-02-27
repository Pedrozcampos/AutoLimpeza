# AutoClean Contábil
O AutoClean Contábil é uma ferramenta de automação em Python desenvolvida para resolver um dos maiores gargalos em escritórios de contabilidade: a limpeza e padronização de extratos contábeis complexos extraídos de sistemas.

## O Problema
Muitos sistemas contábeis geram relatórios onde o nome da conta aparece apenas em uma linha de cabeçalho, e não em cada linha de lançamento. Isso impede a análise rápida via Tabela Dinâmica ou Power BI. Além disso, o layout das colunas pode variar de arquivo para arquivo.

## Solução
Esta automação utiliza **Pandas** e **Regex** para:
1.  **Herança de Conta:** Identifica o cabeçalho "Conta", extrai o Código e Nome e os propaga para todas as transações correspondentes.
2.  **Padronização Inteligente:** Garante as colunas essenciais (`Data`, `Crédito`, `Débito`, `Histórico`, `Conta`) e preserva quaisquer colunas adicionais que existam no arquivo original.
3.  **Interface (GUI):** Construída com **CustomTkinter**, permitindo que usuários sem conhecimento técnico em programação utilizem a ferramenta com facilidade.
4.  **Limpeza Numérica:** Converte automaticamente formatos de moeda brasileiros (1.250,00) para formatos computacionais (1250.00).

## Tecnologias Utilizadas
* **Python**: Principal Tecnologia.
* **Pandas**: Processamento e manipulação de dados.
* **CustomTkinter**: Interface gráfica.
* **Openpyxl**: Engine para leitura e escrita de arquivos Excel (.xlsx).
* **Re (Regex)**: Extração de padrões de texto.

## Como usar
1. Instale as dependências:

   pip install pandas openpyxl customtkinter
