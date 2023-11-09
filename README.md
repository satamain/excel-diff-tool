# Welcome to ExcelDiffTool!

Have you ever been in a situation where you and a colleague are working on the same Excel spreadsheet but end up with different versions? Maybe your colleague saved under a new file name, or a mishap with cloud storage resulted in separate copies. There are several reasons why this could happen, leading to the same challenge: two spreadsheets with different content and the difficulty of knowing exactly what changed.

This is where ExcelDiffTool shines! 🌟 This simple program is designed to help you quickly identify the differences between two spreadsheets. It doesn't just find changes; it highlights exactly where they are—in terms of rows and columns—and shows what's been altered.

## How Does It Work? 🛠️

1. **Sheet Selection**:
   To set up the sheets you want to compare, modify the variables in the script to the paths of your spreadsheet files. Inside the script `ExcelDiffTool.py`, find the following lines:

   ```python
   sheet1_path = 'sheet-origin.xlsx'
   sheet2_path = 'sheet-changed.xlsx'
2. **Script Execution**: Run the `ExcelDiffTool.py` file using Python to start the comparison.
3. **Difference Analysis**: The program then analyzes the spreadsheets and identifies any differences.
4. **Clear Reporting**: Finally, you get a clear report showing where the discrepancies lie, so you can take the necessary actions.

## Sample Comparison Result

Using the ExcelDiffTool, we compared the following sheets from the repository:

- **Sheet 1**: `sheet-origin.xlsx`
- **Sheet 2**: `sheet-changed.xlsx`

The differences found are partially listed in the table below:

| Line | Column | Value in Sheet 1      | Value in Sheet 2  |
|------|--------|-----------------------|-------------------|
| 6    | D      | 1541                  | 1415              |
| 6    | E      | 4                     | 7                 |
| 20   | D      | 4.8                   | 4.2               |
| 28   | B      | W. Cleon Skousen      | Cleon Skousen     |
| ...  | ...    | ...                   | ...               |
| 699  | C      | 4.6                   | 2.3               |

Please note: The table above is just a partial display of the differences. When you run the script, it will show all the discrepancies between the sheets in detail.

The data used for testing this script was obtained from [Amazon Top 50 Bestselling Books 2009-2019](https://www.kaggle.com/datasets/sootersaalu/amazon-top-50-bestselling-books-2009-2019) on Kaggle.

***
***
***

# Bem-vindo ao ExcelDiffTool!


Já esteve em uma situação onde você e um colega estão trabalhando na mesma planilha do Excel, mas acabaram com versões diferentes? Talvez seu colega tenha salvo com um novo nome de arquivo ou um pequeno contratempo com o armazenamento na nuvem criou cópias separadas. Existem várias razões para isso acontecer, levando ao mesmo desafio: duas planilhas com conteúdos diferentes e a dificuldade de saber exatamente o que mudou.

É aqui que o ExcelDiffTool brilha! 🌟 Este simples programa foi criado para ajudá-lo a identificar rapidamente as diferenças entre duas planilhas. Ele não apenas localiza as alterações, mas também destaca exatamente onde elas estão — em termos de linhas e colunas — e mostra o que foi alterado.

## Como Funciona? 🛠️

1. **Seleção das Planilhas**:
   Para configurar as planilhas que deseja comparar, modifique as variáveis no script para os caminhos dos seus arquivos de planilha. Dentro do script `ExcelDiffTool.py`, encontre as seguintes linhas:

   ```python
   sheet1_path = 'sheet-origin.xlsx'
   sheet2_path = 'sheet-changed.xlsx'
2. **Execução do Script**: Execute o arquivo `ExcelDiffTool.py` usando Python para iniciar a comparação.
3. **Análise de Diferenças**: O programa então analisa as planilhas e identifica quaisquer diferenças.
4. **Relatório Claro**: Por fim, você recebe um relatório claro mostrando onde as discrepâncias ocorrem, para que possa tomar as ações necessárias.

## Resultado da Comparação de Amostra

Utilizando o ExcelDiffTool, comparamos as seguintes folhas do repositório:

- **Folha 1**: `sheet-origin.xlsx`
- **Folha 2**: `sheet-changed.xlsx`

As diferenças encontradas estão parcialmente listadas na tabela abaixo:

| Linha | Coluna | Valor na Folha 1      | Valor na Folha 2  |
|-------|--------|-----------------------|-------------------|
| 6     | D      | 1541                  | 1415              |
| 6     | E      | 4                     | 7                 |
| 20    | D      | 4.8                   | 4.2               |
| 28    | B      | W. Cleon Skousen      | Cleon Skousen     |
| ...   | ...    | ...                   | ...               |
| 699   | C      | 4.6                   | 2.3               |

A tabela acima é apenas uma exibição parcial das diferenças. Ao executar o script, ele mostrará todas as discrepâncias entre as folhas em detalhes.

Os dados utilizados para os testes deste script foram obtidos do conjunto [Amazon Top 50 Bestselling Books 2009-2019](https://www.kaggle.com/datasets/sootersaalu/amazon-top-50-bestselling-books-2009-2019) no Kaggle.