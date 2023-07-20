# XML para Excel
#### Video Demo:  <https://www.youtube.com/watch?v=cuUm4yigKQE>
#### Descrição: Projeto final do curso CS50's Introduction to Programming with Python. Esse projeto visa automatizar o processo de preencher uma tabela do excel com informações de uma ou mais notas fiscais eletrônicas (arquivo XML).
Passos de execução do programa:
- ler os itens e quantidade desses itens dos arquivos XML
- guardar essas informações em um dicionario (item: quantidade)
- verificar se há itens repedidos entre as duas empresas, se for caso, o usuário deverá informar de qual empresa é esse item
- perguntar para o usuário qual a janela que envio (1 .. 10) desses itens
- preencher a tabela com as informações coletadas

Estrutura do código:
 - `load_xml() -> dict`: ler os arquivo XML e obter as informações do itens (código do produto e quantidade) e armazenar essas informações em dicionário, separado por empresa.

 - `check_duplicate(mavifer: dict, izamac: dict) -> list`: recebe dois dicionários e verifica os itens que estão presentes nos dois dicionários e retorna uma lista com esses itens.

- `get_items_excel(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict`: recebe um Worksheet e retorna dicionário com itens e linha desses itens.

- `get_first_empty(items: dict) -> int`: recebe um dicionário com os itens e suas linhas (dicionário gerado pela função `get_items_excel()`) e retorna o número da próxima linha que está vazia.

Como executar:
O programa requer a instalação do modulo `openpyxl`, que ser instalado pelo comando:

`pip install openpyxl`

No diretório `/xml` devem ficar os arquivos xml que vão ser lidos.
No diretório `/excel` deve ficar o arquivo da planilha em excel.

Execute o comando: `python project.py`