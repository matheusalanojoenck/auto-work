from xml.dom import minidom
import os
import openpyxl

def main():
    total = load_xml()

    if len(itens := check_duplicate(total["MAVIFER"], total["IZAMAC"])) > 0:
        for item in itens:
            while (True):
                empresa = input(f"{item} é de qual empresa (mavifer ou izamac)? ").upper()
                if empresa in ["MAVIFER", "IZAMAC"]:
                    if empresa == "MAVIFER":
                        value = total["IZAMAC"].pop(item)
                        total[empresa][item] += value
                    elif empresa == "IZAMAC":
                        value = total["MAVIFER"].pop(item)
                        total[empresa][item] += value
                    break

    apontamento_expedicao(total)


# le os arquivos xml e retorna um dicionario com os itens e quantidades desses xml
def load_xml() -> dict:
    xml_path = "C:\\Users\\Nota Fiscal\\Documents\\GitHub\\auto-work\\xml\\"
    total = {"MAVIFER": {}, "IZAMAC": {}}
    for file_name in os.listdir(xml_path):
        nfe = minidom.parse(xml_path + file_name)
        #nnf = nfe.getElementsByTagName("nNF")[0].childNodes[0].nodeValue
        empresa = nfe.getElementsByTagName("xNome")[0].childNodes[0].nodeValue.split(" ")[0].strip()
        itens = nfe.getElementsByTagName("prod")
        #cfop = itens[0].getElementsByTagName("CFOP")[0].childNodes[0].nodeValue
        #print(f"Empresa: {empresa} | NF {nnf} | CFOP: {cfop}")

        for item in itens:
            cod_prod = item.getElementsByTagName("cProd")[0].childNodes[0].nodeValue.strip()
            qtd_prod = item.getElementsByTagName("qCom")[0].childNodes[0].nodeValue
            if cod_prod in total[empresa]:
                total[empresa][cod_prod] += float(qtd_prod)
            else:
                total[empresa][cod_prod] = float(qtd_prod)
    return total

# recebe dois dicionarios e verifica se há algum item duplicado e retorna uma lista com os itens duplicados
def check_duplicate(mavifer: dict, izamac: dict) -> list:
    duplicate = []
    for key in mavifer.keys():
        if key in izamac.keys():
            duplicate.append(key)

    return duplicate

# recebe uma tabela e retorna uma lista com os itens já estão na planilha
def get_items_excel(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict:
    items = {}
    for cell in ws['A6':'A35']:
        if cell[0].value != None:
            items[cell[0].value.upper()] = cell[0].row
    return items

# Retorna um inteiro indicando qual a proxima linha da tabela que esta vazia
def get_first_empty(items: dict) -> int:
    row = 5
    for item in items.keys():
        if items[item] > row:
            row = items[item]
    return row+1

#recebe um dicionario com os itens das duas empresas e salva na planilha
def apontamento_expedicao(total: dict):
    janelas = {
        1: 'B',
        2: 'C',
        3: 'D',
        4: 'E',
        5: 'F',
        6: 'G',
        7: 'H',
        8: 'I',
        9: 'J',
        10: 'K'
    }

    while True:
        num_janela = int(input("Qual janela da carga?(1 .. 10) ")) + 1
        if num_janela in janelas:
            break

    path = "C:\\Users\\Nota Fiscal\\Documents\\GitHub\\auto-work\\excel\\Apontamento Expedição.xlsx"
    wb = openpyxl.load_workbook(path)

    # range header B4:K4
    # range items A6:A35
    # range qtd B6:K35
    # col = B:K
    # row = 6:35

    for emp_name in ["MAVIFER", "IZAMAC"]:
        ws = wb[emp_name]

        for item in total[emp_name].keys():
            items = get_items_excel(ws)
            if item in items.keys():
                cell_coord = f"{janelas[num_janela]}{items[item]}"
                ws[cell_coord] = total[emp_name][item]
            else:
                row = get_first_empty(items)
                cell_coord = f"{janelas[num_janela]}{row}"
                ws[f"A{row}"] = item
                ws[cell_coord] = total[emp_name][item]

    wb.save("C:\\Users\\Nota Fiscal\\Documents\\GitHub\\auto-work\\excel\\Apontamento Expedição.xlsx")

def faturamento():
    pass

def romaneio(itens: dict):
    pass
if __name__ == "__main__":
    main()
