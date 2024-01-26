import csv
def main():
    paths = "C:\\Users\\Nota Fiscal\\Downloads\\"
    estoque = ["produtos.csv", "produtos (1).csv"]
    recebidos = ["Nota Fiscal Entrada Itens.csv", "Nota Fiscal Entrada Itens (1).csv"]

    itens_recebidos = [{}, {}]
    itens_estoque = [{}, {}]

    # lendo estoque
    for i in range(len(estoque)):
        with open(f"{paths}{estoque[i]}", encoding="ANSI") as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';')

            for row in reader:
                item = row["Referencia"]
                qtd = row["Estoque"]
                itens_estoque[i][item] = qtd

    # lendo itens recebidos
    for i in range(len(recebidos)):
        with open(f"{paths}{recebidos[i]}") as csvfile:
            for _ in range(3):
                csvfile.__next__()
            reader = list(csv.DictReader(csvfile, delimiter=';'))[:-4]

            for row in reader:
                item = row["Cod. Forn."]
                qtd = row["Qtde"]
                itens_recebidos[i][item] = qtd

    print(itens_estoque)
    print(itens_recebidos)


if __name__ == "__main__":
    main()
