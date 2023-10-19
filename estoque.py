import csv
import datetime
import os
from tabulate import tabulate

# Nome do arquivo CSV de entrada (PEDIDOS.csv)
input_csv_file_pedidos = "PEDIDOS.csv"

# Nome do arquivo CSV de produtos
input_csv_file_produtos = "PRODUTOS.csv"

# Nome do arquivo de texto de saída
data_de_hoje = datetime.date.today()
output_txt_file = f"estoque_{data_de_hoje}.txt"

try:
    # Dicionário para armazenar a correspondência entre Cod Red e Descrição a partir do arquivo PRODUTOS.csv
    cod_red_to_desc_un = {}

    # Dicionário para armazenar os produtos por grupo
    produtos_por_grupo = {}

    # Ler o arquivo de produtos e criar o mapeamento entre Cod Red, Descrição e Un, bem como agrupar por grupo
    with open(input_csv_file_produtos, 'r', newline='') as produtos_file:
        produtos_reader = csv.DictReader(produtos_file, delimiter=';')
        for produto_row in produtos_reader:
            cod_red = produto_row.get("Reduzido")
            descricao = produto_row.get("Descricao")
            un = produto_row.get("Un")
            grupo = produto_row.get("Grupo")

            # Preencher o dicionário cod_red_to_desc_un
            cod_red_to_desc_un[cod_red] = (descricao, un)

            # Preencher o dicionário produtos_por_grupo
            if grupo in produtos_por_grupo:
                produtos_por_grupo[grupo].append((cod_red, descricao, un))
            else:
                produtos_por_grupo[grupo] = [(cod_red, descricao, un)]

    # Dicionário para armazenar o cálculo da quantidade da coluna "Quant CX" por descrição
    quantidades_por_descricao = {}

    with open(input_csv_file_pedidos, 'r', newline='') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=';')
        for row in csv_reader:
            cod_red = row.get("Cod Red")
            quant_cx = int(row.get("Quant CX").replace(',', '.'))

            romaneio = row.get("Romaneio")
            mot_dev = row.get("Mot Dev")
            ocorr = row.get("Ocorr")

            # Aplicar os filtros, incluindo a verificação de "Romaneio" igual a 0
            if  romaneio == "0" and mot_dev != "1" and quant_cx > 0:
                descricao, un = cod_red_to_desc_un.get(cod_red, ("Descrição não encontrada", "Unidade não encontrada"))
                if descricao in quantidades_por_descricao:
                    quantidades_por_descricao[descricao] = (quantidades_por_descricao[descricao][0] + quant_cx, un)
                else:
                    quantidades_por_descricao[descricao] = (quant_cx, un)

    # Preparando os dados para a tabela
    data = []

    for grupo, produtos in produtos_por_grupo.items():
        produtos_com_quantidade = [p for p in produtos if quantidades_por_descricao.get(p[1], (0, ''))[0] > 0]
        if produtos_com_quantidade:
            data.append(["Grupo: " + grupo, "", ""])

            for cod_red, descricao, un in produtos_com_quantidade:
                quantidade, unidade = quantidades_por_descricao.get(descricao, (0, un))
                data.append([cod_red, descricao, f"{quantidade} {unidade}"])

    # Criando a tabela formatada
    table = tabulate(data, headers=["Código Reduzido", "Descrição", "Quantidade"], tablefmt="grid")

    # Escreva o resultado no arquivo de texto
    with open(output_txt_file, 'w') as txt_file:
        txt_file.write(table)

    print(f"Resultado escrito em {output_txt_file}.")

    if os.name == 'nt':
        os.system(f"notepad.exe {output_txt_file}")

except FileNotFoundError:
    print(f"Arquivo não encontrado.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")
