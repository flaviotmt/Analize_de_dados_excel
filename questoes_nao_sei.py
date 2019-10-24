import xlrd
import xlsxwriter

def questoes_nao_sei(num_coluna, aba, aba_analise):

        #Inicialização de Dicionários para contabilizar votos
        id1 = {"Sim":0,"Não":0, "Não sei":0}
        id2 = {"Sim":0,"Não":0, "Não sei":0}
        id3 = {"Sim":0,"Não":0, "Não sei":0}
        id4 = {"Sim":0,"Não":0, "Não sei":0}
        id5 = {"Sim":0,"Não":0, "Não sei":0}
        id6 = {"Sim":0,"Não":0, "Não sei":0}
        masc = {"Sim":0,"Não":0, "Não sei":0}
        fem = {"Sim":0,"Não":0, "Não sei":0}

        #Identificando e listanto por idade
        for y in range(1, aba_analise.nrows):
            if aba_analise.cell_value(y, num_coluna) != "":
                if aba_analise.cell_value(y, 2) == "Entre 18 e 22 anos":
                    id1[aba_analise.cell_value(y, num_coluna)] += 1
                elif aba_analise.cell_value(y, 2) == "Entre 23 e 27 anos":
                    id2[aba_analise.cell_value(y, num_coluna)] += 1
                elif aba_analise.cell_value(y, 2) == "Entre 28 e 32 anos":
                    id3[aba_analise.cell_value(y, num_coluna)] += 1
                elif aba_analise.cell_value(y, 2) == "Entre 33 e 37 anos":
                    id4[aba_analise.cell_value(y, num_coluna)] += 1
                elif aba_analise.cell_value(y, 2) == "Entre 38 e 42 anos":
                    id5[aba_analise.cell_value(y, num_coluna)] += 1
                elif aba_analise.cell_value(y, 2) == "Acima de 43 anos":
                    id6[aba_analise.cell_value(y, num_coluna)] += 1

        #Identificando e listanto por genero
        for y in range(1, aba_analise.nrows):
            if aba_analise.cell_value(y, num_coluna) != "":
                if aba_analise.cell_value(y, 1) == "Masculino":
                    masc[aba_analise.cell_value(y, num_coluna)] += 1
                elif aba_analise.cell_value(y, 1) == "Feminino":
                    fem[aba_analise.cell_value(y, num_coluna)] += 1

        #Nomes dos cabeçalhos das tabelas
        coluna_1 = ["Entre 18 e 22 anos", "Entre 23 e 27 anos", "Entre 28 e 32 anos", "Entre 33 e 37 anos", "Entre 38 e 42 anos", "Acima de 43 anos"]
        cabecalho = ["", "Sim", "Não", "Não sei"]
        coluna_2 = ["Feminino", "Masculino"]

        #Criando Cabeçalho 1
        for item in range(len(coluna_1)):
            aba.write(item+1, 0, coluna_1[item])
        for item in range(len(cabecalho)):
            aba.write(0, item, cabecalho[item])

        #Criando Cabeçalho 2
        for item in range(len(coluna_2)):
            aba.write(len(coluna_1)+3+item, 0, coluna_2[item])
        for item in range(len(cabecalho)):
            aba.write(len(coluna_1)+2, item, cabecalho[item])
            

        #Listando resultados por idade
        lista = [id1, id2, id3, id4, id5, id6]
        for linha in range(len(coluna_1)):
            for coluna in range(len(cabecalho)-1):
                aba.write(linha+1, coluna+1, lista[linha][cabecalho[coluna+1]])

        #Listanto resultados por genero
        lista2 = [fem, masc]
        for linha in range(len(coluna_2)):
            for coluna in range(len(cabecalho)-1):
                aba.write(len(coluna_1)+linha+3, coluna+1, lista2[linha][cabecalho[coluna+1]])
