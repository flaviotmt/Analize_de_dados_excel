import xlrd
import xlsxwriter
from questoes_escala_linear import questoes_escala_linear
from questoes_nao_sei import questoes_nao_sei
from questoes_sim_nao import questoes_sim_nao


#inicializando o leitor de Excel
path = "Compartilhamento.xlsx"
arquivo_analise = xlrd.open_workbook(path)
aba_analise = arquivo_analise.sheet_by_index(0)

#Inicializando o escritor de Excel
arquivo = xlsxwriter.Workbook("analise.xlsx")
questoes = ["Questao1", "Questao2", "Questao3", "Questao4", "Questao5", "Questao6", "Questao7", "Questao8"]
dicio_questoes = {"Questao1": [0,3], "Questao2": [1,4], "Questao3": [0,5], "Questao4": [0,6], "Questao5": [0,7], "Questao6": [1,8], "Questao7": [1,9], "Questao8": [2,10]}
for nome_aba in questoes:
    aba = arquivo.add_worksheet(nome_aba)
    if dicio_questoes[nome_aba][0] == 0:
        questoes_escala_linear(dicio_questoes[nome_aba][1], aba, aba_analise)
    elif dicio_questoes[nome_aba][0] == 1:
        questoes_sim_nao(dicio_questoes[nome_aba][1], aba, aba_analise)
    elif dicio_questoes[nome_aba][0] == 2:
        questoes_nao_sei(dicio_questoes[nome_aba][1], aba, aba_analise)

arquivo.close()
