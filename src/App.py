import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import model

import formatTable as ft

#= = = Variáveis 
# Obter o diretório atual do arquivo de script
diretorio_atual = os.path.dirname(os.path.realpath(__file__))

func_path = os.path.join(diretorio_atual,"..","data","func.xlsx")

func_atualizado_path = os.path.join(diretorio_atual,"..","data","func_atualizado.xlsx")

func_mesclado_path = os.path.join(diretorio_atual,"..","data","func_final.xlsx")

func_df = model.Dataframe(func_path)

func_df.add_column("SITUACAO","PENDENTE")

func_atualizado_df = model.Dataframe(func_atualizado_path)

func_novos = [] #lista dos funcionarios que foram adicionados na tabela func_atualizado mas não na original

""" 
# Remover acentos de todas as colunas do DataFrame
for coluna in func_df.columns:
    func_df[coluna] = func_df[coluna].apply(lambda x: unidecode(str(x)) if pd.notnull(x) else x)

# Remover acentos de todas as colunas do DataFrame
for coluna in func_atualizado_df.columns:
    func_atualizado_df[coluna] = func_atualizado_df[coluna].apply(lambda x: unidecode(str(x)) if pd.notnull(x) else x)
"""

#Percorremos o dataframe dos funcionarios atualizados
for index, func_atualizado in func_atualizado_df.get_iterrows():
        
    # Encontra o funcionário correspondente no DataFrame original
    func_found = func_df.get_rows_by_collumn('MATRICULA', func_atualizado['MATRICULA']) 
    
    # Se o funcionário estiver presente no DataFrame original, UMA OU + VEZEZ atualizar todos os valores em branco
    if len(func_found) > 0:
        for func in func_found:
            func_index = func[0]  # Pega o índice do funcionário encontrado
            func_df.alter_row(func_index,"SITUACAO","ALTERADO")
    
            # Iterar sobre a linha do funcionario do DataFrame original
            for coluna in func_df.columns:
                # Se o valor da coluna no DataFrame original estiver em branco, preencher com o valor correspondente do DataFrame atualizado
                if func_df.get_column(func_index, coluna) == "":
                    func_df.alter_row(func_index, coluna, func[coluna].upper())  
                    
    #Caso não encontre o funcionario no dataframe, adicione-o no dataframe original
    else:
        func_found.loc["SITUACAO"] = "ADICIONADO"
        func_df.add_row(func_found)     
        
# Salvar o DataFrame mesclado
try:
    func_df.save(func_mesclado_path)
except PermissionError as PE:
    print("Não foi acessar a tabela...\n Tente fechar o arquivo, caso esteja aberto")
    print(PE)
    input()
    # Clearing the Screen
    os.system('cls')
   
# Carregar o arquivo Excel mesclado
workbook = load_workbook(func_mesclado_path)
ws = workbook.active

""" 
quant_func = len(func_df) + len(func_novos)
#Inserir no final da tabela os funcionarios novos
if(len(func_novos)>0):
    func_novo = 0
    for row in ws.iter_rows(min_row=len(func_df) + 2, max_row= quant_func +  1, min_col=1, max_col=len(func_df.columns)):
        if func_novo < len(func_novos):  # Verifica se func_novo não excede o número de elementos em func_novos
            for col_i in range(len(row)):
                row[col_i].value = func_novos[func_novo][col_i]
            
        func_novo += 1
 """
""" #colunas não obrigatorias
colum_no_obligate = [
func_df.columns.get_loc("NUMERO"),
func_df.columns.get_loc("COMPLEMENTO"),
func_df.columns.get_loc("FONE")
]

#Amarelo se as tabelas estão preenchidas corretamente
yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

#Vermelha se estiver faltando alguma informação OBRIGATORIA na coluna
red = PatternFill(start_color="DE4740", end_color="DE4740", fill_type="solid")

#VERDE se for um funcionario for adicionado a planilha original
green = PatternFill(start_color="40DE47", end_color="40DE47", fill_type="solid")

# Iterar sobre as células alteradas e aplicar o estilo de preenchimento amarelo nelas
for row in ws.iter_rows(min_row=2, max_row= quant_func + 1, min_col=1, max_col=len(func_df.columns)):
    #Checa se a linha foi alterada
    if row[func_df.columns.get_loc("SITUACAO")].value == "ALTERADO":
        #Checa se as colunas obrigatorias estão preenchidas
        if ft.check_obligate_collum(row,colum_no_obligate):
            ft.paintRow(row,yellow)
        else:
            ft.paintRow(row,red)
            
        
    elif row[func_df.columns.get_loc("SITUACAO")].value == "ADICIONADO":
            ft.paintRow(row,green)
        
# Salvar as alterações no arquivo Excel mesclado
workbook.save(func_mesclado_path) 
 """


print("*-"*30)
print("Programa finalizado (pressione Enter)....")
print("*-"*30)
input()