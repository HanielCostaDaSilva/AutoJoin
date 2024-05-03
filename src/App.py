import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from unidecode import unidecode

import formatTable as ft

func_path ="./data/func.xlsx"
func_atualizado_path ="./data/func_atualizado.xlsx"

func_mesclado_path ="./data/func_mescla.xlsx"

func_df = pd.read_excel(func_path,dtype=str).fillna("")

# Remover acentos de todas as colunas do DataFrame
for coluna in func_df.columns:
    func_df[coluna] = func_df[coluna].apply(lambda x: unidecode(str(x)) if pd.notnull(x) else x)

func_atualizado_df = pd.read_excel(func_atualizado_path,dtype=str).fillna("")

# Remover acentos de todas as colunas do DataFrame
for coluna in func_atualizado_df.columns:
    func_atualizado_df[coluna] = func_atualizado_df[coluna].apply(lambda x: unidecode(str(x)) if pd.notnull(x) else x)


#Percorremos o dataframe dos funcionarios atualizados
for index, func_atualizado in func_atualizado_df.iterrows():
        
    # Encontrar o funcionário correspondente no DataFrame original
    funcionario_index = func_df.index[func_df['MATRICULA'] == func_atualizado['MATRICULA']]
    
    # Se o funcionário estiver presente no DataFrame original, atualizar os valores em branco
    if len(funcionario_index) > 0:
        funcionario_index = funcionario_index[0]  # Pegar o índice do primeiro funcionário encontrado
        func_df.at[funcionario_index,"SITUACAO"] = "ALTERADO"
    
    # Iterar sobre as colunas do DataFrame original
        for coluna in func_df.columns:
            # Se o valor da coluna no DataFrame original estiver em branco, preencher com o valor correspondente do DataFrame atualizado
            if func_df.at[funcionario_index, coluna] == "":
                func_df.at[funcionario_index, coluna] = unidecode(func_atualizado[coluna]).upper()
                
    #Caso não encontre o funcionario no dataframe, adicione-o no dataframe original
    else:
        print(f"ESTE FUNCIONARIO ESTAVA PRESENTE NOS ATUALIZADOS MAIS NÃO NA ORIGINAL: \n{func_atualizado}")
        func_df= func_df.add(func_atualizado)
        
# Salvar o DataFrame mesclado
func_df.to_excel(func_mesclado_path, index=False)

# Carregar o arquivo Excel mesclado
workbook = load_workbook(func_mesclado_path)
ws = workbook.active

#colunas não obrigatorias
colum_no_obligate = [
func_df.columns.get_loc("NUMERO"),
func_df.columns.get_loc("COMPLEMENTO"),
func_df.columns.get_loc("FONE")
]

#Amarelo se as tabelas estão preenchidas corretamente
yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

#Vermelha se estiver faltando alguma informação OBRIGATORIA na coluna
red = PatternFill(start_color="DE4740", end_color="DE4740", fill_type="solid")

# Iterar sobre as células alteradas e aplicar o estilo de preenchimento amarelo nelas
for row in ws.iter_rows(min_row=2, max_row=len(func_df) + 1, min_col=1, max_col=len(func_df.columns)):
    #Checa se a linha foi alterada
    if row[func_df.columns.get_loc("SITUACAO") ].value == "ALTERADO":
        
        #Checa se as colunas obrigatorias estão preenchidas
        if ft.check_obligate_collum(row,colum_no_obligate):
            ft.paintRow(row,yellow)
        else:
            ft.paintRow(row,red)
            
# Salvar as alterações no arquivo Excel mesclado
workbook.save(func_mesclado_path) 