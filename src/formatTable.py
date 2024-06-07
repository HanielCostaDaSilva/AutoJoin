from unidecode import unidecode
from openpyxl.styles import PatternFill
def format_cell(cell_value:str)-> str:
    cell_value = unidecode(cell_value).upper()
    return cell_value
def check_obligate_collum(row, no_obligate_index: list[int] = [] )->bool:
    '''
    método que verifica se uma linha está devidamente preenchida, caso 
    '''   
    check =True
    
    for i in range(len(row)):
 
        if i in no_obligate_index:
            continue
        
        if row[i].value ==None:
            check = not check
            break
    
    return check
def paintRow(row, collor:PatternFill):
    #pinta todas as colunas de uma determinada linha 
    for cell in row:
        cell.fill = collor