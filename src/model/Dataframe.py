import pandas as pd
import os

class Dataframe:
    __path=""
    __df:pd.DataFrame = None
    
    __naDefault =""
        
    def __init__(self, path:str, naDefault =""):
        self.__path = path
        self.__naDefault = naDefault
        
        if os.path.exists(path):
            self.get_excel_df()    
        else:
            self.__create_df()        
    
    def get_excel_df(self):
        self.__df = pd.read_excel(self.__path, dtype=str).fillna(self.__naDefault)
        
    def __create_df(self):
        '''Create a new dataframe'''
        self.__df = pd.DataFrame()

        # Criamos a planilha na pasta data
        self.df.to_excel(self.__path, index=False)
    
    def save(self) -> bool:
        '''
        return True if the setores was been saved, Exception if not
        '''
        self.__df.to_excel(self.__path, index=False)
        return True
    
    
    def get_rows_by_collumn(self, column:str, value:str):
        '''
        procura no dataframe, todas as que possuem um determinado valor 
        se não for possível, lascou pq eu não fiz tratamento :P
        '''
    
        return self.__df.loc[self.__df[column] == value]
    
    def add_row(self, **object: object):
        '''
        recebe um objeto e o transforma em uma linha para o dataframe 
        '''
        new_row_df = pd.DataFrame([object]).fillna(self.__naDefault)
        
        self.__df = pd.concat([self.__df, new_row_df], ignore_index=True).fillna(self.__naDefault)
        
    
    def delete_row(self,index:int):
        '''
        apaga uma determinada linha do dataframe, caso não seja possível encontrar o índex, retorna None
        '''
        if self.__check_index_(index):
            row= self.__df.loc[index]
            
            self.__df = self.__df.drop(index=index)
            return row
        
        else:
            return None
     
            
    def alter_row(self, index:int, column:str, value:str):
        '''
            altera o valor de uma linha, através do seu índice. 
            caso a `column` não exista no df, criará a linha
        '''
        if column not in self.df.columns:
            self.add_column(column) 
        
        self.__df.loc[index, column] = value 
      
    def fill_collumn(self, col_name:str,value=""):
        '''
        Preenche todos os valores em brancos de uma coluna, 
        return `Bool` caso exista ou não essa coluna
        '''
        if col_name  in self.__df.columns:
            self.__df[col_name] = self.__df[col_name].replace("", value).fillna(value)
            return True
        
        return False
    
    
    def add_column(self,col_name:str,default_value="")-> bool:
        '''
        adicona uma nova coluna no dataframe, se ela não existir no dataframe. 
        return True se foi possível adicionar a nova coluna
        '''
        if col_name not in self.__df.columns:
            self.__df[col_name] = default_value
            return True
        
        return False
    
    def delete_column(self,col_name:str)->bool:
        ''' 
        remove uma determinada coluna do dataframe
        caso ela exista, retorna True
        '''
        if self.__check__column(col_name):
            self.__df = self.__df.drop(columns=[col_name])
            return True
        
        return False
    
    def change_value(self, column:str,old_value:str,new_value):
        '''
        Troca todos os valores de uma coluna por um outro valor
        \nretorna `Bool` se foi possível alterar ou não
        '''
        if column in self.df.columns:
            self.df[column] = self.df[column].replace(old_value, new_value)
            return True
        
        return False
    
    
    def __check_index_(self, index:int)-> bool:
        ''' 
        valida se o index existe dentro do dataframe
        '''
        return index > 0 and index < len(self.__df)  
    
    
    def __check__column(self,col_name:str)-> bool:
        ''' 
        valida se a coluna existe dentro do dataframe
        '''
        return col_name in self.__df.columns
    
    def get_iterrows(self):
        '''
        gera um iterador para o dataframe
        '''
        return self.__df.iterrows()
    
    @property
    def path(self):
        return self.__path

    @path.setter
    def path(self, new_value):
        self.__path = new_value
        
    @property
    def df(self):
        return self.__df.copy()
    
    def __str__(self) -> str:                               
        return str(self.__df)
    
    @property
    def columns(self):
        return self.__df.columns
    
if __name__ =="__main__":
    
    actual_directory = os.path.dirname(os.path.realpath(__file__))
    data_path = os.path.join(actual_directory, "..","..", "data", "func.xlsx")
    
    print("=========ORIGINAL===========")
    df = Dataframe(data_path)
    print(df)
    print("=========ORIGINAL===========")
    print("=========ADICIONANDO NOVA LINHA===========")
    
    func = {
          "MATRICULA":"100",
          "NOME":"JOAO LUCAS",
          "SETOR":"PESQUISA",
          "CPF":"345464432",
    }
    df.add_row(**func)
    
    print(df)
    print("=========ADICIONANDO NOVA LINHA===========")