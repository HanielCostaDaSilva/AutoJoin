## Descrição:

O projeto AutoJoin foi desenvolvido para automatizar o processo de junção e atualização de dados de funcionários armazenados em planilhas do Excel. O objetivo principal é mesclar os dados de uma planilha principal com os dados atualizados de outra planilha, garantindo a consistência e atualização dos registros.

## Funcionalidades:
<ol>
    <li> 
    Mesclagem de Dados: O programa mescla os dados de funcionários de uma planilha principal com os dados atualizados de outra planilha, com base em uma chave única (como a matrícula do funcionário).
    </li>
    <li> 
    Atualização de Registros: Os registros existentes na planilha principal são atualizados com os dados mais recentes da planilha de atualização. Os campos em branco na planilha principal são preenchidos com os dados correspondentes da planilha de atualização.
    </li>
    <li> 
    Identificação de Novos Funcionários: Funcionários que estão presentes na planilha de atualização, mas não na planilha principal, são identificados e adicionados à planilha principal.
    </li>
    <li>
    Marcação de Alterações: As alterações realizadas nos registros existentes são marcadas através de cores determinadas pela empresa, permitindo uma fácil visualização das mudanças feitas.
    </li>
    <li>
    Validação de Dados: O programa verifica se os campos obrigatórios estão preenchidos corretamente e destaca as linhas com informações ausentes.
    </li>
</ol>

## Como Usar:
### 1.  Configuração do Ambiente:
1. Certifique-se de ter o [Python](https://www.python.org/downloads/) instalado em seu ambiente.

2. **Após a instalação do python**. Instale as bibliotecas Python necessárias.

<br>
No terminal escreva os seguintes comandos
```
pip install pandas
pip install openpyxls
```

### 2.  Preparação dos Arquivos:

1. Coloque as planilhas na pasta `data`. 
2. nomeie a planilha com os dados originais com o seguinte nome: `func.xlsx` 
3. nomeie a planilha com os dados que serão cadastrados com o seguinte nome: `func_atualizado.xlsx`.
<br>
*Obs: Você pode mudar o nome das planilhas, através de **hardcode***

### 3. **Execução do Programa:**
   1. Execute o script `App.py` para iniciar o processo de junção e atualização dos dados.
   2. O programa mesclará os dados, atualizará registros existentes e identificará novos funcionários.
   3. O arquivo resultante `func_final.xlsx` será gerado na pasta `data`, contendo os dados atualizados.

### 4. **Análise dos Resultados:**
   - Verifique o arquivo `func_final.xlsx` para revisar os dados mesclados e atualizados.
   - As alterações e adições de funcionários serão destacadas para fácil identificação.

## Considerações:
O projeto AutoJoin oferece uma solução simples, eficiente e automatizada para gerenciar e atualizar dados de funcionários, reduzindo a necessidade de intervenção manual e garantindo a precisão e integridade dos registros.

--- 