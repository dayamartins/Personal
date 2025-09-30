Português 🇧🇷

##Descrição

Este script em VBA remove acentos de caracteres em uma planilha e exporta os dados para um arquivo CSV UTF-8.

Os dados são copiados das colunas A a F até a última linha preenchida.
  
O caminho de saída deve estar definido na célula M1 da planilha ativa.

O arquivo é salvo com o nome areas_bloqueio_yyyy-mm-dd.csv dentro da pasta especificada.
  
Caso a pasta não exista, o código a cria automaticamente.

##Como usar

Insira o código no Editor VBA (Alt + F11).

Ajuste a célula M1 com o caminho onde deseja salvar o arquivo.

Execute a macro Botao_Processar_Acentos.

O arquivo CSV será gerado no local indicado.

----------------------------------------------------------------------------------------------------------

English 🇺🇸

##Description

This VBA script removes accents from characters in a worksheet and exports the data to a UTF-8 CSV file.

Data is copied from columns A to F up to the last filled row.

The output path must be set in cell M1 of the active sheet.

The file is saved as blocked_areas_yyyy-mm-dd.csv in the specified folder.

If the folder does not exist, the code will automatically create it.

##How to use

Insert the code into the VBA Editor (Alt + F11).

Set cell M1 with the path where you want the file saved.

Run the macro Button_Process_Accents.

The CSV file will be generated in the given location.
