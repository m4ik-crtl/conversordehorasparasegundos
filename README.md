# **Conversor de Horas para Segundos via Python**

Este projeto fornece uma solução para converter valores de tempo no formato **HH:MM:SS** para segundos em um arquivo Excel. O código Python usa a biblioteca **openpyxl** para ler, processar e escrever os dados no arquivo Excel.

## **Pré-requisitos**

Antes de começar, você precisará ter o Python instalado em sua máquina e a biblioteca **openpyxl**.

## **Passos para Usar**

1. **Preparar os Arquivos**
   - Coloque o arquivo Excel e o script Python na mesma pasta. 
   - O arquivo Excel deve conter os valores de horas no formato **HH:MM:SS** nas colunas especificadas no código (por exemplo, **E até P**).

2. **Instalar a Biblioteca openpyxl**
   - Abra o terminal (pressione **Win + R**, digite `cmd` e pressione Enter).
   - Navegue até a pasta onde os arquivos estão localizados. Exemplo:
     ```bash
     cd C:/Users/SeuUsuario/Downloads/Pasta1
     ```
   - Instale a biblioteca **openpyxl**:
     ```bash
     pip install openpyxl
     ```

3. **Executar o Script Python**
   - Execute o comando abaixo para rodar o script e processar o arquivo Excel:
     ```bash
     python converter_excel.py
     ```

4. **Resultado Esperado**
   - Após a execução do script, o arquivo Excel será atualizado com os valores convertidos para segundos nas células correspondentes.

## **Script Python**

```python
import openpyxl

def convert_time_to_seconds(time_str):
    """Converte o formato HH:MM:SS para segundos."""
    hours, minutes, seconds = map(int, time_str.split(':'))
    return hours * 3600 + minutes * 60 + seconds

def process_excel(file_path, input_start_column, input_end_column, sheet_name="Sheet1"):
    """Lê valores em HH:MM:SS de várias colunas e escreve os segundos correspondentes em outras células."""
    # Carregar o arquivo Excel
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]

    # Iterar pelas colunas de E até P (ou o intervalo de colunas desejado)
    for row in range(2, sheet.max_row + 1):  # Ignorar o cabeçalho
        for col in range(openpyxl.utils.column_index_from_string(input_start_column),
                         openpyxl.utils.column_index_from_string(input_end_column) + 1):
            cell_value = sheet.cell(row=row, column=col).value
            
            # Verificar se a célula contém uma string no formato HH:MM:SS
            if isinstance(cell_value, str) and ":" in cell_value:
                try:
                    total_seconds = convert_time_to_seconds(cell_value)
                    sheet.cell(row=row, column=col).value = total_seconds
                except ValueError:
                    print(f"Erro ao processar valor na célula {row},{col}: {cell_value}")

    # Salvar o arquivo Excel
    wb.save(file_path)
    print(f"Arquivo processado e salvo em: {file_path}")

# Exemplo de uso
file_path = "tempo.xlsx"  # Caminho para o arquivo Excel
input_start_column = "E"  # Primeira coluna (E)
input_end_column = "P"    # Última coluna (P)
process_excel(file_path, input_start_column, input_end_column)
