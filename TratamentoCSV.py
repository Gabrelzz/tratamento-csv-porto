import pandas as pd
import win32com.client
import os
import time
import sys

''' Script para renomear '''

def rename():  
    renamed_files = {
        "Arquivo1": "[1] - Arquivo1",
        "Arquivo2": "[2] - Arquivo2",
        "Arquivo3": "[3] - Arquivo3",
        "Arquivo4": "[4] - Arquivo4",
        "Arquivo5": "[5] - Arquivo5",
        "Arquivo6": "[6] - Arquivo6",
        "Arquivo7": "[7] - Arquivo7",
        "Arquivo8": "[8] - Arquivo8",
        "Arquivo9": "[9] - Arquivo9",
        "Arquivo10": "[10] - Arquivo10",
    }

    for file in os.listdir():
        for prefixo, new_name in renamed_files.items():
                if file.startswith(prefixo) and file.endswith('.csv'):
                    os.rename(file, new_name)
                    file = new_name

# Definir diretório para o mesmo do script
csv_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
print(f"Diretório do Script: {csv_directory}")

# Listar todos os arquivos CSV no diretório
csv_files = [f for f in os.listdir(csv_directory) if f.endswith('.csv')]

if not csv_files:
    print("Nenhum arquivo CSV encontrado.")
else:
    # Iniciar Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False

    for file in csv_files:
        file_path = os.path.join(csv_directory, file)

        if not os.path.exists(file_path):
            print ("Arquivo não foi encontrado")
            continue
        try:
            # Abrindo pasta de trabalho
            workbook = excel.Workbooks.Open(file_path)
            # Executar macro
            macro_name = "PERSONAL.XLSB!separarDados"
            time.sleep(0.3)
            try:
                excel.Application.Run(macro_name)
                print(f"Macro {macro_name} executada.")
            except Exception as macro_error:
                print(f"Erro: {macro_error}")
        except Exception as e:
            print("Erro ao abrir ou processar macro")

        workbook.SaveAs(file_path, FileFormat=6)
        workbook.Close(SaveChanges=True)
        print("Macro executada e arquivo salvo")
    try:
        excel.Quit()
    except Exception as quit_error:
        print(f"Erro ao fechar excel: {quit_error}")

''' Renomear arquivos '''
try:
    time.sleep(1)
    rename()
    print("Arquivos renomeados.")
except Exception as name_error:
    print("Arquivo não foi renomeado")
