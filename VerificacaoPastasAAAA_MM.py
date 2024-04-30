import os
import pandas as pd

diretorio_principal = r"\\10.1.1.2\FOLHA\GRUPOS"

diretorios_encontrados = []

for raiz, diretorios, arquivos in os.walk(diretorio_principal):
    for diretorio in diretorios:
        if "AAAA_MM" in diretorio:
            diretorios_encontrados.append(os.path.join(raiz, diretorio))

# Criar um DataFrame com os diretórios encontrados
df = pd.DataFrame({"Diretório com AAAA_MM": diretorios_encontrados})

# Salvar o DataFrame em um arquivo Excel
excel_writer = pd.ExcelWriter("diretorios_com_AAAA_MM.xlsx", engine="xlsxwriter")
df.to_excel(excel_writer, sheet_name="Diretórios", index=False)
excel_writer.save()

print("Diretórios salvos em 'diretorios_com_AAAA_MM.xlsx'")
