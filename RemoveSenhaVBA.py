import zipfile
import os
import shutil
import openpyxl
from tkinter import Tk, filedialog
from openpyxl.styles import Protection

def remover_protecao_vba_xlsm(caminho_arquivo):
    backup_path = caminho_arquivo.replace('.xlsm', '_backup.xlsm')
    shutil.copy2(caminho_arquivo, backup_path)

    with zipfile.ZipFile(backup_path, 'r') as zip_in:
        with zipfile.ZipFile(caminho_arquivo, 'w') as zip_out:
            for item in zip_in.infolist():
                data = zip_in.read(item.filename)
                if item.filename == 'xl/vbaProject.bin':
                    print("[✔] Encontrado vbaProject.bin — removendo proteção VBA...")
                    data = data.replace(b'DPB=', b'DPx=')
                zip_out.writestr(item, data)

    print(f"[✔] VBA desbloqueado. Backup criado: {backup_path}")

def remover_protecao_planilhas_xlsx(caminho_arquivo):
    backup_path = caminho_arquivo.replace('.xlsx', '_backup.xlsx')
    shutil.copy2(caminho_arquivo, backup_path)

    wb = openpyxl.load_workbook(caminho_arquivo)
    for sheet in wb.worksheets:
        if sheet.protection.sheet:
            sheet.protection.sheet = False
            print(f"[✔] Proteção de planilha removida: {sheet.title}")
        for row in sheet.iter_rows():
            for cell in row:
                cell.protection = Protection(locked=False)

    novo_arquivo = caminho_arquivo.replace('.xlsx', '_desprotegido.xlsx')
    wb.save(novo_arquivo)
    print(f"[✔] Células desbloqueadas. Salvo como: {novo_arquivo}")
    print(f"[ℹ] Backup criado: {backup_path}")

def escolher_e_processar_arquivo():
    root = Tk()
    root.withdraw()

    caminho = filedialog.askopenfilename(
        title="Selecione um arquivo Excel (.xlsx ou .xlsm)",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm")]
    )

    if not caminho:
        print("Nenhum arquivo selecionado.")
        return

    print(f"[✔] Arquivo selecionado: {caminho}")

    if caminho.endswith(".xlsm"):
        remover_protecao_vba_xlsm(caminho)
    elif caminho.endswith(".xlsx"):
        remover_protecao_planilhas_xlsx(caminho)
    else:
        print("[✖] Tipo de arquivo não suportado.")

if __name__ == "__main__":
    escolher_e_processar_arquivo()
