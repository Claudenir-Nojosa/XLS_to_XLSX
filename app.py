import os
import win32com.client as win32

def converter_xls_para_xlsx(pasta_origem, pasta_destino):
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False

    for arquivo in os.listdir(pasta_origem):
        if arquivo.endswith('.xls'):
            caminho_xls = os.path.join(pasta_origem, arquivo)
            caminho_xlsx = os.path.join(pasta_destino, f"{os.path.splitext(arquivo)[0]}_convertido.xlsx")

            try:
                wb = excel.Workbooks.Open(caminho_xls)
                wb.SaveAs(caminho_xlsx, FileFormat=51)  # 51 é o formato para .xlsx
                wb.Close()
                print(f"Convertido: {arquivo} -> {caminho_xlsx}")
            except Exception as e:
                print(f"Erro ao converter {arquivo}: {e}")

    excel.Quit()

# Caminho das pastas
pasta_origem = "C:\\Users\\Claudenir\\Desktop\\Conversor XLS\\XLS's"
pasta_destino = "C:\\Users\\Claudenir\\Desktop\\Conversor XLS\\XLSX_s"

converter_xls_para_xlsx(pasta_origem, pasta_destino)
