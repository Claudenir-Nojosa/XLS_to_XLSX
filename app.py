import os
import win32com.client as win32

# Função para renomear o arquivo com base no nome original
def renomear_arquivo(nome_original, nome_empresa):
    mapeamento = {
        "Rel Entrada": "Relatório Entrada",
        "Rel Saída": "Relatório Saída",
        "Mov Entrada": "Movimentação Entrada",
        "Mov Saída": "Movimentação Saída",
        "Mov CFe": "Movimentação CFe"
    }
    # Procura uma chave que esteja no nome original
    for chave, novo_nome in mapeamento.items():
        if chave in nome_original:
            return f"{novo_nome} - {nome_empresa}"
    # Retorna o nome original se nenhuma chave for encontrada
    return f"{nome_original} - {nome_empresa}"

def converter_xls_para_xlsx(pasta_origem, pasta_destino):
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False

    for arquivo in os.listdir(pasta_origem):
        if arquivo.endswith('.xls'):
            caminho_xls = os.path.join(pasta_origem, arquivo)

            try:
                wb = excel.Workbooks.Open(caminho_xls)
                ws = wb.Worksheets(1)  # Seleciona a primeira planilha

                # Remove mesclagem de todas as células
                used_range = ws.UsedRange
                if used_range.MergeCells:
                    used_range.UnMerge()
                
                # Pega o valor da célula A2 e extrai o nome da empresa após "Empresa:"
                valor_a2 = ws.Cells(2, 1).Value
                if valor_a2:
                    nome_empresa = valor_a2.split("Empresa:")[1].strip()  # Pega o valor após "Empresa:" e remove espaços extras
                else:
                    nome_empresa = "Desconhecido"
                
                # Determina o novo nome do arquivo com base no nome original e nome da empresa
                novo_nome = renomear_arquivo(os.path.splitext(arquivo)[0], nome_empresa) + ".xlsx"
                caminho_xlsx = os.path.join(pasta_destino, novo_nome)

                # Salva como .xlsx
                wb.SaveAs(caminho_xlsx, FileFormat=51)  # 51 é o formato para .xlsx
                wb.Close()
                print(f"Convertido: {arquivo} -> {caminho_xlsx}")
            except Exception as e:
                print(f"Erro ao converter {arquivo}: {e}")

    excel.Quit()

# Caminho das pastas
pasta_origem = "C:\\Users\\Claudenir\\Desktop\\Conversor XLS\\XLS's"
pasta_destino = "C:\\Users\\Claudenir\\Desktop\\Conversor XLS\\XLSX's"

converter_xls_para_xlsx(pasta_origem, pasta_destino)
