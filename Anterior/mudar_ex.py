import os
import time
from pathlib import Path
import win32com.client as win32

def converter_arquivos_em_lote_com_excel():
    # Caminho da pasta onde estão os relatórios do Domínio
    pasta_relatorios = Path(r"C:\Users\guilherme.rossi\Desktop\PYTHON DOCS\Projects Agrelli\P1 - Inativar Acumulador\Relação Por empresa")
    
    if not pasta_relatorios.exists():
        print(f"[ERRO] A pasta não foi encontrada: {pasta_relatorios}")
        return

    # Busca todos os arquivos que terminam com .xls
    arquivos_xls = list(pasta_relatorios.glob("*.xls"))
    
    if not arquivos_xls:
        print("[INFO] Nenhum arquivo .xls encontrado na pasta.")
        return

    print(f"=== Iniciando conversão de {len(arquivos_xls)} arquivos via Microsoft Excel ===")
    
    sucesso = 0
    erros = 0

    # Inicia o Excel em segundo plano (invisível)
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False # MÁGICA: Ignora os avisos de "Arquivo corrompido/Formato diferente" do Excel
    except Exception as e:
        print(f"[ERRO FATAL] Não foi possível iniciar o Excel em background. Erro: {e}")
        return

    for arquivo in arquivos_xls:
        destino = arquivo.with_suffix(".xlsx")
        
        # Verifica se o arquivo .xlsx já foi criado antes
        if destino.exists():
            print(f"[PULANDO] O arquivo '{destino.name}' já existe.")
            continue
            
        try:
            print(f"Convertendo: '{arquivo.name}'...", end=" ")
            
            # O Excel pelo Python (COM Object) exige caminhos absolutos (como string)
            caminho_abs_xls = str(arquivo.resolve())
            caminho_abs_xlsx = str(destino.resolve())
            
            # 1. Abre o arquivo disfarçado
            wb = excel.Workbooks.Open(caminho_abs_xls)
            
            # 2. Salva como o Excel moderno verdadeiro (FileFormat=51 é a extensão .xlsx)
            wb.SaveAs(caminho_abs_xlsx, FileFormat=51)
            wb.Close()
            
            print("OK!")
            sucesso += 1
            
        except Exception as e:
            print(f"ERRO! -> {e}")
            erros += 1

    # Fecha o Excel fantasma
    excel.Quit()

    print("\n=== RESUMO DA CONVERSÃO ===")
    print(f"Convertidos com sucesso: {sucesso}")
    print(f"Erros: {erros}")
    print("===========================")

if __name__ == "__main__":
    converter_arquivos_em_lote_com_excel()