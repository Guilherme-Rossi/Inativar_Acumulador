import sys
from pywinauto import Desktop

arquivo_log = "raiox_dominio.txt"
# Redireciona tudo que seria impresso na tela para o arquivo de texto
sys.stdout = open(arquivo_log, "w", encoding="utf-8")

print("Iniciando o Raio-X da tela...")

try:
    app = Desktop(backend="uia")
    
    # Pega TODAS as janelas que tenham 'Domínio' ou 'Acumuladores' no título
    janelas = app.windows(title_re=".*(Domínio|Acumuladores).*")
    
    if not janelas:
        print("Nenhuma janela encontrada com esses nomes.")
    else:
        print(f"Foram encontradas {len(janelas)} janela(s)! Fazendo o Raio-X de todas...\n")
        
        for i, janela in enumerate(janelas):
            print(f"=======================================================")
            print(f"JANELA {i+1} - Título: '{janela.window_text()}'")
            print(f"=======================================================\n")
            try:
                # Faz o Raio-X dessa janela específica
                janela.print_control_identifiers()
            except Exception as e_janela:
                print(f"Erro ao mapear a janela {i+1}: {e_janela}")
            print("\n\n")
            
    print("Raio-X finalizado com sucesso!")

except Exception as e:
    print(f"Erro ao tentar mapear a tela: {e}")

finally:
    sys.stdout.close()
    sys.stdout = sys.__stdout__
    print(f"Pronto! Abra o arquivo '{arquivo_log}' e veja o que ele encontrou.")