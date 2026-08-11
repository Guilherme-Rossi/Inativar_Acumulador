"""
Logger centralizado e único para todo o agente RPA.

Singleton que garante que o robô inteiro use o mesmo arquivo de log.
"""

import time
from pathlib import Path


class AgenteLogger:
    """Logger Singleton do agente RPA."""
    
    _instancia = None

    def __new__(cls):
        if cls._instancia is None:
            cls._instancia = super(AgenteLogger, cls).__new__(cls)
            # Por enquanto, deixamos o caminho fixo. Depois o settings.py assume isso.
            pasta_logs = (
                Path.home() 
                / "Desktop" 
                / "PYTHON DOCS" 
                / "Projects Agrelli" 
                / "P1 - Inativar Acumulador" 
                / "Relatorio Final"
            )
            pasta_logs.mkdir(parents=True, exist_ok=True)
            cls._instancia.caminho = pasta_logs / "agente_v2.log"
        return cls._instancia

    def info(self, mensagem: str) -> None:
        """Log em nível INFO."""
        linha = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}"
        print(linha)
        with open(self.caminho, "a", encoding="utf-8") as f:
            f.write(linha + "\n")


# Variável global pronta para ser importada por qualquer outro arquivo!
log = AgenteLogger()

# Função de compatibilidade para código que espera obter um logger
def obter_logger():
    """Retorna a instância global do logger (compatibilidade)."""
    return log
