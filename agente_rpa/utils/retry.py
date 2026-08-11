"""
Motor de "tentar de novo" isoladinho.

Executa uma função repetidas vezes caso dê erro.
Erros de negócio (sem acumulador, parada manual) não fazem retry.
"""

import time
from typing import Callable, Any
from agente_rpa.core.logger import log
from agente_rpa.core.excecoes import (
    SemAcumuladoresException,
    ParadaManualException,
    NovaVigenciaException,
)


def executar_com_retry(
    func: Callable,
    tentativas: int = 3,
    espera: float = 2.0,
    descricao: str = "Operação"
) -> Any:
    """
    Executa uma função repetidas vezes caso dê erro.
    Ignora a repetição se for um erro de Negócio (ex: Sem Acumulador ou Parada Manual).
    
    Args:
        func: Função a executar
        tentativas: Número máximo de tentativas
        espera: Segundos entre tentativas
        descricao: Descrição para logs
        
    Returns:
        Resultado da função
        
    Raises:
        SemAcumuladoresException: Se empresa sem acumuladores (não faz retry)
        ParadaManualException: Se usuário pediu parada (não faz retry)
        Exception: Última exceção após esgotar tentativas
    """
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            return func()
        except (SemAcumuladoresException, ParadaManualException, NovaVigenciaException):
            # Erros de fluxo conhecido: Sobe direto, não faz retry!
            raise
        except Exception as exc:
            ultimo_erro = exc
            log.info(f"[RETRY] {descricao} falhou na tentativa {tentativa}/{tentativas}: {exc}")
            if tentativa < tentativas:
                time.sleep(espera)

    log.info(f"[ERRO FATAL] Esgotadas as tentativas para: {descricao}")
    raise ultimo_erro
