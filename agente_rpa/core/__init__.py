"""Módulo core do agente RPA (Logger, Estado, Exceções)."""

from .logger import log, obter_logger
from .excecoes import SemAcumuladoresException, ParadaManualException, FalhaVisualizacaoException

# Se tiver `estado.py` no core, descomente a importação abaixo
# from .estado import EstadoAgente

__all__ = [
    "log",
    "obter_logger",
    "SemAcumuladoresException",
    "ParadaManualException",
    "FalhaVisualizacaoException",
]
