"""
Agente RPA de Inativação de Acumuladores.

Automação inteligente para sistemas legados corporativos.
"""

from agente_rpa.config import obter_config, cfg
from agente_rpa.core import obter_logger, log

__version__ = "2.0.0"
__all__ = [
    "obter_config",
    "cfg",
    "obter_logger",
    "log",
]
