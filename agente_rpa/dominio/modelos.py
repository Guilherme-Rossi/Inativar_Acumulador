"""
Modelos de Domínio - Entidades Reais do Negócio

Aqui abandonamos dicionários soltos e criamos estruturas sólidas.
O Python (e VS Code) agora sabem exatamente o que é uma "Empresa" e um "Acumulador".
Ativa autocomplete, blinda dados, e torna o código legível.
"""

from dataclasses import dataclass, field
from typing import List, Set


@dataclass
class Acumulador:
    """
    Representa a regra de negócio de um acumulador vindo da planilha Mestre.
    
    Atributos:
        codigo: ID único do acumulador
        deve_inativar: Se deve ser inativado (marcado amarelo, faixa válida)
        deve_marcar_analise: Se deve entrar em fila de análise
    """
    codigo: int
    deve_inativar: bool = False
    deve_marcar_analise: bool = False


@dataclass
class Empresa:
    """
    Representa o cliente e guarda a memória do que foi lido no Domínio.
    
    Atributos:
        codigo: CNPJ ou ID da empresa
        nome: Razão social
        acumuladores_encontrados: Códigos lidos no relatório (para auditoria)
        linhas_processadas: LINHAS processadas no Log (para Resume/Retomada)
    """
    codigo: int
    nome: str
    
    # Onde o robô vai guardar os códigos lidos no relatório do Excel
    acumuladores_encontrados: List[int] = field(default_factory=list)
    
    # Onde o robô vai guardar as LINHAS que ele leu no Log (para Retomada/Resume)
    linhas_processadas: Set[int] = field(default_factory=set)
