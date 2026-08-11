"""
Exceções customizadas do agente RPA.

Classes de erro específicas para diferentes contextos:
- SemAcumuladoresException: Empresa sem acumuladores
- ParadaManualException: Usuário pressionou kill switch
- FalhaVisualizacaoException: OCR/Imagem falhou
"""


class SemAcumuladoresException(Exception):
    """Disparado de propósito quando a empresa não possui nenhum acumulador cadastrado"""
    pass


class ParadaManualException(Exception):
    """Disparado quando o arquivo PARAR.txt é detectado pelo Kill Switch"""
    pass


class FalhaVisualizacaoException(Exception):
    """Disparado quando o OCR ou a busca por imagem não encontra a tela no tempo limite"""
    pass


class NovaVigenciaException(Exception):
    """Disparado quando o Domínio exige a criação de nova vigência antes da inativação."""
    pass
