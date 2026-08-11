"""
Exceções customizadas do agente RPA.

Permitem tratamento específico e mensagens de erro estruturadas.
"""


class ErroAgenteRPA(Exception):
    """Exceção base para todos os erros do agente."""

    def __init__(self, mensagem: str, codigo_erro: str = "ERRO_GENERICO"):
        self.mensagem = mensagem
        self.codigo_erro = codigo_erro
        super().__init__(f"[{codigo_erro}] {mensagem}")


class ErroAutomacao(ErroAgenteRPA):
    """Erro na camada de automação (cliques, teclado, imagens)."""

    def __init__(self, mensagem: str):
        super().__init__(mensagem, "ERRO_AUTOMACAO")


class ErroVisao(ErroAutomacao):
    """Erro em OCR ou template matching."""

    def __init__(self, mensagem: str):
        super().__init__(f"[Visão] {mensagem}")


class ErroTemplateMismatch(ErroVisao):
    """Imagem/Template não encontrado na tela."""

    def __init__(self, nome_imagem: str, confianca_minima: float = 0.80):
        super().__init__(
            f"Template '{nome_imagem}' não encontrado (confiança mínima: {confianca_minima})"
        )


class ErroOCR(ErroVisao):
    """Erro ao executar OCR."""

    def __init__(self, mensagem: str):
        super().__init__(f"Falha no OCR: {mensagem}")


class ErroNegocios(ErroAgenteRPA):
    """Erro na camada de lógica de negócio."""

    def __init__(self, mensagem: str):
        super().__init__(mensagem, "ERRO_NEGOCIOS")


class SemAcumuladoresException(ErroNegocios):
    """Empresa não possui acumuladores para processar."""

    def __init__(self, codigo_empresa: int):
        super().__init__(f"Empresa {codigo_empresa} sem acumuladores")


class ErroInativacao(ErroNegocios):
    """Erro ao inativar acumulador."""

    def __init__(self, codigo_acumulador: int, motivo: str):
        super().__init__(
            f"Falha ao inativar acumulador {codigo_acumulador}: {motivo}"
        )


class ErroValidacao(ErroNegocios):
    """Erro na validação de dados."""

    def __init__(self, mensagem: str):
        super().__init__(f"Validação falhou: {mensagem}")


class ParadaManualException(ErroAgenteRPA):
    """Execução interrompida manualmente pelo usuário."""

    def __init__(self, arquivo_parada: str = ""):
        msg = f"Execução interrompida manualmente"
        if arquivo_parada:
            msg += f" ({arquivo_parada})"
        super().__init__(msg, "PARADA_MANUAL")


class TimeoutAgenteException(ErroAgenteRPA):
    """Timeout na execução de uma ação."""

    def __init__(self, acao: str, timeout_segundos: float):
        super().__init__(
            f"Timeout na ação '{acao}' após {timeout_segundos}s",
            "TIMEOUT"
        )


class ErroConfiguração(ErroAgenteRPA):
    """Erro na configuração do agente."""

    def __init__(self, mensagem: str):
        super().__init__(f"Erro de configuração: {mensagem}", "ERRO_CONFIG")
