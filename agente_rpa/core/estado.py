"""
Estado centralizado do agente RPA.

A máquina de estados permite o agente "saber onde está" e tomar decisões
baseado no contexto. Isso transforma um script procedural em um agente inteligente.
"""

from dataclasses import dataclass, field
from typing import Optional, Any
from enum import Enum
from datetime import datetime


class TelaAtual(Enum):
    """Estados possíveis da tela/sistema."""
    TELA_INICIAL = "tela_inicial"
    LISTAGEM_EMPRESAS = "listagem_empresas"
    LISTAGEM_ACUMULADORES = "listagem_acumuladores"
    DETALHE_ACUMULADOR = "detalhe_acumulador"
    MODAL_CONFIRMACAO = "modal_confirmacao"
    MODAL_ERRO = "modal_erro"
    POPUP_DESCONHECIDO = "popup_desconhecido"
    SISTEMA_TRAVADO = "sistema_travado"
    DESCONHECIDO = "desconhecido"


class AcaoAgente(Enum):
    """Ações que o agente pode executar."""
    ABRIR_EMPRESA = "abrir_empresa"
    LISTAR_ACUMULADORES = "listar_acumuladores"
    INATIVAR_ACUMULADOR = "inativar_acumulador"
    VALIDAR_ACUMULADOR = "validar_acumulador"
    MARCAR_ANALISE = "marcar_analise"
    FECHAR_POPUP = "fechar_popup"
    RECUPERAR_ERRO = "recuperar_erro"
    REINICIAR_FLUXO = "reiniciar_fluxo"
    PARAR = "parar"
    AGUARDAR = "aguardar"


@dataclass
class EstadoEmpresa:
    """Estado da processamento de uma empresa."""
    codigo: int
    nome: str = ""
    acumuladores_totais: int = 0
    acumuladores_processados: int = 0
    acumuladores_inativados: int = 0
    acumuladores_com_erro: int = 0
    timestamp_inicio: datetime = field(default_factory=datetime.now)
    timestamp_fim: Optional[datetime] = None
    erros: list[str] = field(default_factory=list)

    @property
    def percentual_processado(self) -> float:
        if self.acumuladores_totais == 0:
            return 0.0
        return (self.acumuladores_processados / self.acumuladores_totais) * 100


@dataclass
class EstadoAcumulador:
    """Estado da processamento de um acumulador."""
    codigo: int
    empresa_codigo: int = 0
    status: str = "pendente"  # pendente, processando, inativado, marcado, erro
    tentativas: int = 0
    erro_ultimo: Optional[str] = None
    timestamp_processamento: Optional[datetime] = None
    detalhes: dict[str, Any] = field(default_factory=dict)


@dataclass
class EstadoAgente:
    """
    Estado centralizado do agente.
    
    Responde perguntas como:
    - Onde estou agora?
    - Qual empresa estou processando?
    - Quantos acumuladores faltam?
    - Qual foi o último erro?
    - Quantas tentativas fizemos?
    """

    # ============================================================
    # ESTADO ATUAL
    # ============================================================
    tela_atual: TelaAtual = TelaAtual.TELA_INICIAL
    ultima_acao: Optional[AcaoAgente] = None
    acao_pronta: Optional[AcaoAgente] = None

    # ============================================================
    # CONTEXTO DE NEGÓCIO
    # ============================================================
    empresa_atual: Optional[EstadoEmpresa] = None
    acumulador_atual: Optional[EstadoAcumulador] = None

    # ============================================================
    # HISTÓRICO DE EXECUÇÃO
    # ============================================================
    empresas_processadas: dict[int, EstadoEmpresa] = field(default_factory=dict)
    timestamp_inicio_agente: datetime = field(default_factory=datetime.now)

    # ============================================================
    # CONTROLE DE ERRO E RECUPERAÇÃO
    # ============================================================
    tentativas_acao_atual: int = 0
    erro_ultimo: Optional[str] = None
    erro_tipo: Optional[str] = None
    recuperavel: bool = True  # Indica se o erro pode ser recuperado

    # ============================================================
    # PARADA E SINALIZAÇÃO
    # ============================================================
    deve_parar: bool = False
    parada_manual: bool = False
    mensagem_parada: str = ""

    def registrar_empresa(self, codigo: int, nome: str) -> None:
        """Registra que começou a processar uma empresa."""
        self.empresa_atual = EstadoEmpresa(codigo=codigo, nome=nome)
        self.tela_atual = TelaAtual.LISTAGEM_EMPRESAS
        self.tentativas_acao_atual = 0
        self.erro_ultimo = None

    def registrar_acumulador(self, codigo: int) -> None:
        """Registra que começou a processar um acumulador."""
        self.acumulador_atual = EstadoAcumulador(
            codigo=codigo,
            empresa_codigo=self.empresa_atual.codigo if self.empresa_atual else 0
        )
        self.tela_atual = TelaAtual.LISTAGEM_ACUMULADORES
        self.tentativas_acao_atual = 0
        self.erro_ultimo = None

    def registrar_acao(self, acao: AcaoAgente) -> None:
        """Registra a ação que vai ser executada."""
        self.ultima_acao = self.acao_pronta
        self.acao_pronta = acao

    def registrar_erro(self, erro: str, tipo: str = "generico", recuperavel: bool = True) -> None:
        """Registra um erro ocorrido."""
        self.erro_ultimo = erro
        self.erro_tipo = tipo
        self.recuperavel = recuperavel
        self.tentativas_acao_atual += 1

        if self.acumulador_atual:
            self.acumulador_atual.erro_ultimo = erro
            self.acumulador_atual.tentativas += 1

    def registrar_sucesso_acumulador(self, tipo_sucesso: str = "inativado") -> None:
        """Registra sucesso na processamento de um acumulador."""
        if not self.acumulador_atual:
            return

        self.acumulador_atual.status = tipo_sucesso
        self.acumulador_atual.timestamp_processamento = datetime.now()

        if self.empresa_atual:
            self.empresa_atual.acumuladores_processados += 1

            if tipo_sucesso == "inativado":
                self.empresa_atual.acumuladores_inativados += 1

        self.tentativas_acao_atual = 0
        self.erro_ultimo = None

    def registrar_sucesso_empresa(self) -> None:
        """Registra fim de processamento de empresa."""
        if not self.empresa_atual:
            return

        self.empresa_atual.timestamp_fim = datetime.now()
        self.empresas_processadas[self.empresa_atual.codigo] = self.empresa_atual
        self.empresa_atual = None

    def mudar_tela(self, nova_tela: TelaAtual) -> None:
        """Altera a tela atual (com log)."""
        self.tela_atual = nova_tela

    def resumo_execucao(self) -> dict:
        """Retorna um resumo do progresso até agora."""
        return {
            "empresas_totais": len(self.empresas_processadas),
            "acumuladores_inativados": sum(
                e.acumuladores_inativados for e in self.empresas_processadas.values()
            ),
            "acumuladores_com_erro": sum(
                e.acumuladores_com_erro for e in self.empresas_processadas.values()
            ),
            "tempo_decorrido_minutos": (
                datetime.now() - self.timestamp_inicio_agente
            ).total_seconds() / 60,
            "ultima_empresa": (
                self.empresa_atual.codigo if self.empresa_atual else None
            ),
            "tela_atual": self.tela_atual.value,
            "erro_ultimo": self.erro_ultimo,
        }

    def __repr__(self) -> str:
        return (
            f"EstadoAgente("
            f"tela={self.tela_atual.value}, "
            f"empresa={self.empresa_atual.codigo if self.empresa_atual else 'None'}, "
            f"acumulador={self.acumulador_atual.codigo if self.acumulador_atual else 'None'}, "
            f"tentativas={self.tentativas_acao_atual}, "
            f"erro={self.erro_tipo}"
            f")"
        )
