"""
Configuração centralizada e única para todo o agente RPA.

Centraliza coordenadas do monitor Dell, tempos de espera e regras de negócio.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class AgenteRPASettings:
    """Configuração unificada do Agente RPA de Inativação de Acumuladores."""

    # ============================================================
    # CAMINHOS
    # ============================================================
    base_dir: Path = Path(__file__).resolve().parents[2]

    planilha_empresas: Path = base_dir / "EMPRESA PARA INATIVAR ACUMULADOR.xlsx"
    planilha_acumuladores: Path = base_dir / "RELAÇÃO DE ACUMULADORES.xlsx"

    pasta_relatorios: Path = base_dir / "Relação Por empresa"
    pasta_logs: Path = base_dir / "Relatorio Final"
    pasta_screenshots: Path = pasta_logs / "screenshots"
    pasta_debug_agente: Path = pasta_logs / "debug_agente"
    pasta_imagens: Path = base_dir / "imagens_rpa"
    pasta_inativacao: Path = base_dir / "Relatório Inativação"
    pasta_relatorios_finais: Path = base_dir / "Relação por empresa 3"

    arquivo_execucao_csv: Path = pasta_logs / "execucao_consolidada.csv"
    arquivo_parada_manual: Path = pasta_logs / "PARAR.txt"
    caminho_tesseract: str = r"C:\Users\guilherme.rossi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

    # ============================================================
    # REGRAS DE NEGÓCIO
    # ============================================================
    faixa_inativar_min: int = 1
    faixa_inativar_max: int = 211
    acumuladores_excluidos_inativacao: tuple[int, ...] = (2,)

    faixa_marcacao_min: int = 212
    faixa_marcacao_max: int = 1000

    # Sem barra para evitar erros de teclado
    data_inativacao_label: str = "042026"

    # ============================================================
    # COMPORTAMENTO
    # ============================================================
    dry_run: bool = False
    usar_ocr: bool = True
    reabrir_empresa_a_cada_item: bool = False

    # ============================================================
    # TIMING (em segundos)
    # ============================================================
    pyautogui_pause: float = 0.5
    tempo_dominio_baixo: float = 7.5
    tempo_dominio: float = 16.0
    tempo_dominio_max: float = 27.0
    tempo_excel: float = 4.0
    segundos_iniciais: int = 5

    retry_tentativas: int = 3
    retry_espera: float = 2.0

    esc_loop_qtd: int = 6
    esc_loop_intervalo: float = 0.3

    # ============================================================
    # COORDENADAS (Monitor Dell 1920x1080 100%)
    # ============================================================
    x_cancelar: int = 1316
    y_cancelar: int = 304

    x_foco_listagem: int = 1527
    y_foco_listagem: int = 308

    x_excel: int = 29
    y_excel: int = 302

    x_nome_arquivo: int = 359
    y_nome_arquivo: int = 418

    x_salvar: int = 603
    y_salvar: int = 414

    x_voltar_dominio: int = 559
    y_voltar_dominio: int = 581

    x_situacao: int = 932
    y_situacao: int = 349

    x_inativo: int = 922
    y_inativo: int = 378

    x_data: int = 1192
    y_data: int = 356

    x_data_retry: int = 1182
    y_data_retry: int = 352

    x_gravar: int = 1326
    y_gravar: int = 370

    x_aba_impostos: int = 741
    y_aba_impostos: int = 462

    x_excluir_imposto: int = 1182
    y_excluir_imposto: int = 768

    x_listagem: int = 1310
    y_listagem: int = 406

    x_buscar_campo: int = 1694
    y_buscar_campo: int = 700

    x_buscar_botao: int = 1700
    y_buscar_botao: int = 699

    x_segundo_codigo_1: int = 1525
    y_segundo_codigo_1: int = 351

    # ============================================================
    # VISÃO (OCR + Template Matching)
    # ============================================================
    confidence_template_matching: float = 0.80
    usar_template_matching: bool = True
    salvar_screenshots_debug: bool = True

    cor_inativar_rgb: str = "92D050"
    cor_possivel_rgb: str = "FFF2CC"
    cor_nao_existe_rgb: str = "F4CCCC"

    regiao_situacao: tuple[int, int, int, int] = (892, 333, 90, 40)
    regiao_popup_central: tuple[int, int, int, int] = (600, 300, 720, 480)
    regiao_botao_cancelar_novo: tuple[int, int, int, int] = (1273, 291, 86, 27)

    def criar_pastas(self) -> None:
        """Cria todas as pastas necessárias."""
        pastas = [
            self.pasta_relatorios,
            self.pasta_logs,
            self.pasta_screenshots,
            self.pasta_debug_agente,
            self.pasta_imagens,
            self.pasta_inativacao,
            self.pasta_relatorios_finais,
        ]
        for pasta in pastas:
            pasta.mkdir(parents=True, exist_ok=True)


_config: Optional[AgenteRPASettings] = None


def obter_config() -> AgenteRPASettings:
    """Retorna a instância global de configuração (lazy loading)."""
    global _config
    if _config is None:
        _config = AgenteRPASettings()
        _config.criar_pastas()
    return _config


def definir_config(config: AgenteRPASettings) -> None:
    """Define a instância global de configuração (útil para testes)."""
    global _config
    _config = config


CFG = obter_config()
cfg = CFG
