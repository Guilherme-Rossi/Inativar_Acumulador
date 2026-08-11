"""
Motor de interface simplificado (AutomacaoBase).

Esta versão é intencionalmente minimalista: não conhece regras de negócio,
apenas ações de mouse/teclado e verificações básicas (arquivo de parada).
"""

import re
import time
from pathlib import Path
from typing import Optional

import pyautogui

from agente_rpa.core.logger import log
from agente_rpa.core.excecoes import ParadaManualException


class AutomacaoBase:
    """ Wrapper simplificado do PyAutoGUI (Motor Base) """

    def __init__(self, pasta_imagens: Path, arquivo_parada: Path):
        self.pasta_imagens = pasta_imagens
        self.arquivo_parada = arquivo_parada
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5

    def verificar_parada(self) -> None:
        if self.arquivo_parada.exists():
            raise ParadaManualException(f"Parada acionada via arquivo: {self.arquivo_parada.name}")

    def aguardar(self, segundos: float) -> None:
        time.sleep(segundos)

    def clicar_coordenada(self, x: int, y: int, espera: float = 0.5, log_msg: str = "") -> None:
        self.verificar_parada()
        if log_msg:
            log.info(f"[AÇÃO] Clicando em {log_msg} (X:{x}, Y:{y})")
        pyautogui.click(x, y)
        self.aguardar(espera)

    def escrever(self, texto: str, espera: float = 1.0) -> None:
        self.verificar_parada()
        pyautogui.write(str(texto), interval=0.04)
        self.aguardar(espera)

    def pressionar(self, tecla: str, espera: float = 0.5) -> None:
        self.verificar_parada()
        pyautogui.press(tecla)
        self.aguardar(espera)

    def hotkey(self, tecla1: str, tecla2: str, espera: float = 1.0) -> None:
        self.verificar_parada()
        pyautogui.hotkey(tecla1, tecla2)
        self.aguardar(espera)

    def sequencia_alt_aa(self, seg_entre_teclas: float = 0.5, seg_final: Optional[float] = None) -> None:
        try:
            pyautogui.keyDown("alt")
            time.sleep(seg_entre_teclas)
            pyautogui.press("a")
            time.sleep(seg_entre_teclas)
            pyautogui.press("a")
            time.sleep(seg_entre_teclas)
        finally:
            pyautogui.keyUp("alt")

        if seg_final is not None:
            time.sleep(seg_final)

    def localizar_imagem(self, nome_arquivo: str, confidence: float = 0.8, region: tuple = None):
        caminho = self.pasta_imagens / nome_arquivo
        if not caminho.exists():
            return None

        kwargs = {"confidence": confidence, "grayscale": False}
        if region:
            kwargs["region"] = region

        try:
            return pyautogui.locateCenterOnScreen(str(caminho), **kwargs)
        except Exception:
            return None

    def ocr_regiao(self, regiao: tuple[int, int, int, int], nome_debug: str = "") -> str:
        try:
            import pytesseract
            from agente_rpa.config.settings import CFG

            pytesseract.pytesseract.tesseract_cmd = CFG.caminho_tesseract
            img = pyautogui.screenshot(region=regiao)

            if nome_debug and CFG.salvar_screenshots_debug:
                CFG.pasta_debug_agente.mkdir(parents=True, exist_ok=True)
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                img.save(str(CFG.pasta_debug_agente / f"{nome_debug}_{timestamp}.png"))

            texto = pytesseract.image_to_string(img, lang="por")
            return re.sub(r"\s+", " ", str(texto or "")).strip()
        except Exception:
            return ""

    def screenshot_debug(self, nome: str) -> str:
        from agente_rpa.config.settings import CFG

        if not CFG.salvar_screenshots_debug:
            return ""

        CFG.pasta_debug_agente.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        caminho = CFG.pasta_debug_agente / f"{nome}_{timestamp}.png"
        pyautogui.screenshot(str(caminho))
        return str(caminho)

    def fechar_sistema_para_loop(self, quantidade_esc: int = 10, intervalo: float = 0.5) -> None:
        for _ in range(quantidade_esc):
            self.pressionar("esc", intervalo)
