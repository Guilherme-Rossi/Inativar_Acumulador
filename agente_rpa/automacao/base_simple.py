"""
Camada de automação de baixo nível (mouse, teclado, visão).

Responsável por:
- Cliques e movimentação do mouse
- Pressionamento de teclado
- OCR (Tesseract)
- Screenshots de debug
"""

from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Optional

try:
    import pyautogui
except ImportError:
    pyautogui = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

from agente_rpa.config import CFG
from agente_rpa.core.logger import log


class AutomacaoDominio:
    """
    Camada baixo-nível de automação RPA.

    Métodos:
    - click, write, press, hotkey: Ações básicas
    - take_screenshot: Captura de tela
    - ocr_region: Reconhecimento de texto
    """

    def __init__(self):
        """Inicializa a automação."""
        config = CFG

        if pyautogui is None:
            raise RuntimeError("Biblioteca 'pyautogui' não encontrada")

        # Configurar pyautogui globalmente
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = config.pyautogui_pause

        # Verificar Tesseract se usando OCR
        if config.usar_ocr:
            try:
                caminho_tess = Path(config.caminho_tesseract)
                if caminho_tess.exists():
                    pytesseract.pytesseract.tesseract_cmd = str(caminho_tess)
                    log.info(f"Tesseract configurado: {caminho_tess}")
            except Exception as e:
                log.info(f"Aviso ao configurar Tesseract: {e}")

        log.info("AutomacaoDominio inicializado")

    def click(self, x: int, y: int, tempo_espera: Optional[float] = None) -> None:
        """Clica em uma coordenada."""
        pyautogui.click(x, y)
        if tempo_espera is not None:
            time.sleep(tempo_espera)

    def write(self, texto: str, tempo_espera: float = 1.0) -> None:
        """Escreve texto como se fosse teclado."""
        texto_str = str(texto)
        pyautogui.write(texto_str, interval=0.04)
        time.sleep(tempo_espera)

    def press(self, tecla: str, tempo_espera: Optional[float] = None) -> None:
        """Pressiona uma tecla."""
        pyautogui.press(tecla)
        if tempo_espera is not None:
            time.sleep(tempo_espera)

    def hotkey(self, *teclas: str, tempo_espera: Optional[float] = None) -> None:
        """Pressiona combinação de teclas (Ctrl+S, etc)."""
        pyautogui.hotkey(*teclas)
        if tempo_espera is not None:
            time.sleep(tempo_espera)

    def wait(self, segundos: float) -> None:
        """Aguarda X segundos."""
        time.sleep(segundos)

    def clear_field(self, tempo_espera: float = 1.0) -> None:
        """Limpa campo de texto (Ctrl+A + Delete)."""
        pyautogui.hotkey("ctrl", "a")
        self.wait(tempo_espera)
        pyautogui.press("backspace")
        self.wait(tempo_espera)

    def close_windows_with_esc(self, quantidade: int = 6, intervalo: float = 0.1) -> None:
        """Pressiona ESC múltiplas vezes para fechar popups/modals."""
        for _ in range(quantidade):
            pyautogui.press("esc")
            time.sleep(intervalo)

    def take_screenshot(self, nome: str = "") -> str:
        """Captura uma screenshot da tela inteira."""
        config = CFG
        timestamp = time.strftime("%Y%m%d_%H%M%S")

        if nome:
            nome_arquivo = f"{nome}_{timestamp}.png"
        else:
            nome_arquivo = f"screenshot_{timestamp}.png"

        caminho = config.pasta_screenshots / nome_arquivo
        pyautogui.screenshot(str(caminho))

        log.info(f"Screenshot salvo: {caminho}")
        return str(caminho)

    def ocr_region(
        self,
        regiao: tuple[int, int, int, int],
        idioma: str = "por",
        nome_debug: str = "",
    ) -> str:
        """
        Executa OCR em uma região da tela.

        Args:
            regiao: (left, top, width, height)
            idioma: Idioma para OCR (padrão: português)
            nome_debug: Se fornecido, salva screenshot dessa região

        Returns:
            Texto extraído
        """
        config = CFG

        if not config.usar_ocr or pytesseract is None:
            return ""

        try:
            imagem = pyautogui.screenshot(region=regiao)

            if nome_debug and config.pasta_debug_agente:
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                debug_path = config.pasta_debug_agente / f"{nome_debug}_{timestamp}.png"
                imagem.save(str(debug_path))

            texto = pytesseract.image_to_string(imagem, lang=idioma)
            texto_limpo = re.sub(r"\s+", " ", str(texto or "")).strip()

            return texto_limpo

        except Exception as e:
            log.info(f"Erro ao executar OCR: {e}")
            return ""
