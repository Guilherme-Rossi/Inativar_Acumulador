from __future__ import annotations

import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pyautogui


@dataclass(frozen=True)
class AgenteConfig:
    pasta_imagens: Path
    confidence: float = 0.80
    grayscale: bool = False
    pyautogui_pause: float = 0.5
    tempo_padrao: float = 5.0
    tempo_max: float = 15.0
    usar_imagem: bool = True
    salvar_screenshots_debug: bool = True
    pasta_debug: Optional[Path] = None
    failsafe: bool = True
    arquivo_parada_manual: Optional[Path] = None


class AgenteDominio:
    def __init__(self, cfg: AgenteConfig):
        if pyautogui is None:
            raise RuntimeError("Biblioteca 'pyautogui' não encontrada.")

        self.cfg = cfg
        pyautogui.FAILSAFE = cfg.failsafe
        pyautogui.PAUSE = cfg.pyautogui_pause

        if self.cfg.pasta_debug:
            self.cfg.pasta_debug.mkdir(parents=True, exist_ok=True)

    def verificar_parada_manual(self) -> None:
        if self.cfg.arquivo_parada_manual and self.cfg.arquivo_parada_manual.exists():
            raise RuntimeError(f"Execução interrompida manualmente: {self.cfg.arquivo_parada_manual}")

    def aguardar(self, seg: Optional[float] = None) -> None:
        time.sleep(seg if seg is not None else self.cfg.tempo_padrao)

    def screenshot_debug(self, nome: str) -> str:
        if not self.cfg.salvar_screenshots_debug or not self.cfg.pasta_debug:
            return ""

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        caminho = self.cfg.pasta_debug / f"{nome}_{timestamp}.png"
        pyautogui.screenshot(str(caminho))
        return str(caminho)

    def clicar_coordenada(self, x: int, y: int, seg: Optional[float] = None, descricao: str = "") -> None:
        pyautogui.click(x, y)
        self.aguardar(seg)

    def escrever(self, texto: str, seg: float = 1.0) -> None:
        pyautogui.write(str(texto), interval=0.04)
        self.aguardar(seg)

    def pressionar(self, tecla: str, seg: Optional[float] = None) -> None:
        pyautogui.press(tecla)
        self.aguardar(seg)

    def hotkey(self, *teclas: str, seg: Optional[float] = None) -> None:
        pyautogui.hotkey(*teclas)
        self.aguardar(seg)

    def limpar_campo(self, seg: float = 1.0) -> None:
        pyautogui.hotkey("ctrl", "a")
        self.aguardar(seg)
        pyautogui.press("backspace")
        self.aguardar(seg)

    def fechar_sistema_para_loop(self, quantidade_esc: int = 10, intervalo: float = 0.5) -> None:
        for _ in range(quantidade_esc):
            pyautogui.press("esc")
            time.sleep(intervalo)

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

    def localizar_imagem(
        self,
        nome_arquivo: str,
        confidence: Optional[float] = None,
        region: Optional[tuple[int, int, int, int]] = None,
    ):
        if not self.cfg.usar_imagem:
            return None

        caminho = self.cfg.pasta_imagens / nome_arquivo
        if not caminho.exists():
            return None

        try:
            kwargs = {
                "confidence": confidence or self.cfg.confidence,
                "grayscale": self.cfg.grayscale,
            }

            if region is not None:
                kwargs["region"] = region

            return pyautogui.locateCenterOnScreen(
                str(caminho),
                **kwargs,
            )

        except Exception:
            return None

    def clicar_imagem_ou_coordenada(
        self,
        imagem: str,
        fallback_x: int,
        fallback_y: int,
        seg: Optional[float] = None,
        confidence: Optional[float] = None,
        descricao: str = "",
    ) -> tuple[int, int, str]:
        pos = self.localizar_imagem(imagem, confidence=confidence)

        if pos:
            pyautogui.click(pos.x, pos.y)
            self.aguardar(seg)
            return pos.x, pos.y, f"imagem:{imagem}"

        pyautogui.click(fallback_x, fallback_y)
        self.aguardar(seg)
        return fallback_x, fallback_y, f"fallback:{descricao or imagem}"

    def ocr_regiao(self, regiao: tuple[int, int, int, int], nome_debug: str = "") -> str:
        try:
            import pytesseract

            img = pyautogui.screenshot(region=regiao)

            if nome_debug and self.cfg.pasta_debug:
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                img.save(str(self.cfg.pasta_debug / f"{nome_debug}_{timestamp}.png"))

            texto = pytesseract.image_to_string(img, lang="por")
            return re.sub(r"\s+", " ", str(texto or "")).strip()

        except Exception:
            return ""