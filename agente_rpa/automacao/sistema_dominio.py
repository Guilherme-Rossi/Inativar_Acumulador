import time

import pyautogui

from typing import Optional

from agente_rpa.config.settings import CFG
from agente_rpa.core.logger import log
from agente_rpa.core.excecoes import SemAcumuladoresException
from agente_rpa.automacao.base import AutomacaoBase


class SistemaDominio(AutomacaoBase):
    """Operador Sniper do Domínio: busca, cria nova vigência e inativa."""

    def entrar_em_empresa(self, codigo_empresa: int) -> None:
        self.verificar_parada()
        log.info(f"[AÇÃO] Trocando para a Empresa {codigo_empresa}...")
        self.pressionar("f8", CFG.tempo_dominio_baixo)
        self.escrever(str(codigo_empresa), CFG.tempo_dominio_baixo)
        self.pressionar("enter", CFG.tempo_dominio_baixo)

    def entrar_em_acumuladores_por_atalho(self) -> None:
        max_tentativas = 3
        tela_aberta = False
        cadastro_vazio = False

        for tentativa in range(1, max_tentativas + 1):
            self.verificar_parada()
            log.info(f"[INFO] Tentando entrar em Acumuladores (Alt+A+A) - Tentativa {tentativa}/{max_tentativas}")

            if tentativa > 1:
                log.info("[INFO] Limpando tela com ESCs antes de tentar novamente.")
                self.fechar_sistema_para_loop(quantidade_esc=3, intervalo=0.2)

            self.sequencia_alt_aa(seg_entre_teclas=0.5, seg_final=1.0)
            self.pressionar("enter", 0.5)

            tempo_espera = time.time()
            while (time.time() - tempo_espera) < CFG.tempo_dominio_max:
                if self.localizar_imagem("tela_aviso1.png", 0.7) or self.localizar_imagem("tela_aviso.png", 0.7):
                    cadastro_vazio = True
                    tela_aberta = True
                    break

                if (
                    self.localizar_imagem("tela_inteira.png", 0.7)
                    or self.localizar_imagem("tela_acumuladores.png", 0.7)
                    or self.localizar_imagem("btn_cancelar2.png", 0.7)
                    or self.localizar_imagem("tela_novo.png", 0.7)
                    or self.localizar_imagem("btn_novo2.png", 0.7)
                ):
                    tela_aberta = True
                    break
                time.sleep(1)

            if tela_aberta:
                break

        if not tela_aberta:
            raise RuntimeError("Falha ao abrir a tela de Acumuladores.")

        if cadastro_vazio:
            self.pressionar("enter", 0.5)
            raise SemAcumuladoresException("Empresa não possui acumuladores cadastrados.")

        time.sleep(3.0)

        log.info("[ATALHO] Pressionando Alt+C (Cancelar) por segurança.")
        self.hotkey("alt", "c", espera=2.0)

        log.info("[ATALHO] Abrindo Listagem (Alt+L)...")
        self.hotkey("alt", "l", espera=2.0)

    def verificar_se_ja_inativo(self, str_cod: str) -> bool:
        if (
            self.localizar_imagem("situacao_inativa.png", 0.75, CFG.regiao_situacao)
            or self.localizar_imagem("situacao_inativa_azul.png", 0.75, CFG.regiao_situacao)
        ):
            return True

        if CFG.usar_ocr:
            texto = self.ocr_regiao(CFG.regiao_situacao, f"ocr_{str_cod}").lower()
            if "inativ" in texto:
                return True
        return False

    def gerar_relatorio_acumuladores(self, codigo_empresa: int) -> None:
        """Gera o relatório de acumuladores a partir da tela aberta no Domínio."""
        self.verificar_parada()
        log.info(f"[AÇÃO] Gerando relatório de acumuladores para a empresa {codigo_empresa}...")

        self.hotkey("alt", "l", espera=CFG.tempo_dominio_baixo)
        self.hotkey("alt", "r", espera=CFG.tempo_dominio_baixo)

        tempo_espera = time.time()
        while (time.time() - tempo_espera) < CFG.tempo_dominio_max:
            if self.localizar_imagem("tela_listagem_aberta.png", 0.7) or self.localizar_imagem("tela_listagem.png", 0.7):
                break
            time.sleep(1.0)

        self.pressionar("down", CFG.tempo_dominio_baixo)

        pausa_original = pyautogui.PAUSE
        pyautogui.PAUSE = 0.0
        pyautogui.press("tab", presses=20, interval=0.02)
        pyautogui.PAUSE = pausa_original

        self.pressionar("enter", CFG.tempo_dominio)

        salvou = False
        for tentativa in range(3):
            pos_excel = self.localizar_imagem("tela_excel.png", 0.7)
            if pos_excel:
                self.clicar_coordenada(pos_excel.x, pos_excel.y, 0.5, log_msg="ícone do Excel")
            else:
                self.clicar_coordenada(CFG.x_excel, CFG.y_excel, 0.5, log_msg="opção Excel")

            time.sleep(1.0)

            tempo_salvar = time.time()
            pos_salvar = None
            while (time.time() - tempo_salvar) < 10.0:
                pos_salvar = self.localizar_imagem("tela_salvar_excel.png", 0.8)
                if pos_salvar:
                    break
                time.sleep(1.0)

            if pos_salvar:
                self.clicar_coordenada(CFG.x_nome_arquivo, CFG.y_nome_arquivo, 0.5, log_msg="campo de nome do arquivo")
                self.escrever(f" - {codigo_empresa}", 0.5)
                self.clicar_coordenada(CFG.x_salvar, CFG.y_salvar, 0.5, log_msg="botão salvar")
                salvou = True
                break

        if not salvou:
            self.clicar_coordenada(CFG.x_nome_arquivo, CFG.y_nome_arquivo, 0.5, log_msg="campo de nome do arquivo")
            self.escrever(f" - {codigo_empresa}", 0.5)
            self.clicar_coordenada(CFG.x_salvar, CFG.y_salvar, 0.5, log_msg="botão salvar")

        time.sleep(CFG.tempo_dominio_baixo)
        self.hotkey("alt", "f4", espera=CFG.tempo_dominio_baixo)
        self.clicar_coordenada(CFG.x_voltar_dominio, CFG.y_voltar_dominio, 0.5, log_msg="Domínio")
        self.pressionar("esc", CFG.tempo_dominio_baixo)

    def inativar_por_nova_vigencia(self, cod_alvo: str) -> str:
        """Busca o acumulador e sempre cria nova vigência já inativada."""
        self.verificar_parada()

        log.info(f"[ATALHO] Focando na busca (Alt+S) para o código {cod_alvo}...")
        self.hotkey("alt", "s", espera=1.0)
        self.escrever(cod_alvo, 0.5)

        if cod_alvo == "1":
            self.clicar_coordenada(CFG.x_segundo_codigo_1, CFG.y_segundo_codigo_1, 0.5)

        log.info("[ATALHO] Buscando (Alt+B)...")
        self.hotkey("alt", "b", espera=3.0)

        if self.localizar_imagem("registro_nao_encontrado.png", 0.7):
            log.info(f"[AVISO] O Acumulador {cod_alvo} não existe nesta empresa. Fechando aviso.")
            self.pressionar("enter", 1.0)
            return "NAO_EXISTE"

        # Entrar em edição do item encontrado.
        log.info("[AÇÃO] Entrando no acumulador encontrado (Enter)...")
        self.pressionar("enter", 1.0)

        # Garantir que não estamos alterando a vigência atual diretamente.
        log.info("[ATALHO] Cancelando edição atual (Alt+C) para abrir Nova Vigência limpa...")
        self.hotkey("alt", "c", espera=1.2)

        log.info("[ATALHO] Criando NOVA VIGÊNCIA (Alt+O)...")
        self.hotkey("alt", "o", espera=2.0)

        log.info("[AÇÃO] Alterando Situação para Inativo (seta para baixo + enter)...")
        self.clicar_coordenada(CFG.x_situacao, CFG.y_situacao, 0.8)
        self.pressionar("down", 0.6)
        self.pressionar("enter", 0.8)

        log.info("[AÇÃO] Ajustando período da nova vigência (TAB + 2x seta para baixo)...")
        # self.pressionar("tab", 0.8) sem necessidade
        self.pressionar("down", 0.6)
        self.pressionar("down", 0.8)
        
        resultado_gravacao = self._gravar_e_tratar_impostos(cod_alvo)

        if resultado_gravacao == "SIMPLES_NACIONAL":
            return resultado_gravacao

        self.clicar_coordenada(CFG.x_foco_listagem, CFG.y_foco_listagem, 0.5)
        return "INATIVADO"

    def _gravar_e_tratar_impostos(self, str_cod: str) -> str:
        log.info("[ATALHO] Gravando alteração (Alt+G)...")
        self.hotkey("alt", "g", espera=2.5)

        tempo_verif = time.time()
        while (time.time() - tempo_verif) < 5.0:
            self.verificar_parada()

            if self.localizar_imagem("data_inativacao.png", 0.7):
                log.info("[VISÃO] Aviso de data detectado. Confirmando e ajustando data (2x seta para baixo)...")
                self.pressionar("enter", 1.0)
                self.pressionar("down", 0.6)
                self.pressionar("down", 0.8)
                log.info("[ATALHO] Regravando após ajuste de data (Alt+G)...")
                self.hotkey("alt", "g", espera=2.2)
                tempo_verif = time.time()
                continue

            if self.localizar_imagem("tela_data_invalida.png", 0.7):
                self.pressionar("enter", 0.5)
                self.pressionar("esc", 0.5)
                raise RuntimeError("O sistema recusou a data de inativação (Data Inválida).")

            if self.localizar_imagem("sn1.png", 0.8):
                log.info("[VISÃO] Aviso de Simples Nacional detectado. Confirmando, cancelando, TAB/ENTER e pulando acumulador...")
                self.pressionar("enter", 1.0)
                self.hotkey("alt", "c", espera=1.0)
                self.pressionar("tab", 0.5)
                self.pressionar("enter", 0.8)
                self.clicar_coordenada(CFG.x_foco_listagem, CFG.y_foco_listagem, 0.5)
                return "SIMPLES_NACIONAL"

            texto_popup = ""
            if CFG.usar_ocr:
                try:
                    texto_popup = self.ocr_regiao(CFG.regiao_popup_central).lower()
                except Exception:
                    texto_popup = ""

            if (
                self.localizar_imagem("aviso_imposto.png", 0.7)
                or self.localizar_imagem("problema_imposto.png", 0.7)
                or self.localizar_imagem("Imposto3.png", 0.7)
                or self.localizar_imagem("Imposto4.png", 0.7)
                or self.localizar_imagem("Imposto5.png", 0.7)
                or "cadastrado" in texto_popup
                or "parâmetros" in texto_popup
            ):
                log.info("[VISÃO] Aviso de Imposto detectado! Limpando...")
                self.pressionar("enter", 1.0)

                pos_impostos = self.localizar_imagem("aba_impostos.png", 0.8)
                if pos_impostos:
                    self.clicar_coordenada(pos_impostos.x, pos_impostos.y, 1.0)
                else:
                    self.clicar_coordenada(CFG.x_aba_impostos, CFG.y_aba_impostos, 1.0)

                for _ in range(5):
                    pos_excluir = self.localizar_imagem("btn_excluir.png", 0.8)
                    if pos_excluir:
                        self.clicar_coordenada(pos_excluir.x, pos_excluir.y, 0.4)
                    else:
                        self.clicar_coordenada(CFG.x_excluir_imposto, CFG.y_excluir_imposto, 0.4)

                time.sleep(1.2)
                self.clicar_coordenada(CFG.x_situacao, CFG.y_situacao, 0.5)
                log.info("[ATALHO] Regravando após exclusão de impostos (Alt+G)...")
                self.hotkey("alt", "g", espera=2.0)
                tempo_verif = time.time()
                continue

            if self.localizar_imagem("cfop_selecao.png", 0.7):
                log.info("[VISÃO] Aviso CFOP de seleção detectado. Confirmando seleção e regravando...")
                # Centro aproximado informado: (608, 496)
                self.pressionar("enter", 0.8)
                self.clicar_coordenada(608, 496, 0.6)
                # self.pressionar("tab", 0.5)
                
                self.hotkey("alt", "g", espera=2.0)
                tempo_verif = time.time()
                continue

            if self.localizar_imagem("cfop2.png", 0.7):
                log.info("[VISÃO] Aviso CFOP2 detectado. Confirmando com TAB/ENTER e regravando...")
                self.pressionar("tab", 0.5)
                self.pressionar("enter", 0.8)
                self.hotkey("alt", "g", espera=2.0)
                tempo_verif = time.time()
                continue

            if self.localizar_imagem("vigencia1.png", 0.7):
                log.info("[VISÃO] Aviso Vigencia1 detectado. Confirmando com Enter e continuando...")
                self.pressionar("enter", 1.0)
                tempo_verif = time.time()
                continue

            if self.localizar_imagem("vigencia2.png", 0.7):
                log.info("[VISÃO] Aviso Vigencia2 detectado. Confirmando com TAB/ENTER e regravando...")
                self.pressionar("tab", 0.5)
                self.pressionar("enter", 0.8)
                self.hotkey("alt", "g", espera=2.0)
                tempo_verif = time.time()
                continue

            if self.localizar_imagem("ex_cfop.png", 0.7) or self.localizar_imagem("mostrar_rateio.png", 0.7):
                self.pressionar("enter", 1.0)
                tempo_verif = time.time()
                continue

            time.sleep(0.5)

        self.pressionar("enter", 1.0)
        return "INATIVADO"
