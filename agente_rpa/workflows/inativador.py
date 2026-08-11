import csv
import re
import time
from pathlib import Path
from typing import Dict, List, Tuple

from agente_rpa.config.settings import CFG
from agente_rpa.core.logger import log
from agente_rpa.core.excecoes import SemAcumuladoresException, ParadaManualException
from agente_rpa.utils.retry import executar_com_retry
from agente_rpa.utils.excel import (
    aguardar_relatorio_empresa,
    carregar_empresas_ja_inativadas_por_cor,
    carregar_empresas,
    carregar_acumuladores_principais,
    processar_relatorio_extraido,
    verificar_relatorio_existente,
)
from agente_rpa.automacao.sistema_dominio import SistemaDominio


class WorkflowInativacao:
    """Gerente Sniper: lê a planilha mestra e atira código por código."""

    def __init__(self, sistema: SistemaDominio):
        self.sistema = sistema

    def _buscar_caminho_relatorio(self, codigo_empresa: int) -> Path | None:
        """Busca o relatório da empresa nas pastas padrão do projeto."""
        caminho_relatorio = verificar_relatorio_existente(codigo_empresa)

        if caminho_relatorio is None:
            nome_base = f"RELAÇÃO DE ACUMULADORES - {codigo_empresa}"
            pasta_relatorios_alt = CFG.base_dir / "Relação Por empresa 2"
            candidatos = [
                pasta_relatorios_alt / f"{nome_base}.xlsx",
                pasta_relatorios_alt / f"{nome_base}.xls",
                pasta_relatorios_alt / f"{codigo_empresa}.xlsx",
                pasta_relatorios_alt / f"{codigo_empresa}.xls",
            ]
            for candidato in candidatos:
                if candidato.exists():
                    caminho_relatorio = candidato
                    break

        return caminho_relatorio

    def _relatorio_disponivel(self, codigo_empresa: int) -> bool:
        return self._buscar_caminho_relatorio(codigo_empresa) is not None

    def _garantir_relatorio_empresa(self, codigo_empresa: int) -> bool:
        """Verifica o relatório da empresa e, se necessário, executa a extração antes da inativação."""
        if self._relatorio_disponivel(codigo_empresa):
            log.info(f"[ETAPA 1] Relatório da empresa {codigo_empresa} já está disponível.")
            return True

        log.info(f"[ETAPA 1] Relatório da empresa {codigo_empresa} não encontrado. Iniciando extração...")
        self.sistema.gerar_relatorio_acumuladores(codigo_empresa)

        try:
            caminho_relatorio = aguardar_relatorio_empresa(codigo_empresa, timeout=20)
        except FileNotFoundError as exc:
            log.info(f"[AVISO] A extração da empresa {codigo_empresa} não gerou o relatório: {exc}")
            return False

        if caminho_relatorio is not None:
            log.info(f"[ETAPA 1] Relatório da empresa {codigo_empresa} gerado com sucesso.")
            return True

        return False

    def _carregar_acumuladores_empresa_unicos(self, codigo_empresa: int, gerar_se_ausente: bool = False) -> List[int]:
        """Lê acumuladores reais da empresa e remove repetições preservando ordem."""
        caminho_relatorio = self._buscar_caminho_relatorio(codigo_empresa)

        if caminho_relatorio is None and gerar_se_ausente:
            log.info(f"[AÇÃO] Relatório da empresa {codigo_empresa} não encontrado. Gerando novo relatório no Domínio...")
            self.sistema.gerar_relatorio_acumuladores(codigo_empresa)
            try:
                caminho_relatorio = aguardar_relatorio_empresa(codigo_empresa, timeout=20)
            except FileNotFoundError as exc:
                log.info(f"[AVISO] Não foi possível gerar o relatório da empresa {codigo_empresa}: {exc}")
                return []

        if caminho_relatorio is None:
            return []

        acumuladores = processar_relatorio_extraido(caminho_relatorio)
        acumuladores_unicos: List[int] = []
        vistos = set()
        for cod in acumuladores:
            if cod in vistos:
                continue
            vistos.add(cod)
            acumuladores_unicos.append(cod)
        return acumuladores_unicos

    @staticmethod
    def _escolher_status_consolidado(status_atual: str, novo_status: str) -> str:
        """Consolida múltiplos status do mesmo acumulador vindos de logs antigos com vigências."""
        prioridade = {
            "ERRO": 4,
            "Pendente": 3,
            "PROCESSANDO...": 2,
            "Inativado": 1,
            "Já inativo": 1,
            "Pulado (Simples Nacional)": 1,
            "Nova Vigência (Inativado)": 1,
            "Nova vigência criada": 1,
            "Sobreposto por Nova Vigência": 1,
            "Superseded por nova vigência": 1,
            "Não existe na empresa": 1,
            "Ignorado (Fora da regra)": 1,
        }
        if not status_atual:
            return novo_status
        if prioridade.get(novo_status, 0) > prioridade.get(status_atual, 0):
            return novo_status
        return status_atual

    def analisar_log_existente(self, codigo_empresa: int) -> Tuple[bool, Dict[int, str], List[str]]:
        caminho_log = CFG.pasta_inativacao / f"{codigo_empresa}.txt"
        if not caminho_log.exists():
            return False, {}, []

        status_empresa = ""
        checklist_lida: Dict[int, str] = {}
        erros: List[str] = []

        try:
            with open(caminho_log, "r", encoding="utf-8") as arquivo:
                for linha in arquivo:
                    texto = linha.strip()
                    if texto.startswith("Status Geral da Empresa:"):
                        status_empresa = texto.split(":", 1)[1].strip()
                    elif re.match(r"Acumulador\s+(\d+)\s*\(Vig[eê]ncia\s+\d+\)\s*-\s*(.+)", texto, re.IGNORECASE):
                        # Formato legado: Acumulador 6 (Vigência 1) - Inativado
                        match = re.match(
                            r"Acumulador\s+(\d+)\s*\(Vig[eê]ncia\s+\d+\)\s*-\s*(.+)",
                            texto,
                            re.IGNORECASE,
                        )
                        cod = int(match.group(1))
                        status_item = match.group(2).strip()
                        checklist_lida[cod] = self._escolher_status_consolidado(
                            checklist_lida.get(cod, ""),
                            status_item,
                        )
                    elif re.match(r"Acumulador\s+(\d+)\s*-\s*(.+)", texto, re.IGNORECASE):
                        # Formato atual: Acumulador 6 - Inativado
                        match = re.match(r"Acumulador\s+(\d+)\s*-\s*(.+)", texto, re.IGNORECASE)
                        cod = int(match.group(1))
                        status_item = match.group(2).strip()
                        checklist_lida[cod] = self._escolher_status_consolidado(
                            checklist_lida.get(cod, ""),
                            status_item,
                        )
                    elif (
                        texto
                        and not texto.startswith("=")
                        and texto != "Nenhum"
                        and "Acumulador" not in texto
                        and "Erros:" not in texto
                        and "Checklist" not in texto
                        and "Fim" not in texto
                        and "Empresa:" not in texto
                        and "Status Geral" not in texto
                    ):
                        erros.append(texto)

            tem_pendencias = any(res in {"Pendente", "ERRO"} for res in checklist_lida.values())
            tem_relatorio = self._relatorio_disponivel(codigo_empresa)
            finalizada = (
                status_empresa in {
                    "OK",
                    "PROCESSADA (SEM ALVOS)",
                    "PROCESSADA (VAZIA)",
                    "PROCESSADA (SEM ACUMULADORES CADASTRADOS)",
                    "PROCESSADA (JÁ INATIVADA NA PLANILHA)",
                }
                and not tem_pendencias
                and tem_relatorio
            )
            return finalizada, checklist_lida, erros
        except Exception:
            return False, {}, []

    def registrar_csv(self, empresa: int, nome: str, acumulador: str, status: str, motivo: str) -> None:
        existe = CFG.arquivo_execucao_csv.exists()
        campos = ["timestamp", "empresa", "nome_empresa", "acumulador", "status", "motivo"]
        with open(CFG.arquivo_execucao_csv, "a", newline="", encoding="utf-8-sig") as arquivo:
            writer = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
            if not existe:
                writer.writeheader()
            writer.writerow({
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                "empresa": empresa,
                "nome_empresa": nome,
                "acumulador": acumulador,
                "status": status,
                "motivo": motivo,
            })

    def salvar_log(self, codigo: int, nome: str, checklist: Dict[int, str], erros: List[str], status: str) -> None:
        caminho = CFG.pasta_inativacao / f"{codigo}.txt"
        with open(caminho, "w", encoding="utf-8") as arquivo:
            arquivo.write("===================================\n")
            arquivo.write(f"Empresa: {codigo} - {nome}\n")
            arquivo.write(f"Status Geral da Empresa: {status}\n")
            arquivo.write("===================================\n")
            arquivo.write("Checklist de Acumuladores:\n")
            if checklist:
                for cod, res in sorted(checklist.items()):
                    arquivo.write(f"Acumulador {cod} - {res}\n")
            else:
                arquivo.write("Nenhum\n")
            arquivo.write("===================================\n")
            arquivo.write("Erros:\n")
            for erro in erros:
                arquivo.write(f"{erro}\n")
            if not erros:
                arquivo.write("Nenhum\n")
            arquivo.write("===================================\n")
            arquivo.write("Fim do log.\n")

    def executar_lote(self) -> None:
        empresas = carregar_empresas(CFG.planilha_empresas)
        lista_mestre = carregar_acumuladores_principais(CFG.planilha_acumuladores)
        caminho_planilha_controle = CFG.base_dir / "Cópia de RELAÇÃO DE EMPRESAS- matheus (1).xls"
        empresas_ja_inativadas = carregar_empresas_ja_inativadas_por_cor(caminho_planilha_controle)

        alvos_oficiais = [item.codigo for item in lista_mestre if item.deve_inativar]

        log.info(f"Posicione a tela do Domínio ativa. Iniciando em {CFG.segundos_iniciais} segundos.")
        time.sleep(CFG.segundos_iniciais)

        for empresa in empresas:
            codigo_empresa = int(empresa.codigo)
            nome_empresa = str(empresa.nome)

            if codigo_empresa in empresas_ja_inativadas:
                motivo = "Empresa marcada em amarelo na planilha de controle (já inativada)."
                log.info(f"--- [PULANDO] Empresa {codigo_empresa} marcada como concluída na planilha amarela. ---")
                self.salvar_log(
                    codigo_empresa,
                    nome_empresa,
                    {},
                    [motivo],
                    "PROCESSADA (JÁ INATIVADA NA PLANILHA)",
                )
                self.registrar_csv(codigo_empresa, nome_empresa, "N/A", "PULADO", motivo)
                continue

            self.sistema.verificar_parada()
            finalizada, checklist_atual, erros = self.analisar_log_existente(codigo_empresa)

            if finalizada:
                log.info(f"--- [PULANDO] Empresa {codigo_empresa} já concluída. ---")
                continue

            log.info(f"--- Processando Empresa: {codigo_empresa} ({nome_empresa}) ---")
            status_log = "OK"

            try:
                self.sistema.entrar_em_empresa(codigo_empresa)
                executar_com_retry(
                    lambda: self.sistema.entrar_em_acumuladores_por_atalho(),
                    descricao="Entrar Acumuladores",
                )

                relatorio_ok = self._garantir_relatorio_empresa(codigo_empresa)
                if not relatorio_ok:
                    status_log = "ERRO FATAL"
                    erros.append(f"Falha na extração do relatório da empresa {codigo_empresa}.")
                    self.registrar_csv(codigo_empresa, nome_empresa, "N/A", "ERRO", "Relatório não gerado")
                    self.salvar_log(codigo_empresa, nome_empresa, checklist_atual, erros, status_log)
                    continue

                if not checklist_atual:
                    acumuladores_empresa = self._carregar_acumuladores_empresa_unicos(
                        codigo_empresa, gerar_se_ausente=False
                    )
                    if acumuladores_empresa:
                        alvos_set = set(alvos_oficiais)
                        for cod in acumuladores_empresa:
                            if cod in alvos_set:
                                checklist_atual[cod] = "Pendente"
                            else:
                                checklist_atual[cod] = "Ignorado (Fora da regra)"
                    else:
                        log.info(
                            f"[AVISO] Relatório da empresa {codigo_empresa} não encontrado em Relação Por empresa/Relação Por empresa 2."
                        )
                        for cod in alvos_oficiais:
                            checklist_atual[cod] = "Pendente"

                self.salvar_log(codigo_empresa, nome_empresa, checklist_atual, erros, "PROCESSANDO...")

                alvos_pendentes = [
                    cod for cod, status in checklist_atual.items() if status in {"Pendente", "ERRO"}
                ]

                if not alvos_pendentes:
                    self.salvar_log(codigo_empresa, nome_empresa, checklist_atual, erros, "OK")
                    continue

                for cod in sorted(alvos_pendentes):
                    self.sistema.verificar_parada()
                    log.info(f"--- [SNIPER] Atirando no Acumulador {cod} ---")

                    try:
                        resultado = executar_com_retry(
                            lambda c=cod: self.sistema.inativar_por_nova_vigencia(str(c)),
                            descricao=f"Inativar acumulador {cod}",
                        )

                        if resultado == "INATIVADO":
                            checklist_atual[cod] = "Inativado"
                            self.registrar_csv(
                                codigo_empresa, nome_empresa, str(cod), "INATIVADO", "Nova Vigência Criada"
                            )
                        elif resultado == "JA_INATIVO":
                            checklist_atual[cod] = "Já inativo"
                            self.registrar_csv(
                                codigo_empresa, nome_empresa, str(cod), "PULADO", "Já estava inativo"
                            )
                        elif resultado == "NAO_EXISTE":
                            checklist_atual[cod] = "Não existe na empresa"
                            self.registrar_csv(
                                codigo_empresa, nome_empresa, str(cod), "PULADO", "Código inexistente"
                            )
                        elif resultado == "SIMPLES_NACIONAL":
                            checklist_atual[cod] = "Pulado (Simples Nacional)"
                            self.registrar_csv(
                                codigo_empresa,
                                nome_empresa,
                                str(cod),
                                "PULADO",
                                "Simples Nacional",
                            )
                    except Exception as exc:
                        checklist_atual[cod] = "ERRO"
                        erros.append(f"Falha Acumulador {cod}: {exc}")
                        self.registrar_csv(codigo_empresa, nome_empresa, str(cod), "ERRO", str(exc))
                        self.sistema.screenshot_debug(f"erro_sniper_{codigo_empresa}_{cod}")

                        self.sistema.fechar_sistema_para_loop(2, 0.3)
                        self.sistema.hotkey("alt", "a", espera=0.5)
                        self.sistema.pressionar("a", 1.0)
                        self.sistema.pressionar("enter", CFG.tempo_dominio)
                        self.sistema.hotkey("alt", "c", espera=1.0)
                        self.sistema.hotkey("alt", "l", espera=1.5)

                    self.salvar_log(codigo_empresa, nome_empresa, checklist_atual, erros, "PROCESSANDO...")

            except SemAcumuladoresException:
                status_log = "PROCESSADA (VAZIA)"
                self.registrar_csv(codigo_empresa, nome_empresa, "N/A", "PULADO", "Empresa vazia")
            except ParadaManualException as exc:
                status_log = "INTERROMPIDO MANUALMENTE"
                erros.append(str(exc))
                self.salvar_log(codigo_empresa, nome_empresa, checklist_atual, erros, status_log)
                raise
            except Exception as exc:
                status_log = "ERRO FATAL"
                erros.append(str(exc))
                self.registrar_csv(codigo_empresa, nome_empresa, "N/A", "ERRO_EMPRESA", str(exc))
            finally:
                if erros and status_log == "OK":
                    status_log = "ERRO PARCIAL"
                self.salvar_log(codigo_empresa, nome_empresa, checklist_atual, erros, status_log)
                self.sistema.fechar_sistema_para_loop(CFG.esc_loop_qtd, CFG.esc_loop_intervalo)

    def executar(self) -> None:
        """Alias para compatibilidade."""
        self.executar_lote()
