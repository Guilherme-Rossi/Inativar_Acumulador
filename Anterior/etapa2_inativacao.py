from __future__ import annotations

import csv
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import pandas as pd
import pyautogui

from agente_dominio import AgenteConfig, AgenteDominio


# ============================================================
# CONFIGURACAO
# ============================================================
@dataclass(frozen=True)
class Config:
    base_dir: Path = Path(r"C:\Users\guilherme.rossi\Desktop\PYTHON DOCS\Projects Agrelli\P1 - Inativar Acumulador")
    arquivo_validados: Path = base_dir / "VALIDADOS.xlsx"
    pasta_logs: Path = base_dir / "Relatorio Final"
    pasta_screenshots: Path = pasta_logs / "screenshots"
    arquivo_execucao_csv: Path = pasta_logs / "execucao_detalhada.csv"
    pasta_imagens: Path = base_dir / "imagens_rpa"
    pasta_debug_agente: Path = pasta_logs / "debug_agente"

    data_inativacao: str = "03/2026"
    dry_run: bool = False
    reabrir_empresa_a_cada_item: bool = False
    
    # --- OTIMIZACOES DE TEMPO ---
    segundos_iniciais: int = 5
    retry_tentativas: int = 3
    retry_espera: float = 2.0

    tempo_dominio_baixo: float = 2.0
    tempo_dominio: float = 11.0
    tempo_dominio_max: float = 14.0
    pyautogui_pause: float = 0.1

    esc_loop_qtd: int = 6
    esc_loop_intervalo: float = 0.1
    # ----------------------------

    # fallback coordenadas
    x_listagem: int = 1310
    y_listagem: int = 406
    x_buscar_campo: int = 1694
    y_buscar_campo: int = 700
    x_buscar_botao: int = 1700
    y_buscar_botao: int = 699
    x_segundo_codigo_1: int = 1525
    y_segundo_codigo_1: int = 351
    x_situacao: int = 926
    y_situacao: int = 348
    x_inativo: int = 922
    y_inativo: int = 378
    x_data: int = 1192
    y_data: int = 356
    x_gravar: int = 1326
    y_gravar: int = 370


CFG = Config()

CFG.pasta_logs.mkdir(parents=True, exist_ok=True)
CFG.pasta_screenshots.mkdir(parents=True, exist_ok=True)
CFG.pasta_debug_agente.mkdir(parents=True, exist_ok=True)

pyautogui.PAUSE = CFG.pyautogui_pause
pyautogui.FAILSAFE = True


class Logger:
    def __init__(self) -> None:
        self.log_path = CFG.pasta_logs / "etapa2_inativacao.log"

    def info(self, mensagem: str) -> None:
        linha = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}"
        print(linha)
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(linha + "\n")


LOG = Logger()


def registrar_execucao_csv(registro: dict) -> None:
    arquivo_existe = CFG.arquivo_execucao_csv.exists()
    campos =[
        "timestamp", "empresa", "nome_empresa", "acumulador", "status",
        "motivo", "screenshot"
    ]
    with open(CFG.arquivo_execucao_csv, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=campos, delimiter=';')
        if not arquivo_existe:
            writer.writeheader()
        writer.writerow(registro)


def executar_com_retry(func: Callable, tentativas: int, espera: float, descricao: str):
    ultimo_erro = None
    for tentativa in range(1, tentativas + 1):
        try:
            return func()
        except Exception as exc:
            ultimo_erro = exc
            LOG.info(f"[RETRY] {descricao} falhou na tentativa {tentativa}/{tentativas}: {exc}")
            if tentativa < tentativas:
                time.sleep(espera)
    raise ultimo_erro


def screenshot_erro(empresa: str, acumulador: str, sufixo: str) -> str:
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    nome = f"{empresa}_{acumulador}_{sufixo}_{timestamp}.png"
    caminho = CFG.pasta_screenshots / nome
    pyautogui.screenshot(str(caminho))
    return str(caminho)


def criar_agente() -> AgenteDominio:
    return AgenteDominio(
        AgenteConfig(
            pasta_imagens=CFG.pasta_imagens,
            confidence=0.80,
            grayscale=True,
            pyautogui_pause=CFG.pyautogui_pause,
            tempo_padrao=CFG.tempo_dominio,
            tempo_max=CFG.tempo_dominio_max,
            usar_imagem=False, # <--- IMAGEM DESATIVADA AQUI TAMBEM
            usar_acessibilidade=False,
            salvar_screenshots_debug=True,
            pasta_debug=CFG.pasta_debug_agente,
        )
    )


# ============================================================
# RPA
# ============================================================
def entrar_em_acumuladores_por_atalho(agente: AgenteDominio) -> None:
    LOG.info("[INFO] Entrando em Acumuladores via ALT + A + A (ALT segurado)")
    agente.sequencia_alt_aa(seg_entre_teclas=0.2, seg_final=CFG.tempo_dominio_max)


def abrir_empresa(agente: AgenteDominio, codigo_empresa: str) -> None:
    agente.pressionar("f8", CFG.tempo_dominio_baixo)
    agente.escrever(codigo_empresa, CFG.tempo_dominio_baixo)
    agente.pressionar("enter", CFG.tempo_dominio_baixo)

    entrar_em_acumuladores_por_atalho(agente)

    agente.clicar_imagem_ou_coordenada(
        "listagem.png",
        CFG.x_listagem,
        CFG.y_listagem,
        CFG.tempo_dominio,
        descricao="Listagem",
    )


def preparar_busca(agente: AgenteDominio) -> None:
    agente.clicar_imagem_ou_coordenada(
        "campo_busca.png",
        CFG.x_buscar_campo,
        CFG.y_buscar_campo,
        CFG.tempo_dominio,
        descricao="Campo busca",
    )
    agente.limpar_campo(CFG.tempo_dominio_baixo)


def buscar_acumulador(agente: AgenteDominio, codigo_acumulador: str) -> None:
    preparar_busca(agente)
    agente.escrever(codigo_acumulador, CFG.tempo_dominio_baixo)
    agente.clicar_imagem_ou_coordenada(
        "botao_buscar.png",
        CFG.x_buscar_botao,
        CFG.y_buscar_botao,
        CFG.tempo_dominio_max,
        descricao="Botao buscar",
    )


def tratar_codigo_duplicado_um(agente: AgenteDominio, codigo_acumulador: str) -> None:
    if str(codigo_acumulador) == "1":
        agente.clicar_imagem_ou_coordenada(
            "segundo_codigo_1.png",
            CFG.x_segundo_codigo_1,
            CFG.y_segundo_codigo_1,
            CFG.tempo_dominio_baixo,
            descricao="Segundo codigo 1",
        )
        agente.clicar_imagem_ou_coordenada(
            "botao_buscar.png",
            CFG.x_buscar_botao,
            CFG.y_buscar_botao,
            CFG.tempo_dominio_max,
            descricao="Botao buscar",
        )


def inativar_acumulador(agente: AgenteDominio) -> None:
    agente.clicar_imagem_ou_coordenada(
        "situacao.png",
        CFG.x_situacao,
        CFG.y_situacao,
        CFG.tempo_dominio_baixo,
        descricao="Situacao",
    )
    agente.clicar_imagem_ou_coordenada(
        "inativo.png",
        CFG.x_inativo,
        CFG.y_inativo,
        CFG.tempo_dominio_baixo,
        descricao="Inativo",
    )
    agente.clicar_imagem_ou_coordenada(
        "campo_data.png",
        CFG.x_data,
        CFG.y_data,
        CFG.tempo_dominio_baixo,
        descricao="Data",
    )
    agente.limpar_campo(CFG.tempo_dominio_baixo)
    agente.escrever(CFG.data_inativacao, CFG.tempo_dominio_baixo)

    if not CFG.dry_run:
        agente.clicar_imagem_ou_coordenada(
            "gravar.png",
            CFG.x_gravar,
            CFG.y_gravar,
            CFG.tempo_dominio_baixo,
            descricao="Gravar",
        )
    else:
        LOG.info("[DRY RUN] Gravacao pulada intencionalmente.")


# ============================================================
# VALIDACOES PREVIAS
# ============================================================
def validar_arquivo_validados(df: pd.DataFrame) -> None:
    obrigatorias = {
        "EMPRESA", "NOME_EMPRESA", "ACUMULADOR", "EXISTE", "INATIVAR"
    }
    faltando = obrigatorias - set(df.columns)
    if faltando:
        raise ValueError(f"VALIDADOS.xlsx nao possui as colunas obrigatorias: {sorted(faltando)}")


def carregar_fila_execucao() -> pd.DataFrame:
    df = pd.read_excel(CFG.arquivo_validados)
    validar_arquivo_validados(df)

    df = df[(df["INATIVAR"] == "SIM") & (df["EXISTE"] == "SIM")].copy()
    if df.empty:
        raise ValueError("Nao ha registros com INATIVAR=SIM e EXISTE=SIM no VALIDADOS.xlsx")

    df["EMPRESA"] = df["EMPRESA"].apply(lambda x: str(int(float(x))) if pd.notna(x) else "")
    df["ACUMULADOR"] = df["ACUMULADOR"].apply(lambda x: str(int(float(x))) if pd.notna(x) else "")
    df["NOME_EMPRESA"] = df["NOME_EMPRESA"].fillna("").astype(str)
    df = df.drop_duplicates(subset=["EMPRESA", "ACUMULADOR"]).sort_values(["EMPRESA", "ACUMULADOR"])
    return df


# ============================================================
# LOG POR EMPRESA
# ============================================================
def criar_log_empresa(codigo_empresa: str, nome_empresa: str, inativados: list[str], pulados: list[str], erros: list[str]) -> None:
    caminho_log = CFG.pasta_logs / f"{codigo_empresa}.txt"
    with open(caminho_log, "w", encoding="utf-8") as f:
        f.write("===================================\n")
        f.write(f"Empresa: {codigo_empresa} - {nome_empresa}\n")
        f.write("===================================\n")
        f.write("Acumuladores inativados:\n")
        if inativados:
            for item in inativados:
                f.write(f"{item}\n")
        else:
            f.write("Nenhum\n")

        f.write("===================================\n")
        f.write("Acumuladores pulados / nao confirmados:\n")
        if pulados:
            for item in pulados:
                f.write(f"{item}\n")
        else:
            f.write("Nenhum\n")

        f.write("===================================\n")
        f.write("Erros:\n")
        if erros:
            for item in erros:
                f.write(f"{item}\n")
        else:
            f.write("Nenhum\n")

        f.write("===================================\n")
        f.write("Fim do log.\n")


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar_item(agente: AgenteDominio, empresa: str, nome_empresa: str, acumulador: str) -> tuple[str, str, str]:
    try:
        executar_com_retry(
            lambda: buscar_acumulador(agente, acumulador),
            CFG.retry_tentativas,
            CFG.retry_espera,
            f"Busca do acumulador {acumulador}",
        )

        tratar_codigo_duplicado_um(agente, acumulador)

        executar_com_retry(
            lambda: inativar_acumulador(agente),
            CFG.retry_tentativas,
            CFG.retry_espera,
            f"Inativacao do acumulador {acumulador}",
        )

        if CFG.dry_run:
            return "DRY_RUN", "Simulado sem gravar", ""
        return "INATIVADO", "Inativado com sucesso", ""

    except Exception as exc:
        screenshot = ""
        try:
            screenshot = screenshot_erro(empresa, acumulador, "erro")
            agente.screenshot_debug(f"erro_item_{empresa}_{acumulador}")
        except Exception as ss_exc:
            LOG.info(f"[AVISO] Nao foi possivel gerar screenshot: {ss_exc}")
        return "ERRO", str(exc), screenshot


def processar() -> None:
    LOG.info("[INFO] Iniciando etapa 2 com atalho ALT+A")
    LOG.info(f"[INFO] Dry run: {'SIM' if CFG.dry_run else 'NAO'}")
    LOG.info("[INFO] Posicione o sistema Dominio corretamente.")
    time.sleep(CFG.segundos_iniciais)

    agente = criar_agente()
    df = carregar_fila_execucao()

    for empresa, grupo in df.groupby("EMPRESA"):
        nome_empresa = grupo["NOME_EMPRESA"].iloc[0]
        inativados: list[str] =[]
        pulados: list[str] = []
        erros: list[str] =[]

        LOG.info(f"[INFO] Processando empresa {empresa} - {nome_empresa}")

        try:
            executar_com_retry(
                lambda: abrir_empresa(agente, empresa),
                CFG.retry_tentativas,
                CFG.retry_espera,
                f"Abertura da empresa {empresa}",
            )

            for _, row in grupo.iterrows():
                acumulador = row["ACUMULADOR"]

                if CFG.reabrir_empresa_a_cada_item:
                    executar_com_retry(
                        lambda: abrir_empresa(agente, empresa),
                        CFG.retry_tentativas,
                        CFG.retry_espera,
                        f"Reabertura da empresa {empresa}",
                    )

                status, motivo, screenshot = processar_item(agente, empresa, nome_empresa, acumulador)

                if status in {"INATIVADO", "DRY_RUN"}:
                    inativados.append(acumulador)
                    LOG.info(f"[OK] Empresa {empresa} - acumulador {acumulador} - {motivo}")
                else:
                    pulados.append(acumulador)
                    erros.append(f"Acumulador {acumulador}: {motivo}")
                    LOG.info(f"[ERRO] Empresa {empresa} - acumulador {acumulador} - {motivo}")

                registrar_execucao_csv({
                    "timestamp": time.strftime('%Y-%m-%d %H:%M:%S'),
                    "empresa": empresa,
                    "nome_empresa": nome_empresa,
                    "acumulador": acumulador,
                    "status": status,
                    "motivo": motivo,
                    "screenshot": screenshot,
                })

        except Exception as exc:
            screenshot = ""
            try:
                screenshot = screenshot_erro(empresa, "empresa", "falha_geral")
                agente.screenshot_debug(f"falha_geral_empresa_{empresa}")
            except Exception:
                pass
            msg = f"Falha geral na empresa {empresa}: {exc}"
            erros.append(msg)
            LOG.info(f"[ERRO] {msg}")
            registrar_execucao_csv({
                "timestamp": time.strftime('%Y-%m-%d %H:%M:%S'),
                "empresa": empresa,
                "nome_empresa": nome_empresa,
                "acumulador": "",
                "status": "ERRO_EMPRESA",
                "motivo": str(exc),
                "screenshot": screenshot,
            })

        finally:
            criar_log_empresa(empresa, nome_empresa, inativados, pulados, erros)
            agente.fechar_sistema_para_loop(CFG.esc_loop_qtd, CFG.esc_loop_intervalo)

    LOG.info("[OK] Etapa 2 finalizada.")


if __name__ == "__main__":
    processar()