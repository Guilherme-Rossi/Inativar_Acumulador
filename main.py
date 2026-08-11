import time

from agente_rpa.config.settings import CFG
from agente_rpa.core.logger import log
from agente_rpa.automacao.sistema_dominio import SistemaDominio
from agente_rpa.workflows.inativador import WorkflowInativacao


def main() -> None:
    log.info("=====================================================")
    log.info("🤖 AGENTE DOMÍNIO V2 - INATIVAÇÃO DE ACUMULADORES 🤖")
    log.info("=====================================================")

    operador_dominio = SistemaDominio(
        pasta_imagens=CFG.pasta_imagens,
        arquivo_parada=CFG.arquivo_parada_manual,
    )

    gerente_inativador = WorkflowInativacao(sistema=operador_dominio)

    try:
        gerente_inativador.executar_lote()
        log.info("=== PROCESSO TOTAL FINALIZADO COM SUCESSO ===")
    except Exception as exc:
        log.info(f"=== PROCESSO INTERROMPIDO: {exc} ===")


if __name__ == "__main__":
    main()
