"""
EXEMPLO DE USO: Como usar a nova arquitetura do agente RPA.

Este arquivo demonstra como a nova arquitetura MUDA sua forma de programar.

ANTES (código procedural):
    clicar_em_xyz()
    esperar(5)
    se_isso_entao_aquilo()
    escrever_texto()
    clicar_em_abc()

DEPOIS (código com estado + decisão):
    agente = AgentePrincipal()
    for empresa in empresas:
        resultado = agente.processar_empresa(empresa)
        print(resultado)
"""

from agente_rpa.config import cfg, obter_config
from agente_rpa.core import log, EstadoAgente, TelaAtual, AcaoAgente
from agente_rpa.automacao.base import AutomacaoDominio
from agente_rpa.dominio.modelos import Empresa, Acumulador
from agente_rpa.utils.excel import carregar_empresas_com_acumuladores
from agente_rpa.utils.retry import com_retry


class AgentePrincipal:
    """
    Orquestrador principal do agente.
    
    Responsável por:
    - Carregar empresas e acumuladores
    - Manter estado centralizado
    - Decidir próxima ação
    - Executar workflows
    - Tratar erros
    """

    def __init__(self):
        """Inicializa o agente."""
        self.config = cfg()
        self.logger = log()
        self.automacao = AutomacaoDominio()
        self.estado = EstadoAgente()

        self.logger.info("="*60)
        self.logger.info("🤖 AGENTE RPA INICIALIZADO")
        self.logger.info("="*60)

    # ============================================================
    # DECISÃO: Para onde vou agora?
    # ============================================================

    def decidir_proxima_acao(self) -> AcaoAgente:
        """
        Motor de decisão: com base no estado, qual é a próxima ação?
        
        Este é o "cérebro" do agente. Aqui entra IA depois (LLM).
        """
        # Se está em error recovery
        if self.estado.erro_ultimo and self.estado.tentativas_acao_atual < self.config.retry_tentativas:
            if self.estado.recuperavel:
                return AcaoAgente.RECUPERAR_ERRO
            else:
                return AcaoAgente.REINICIAR_FLUXO

        # Se deve parar
        if self.estado.deve_parar or self.estado.parada_manual:
            return AcaoAgente.PARAR

        # Se não tem empresa atual
        if not self.estado.empresa_atual:
            return AcaoAgente.ABRIR_EMPRESA

        # Se não tem acumulador atual
        if not self.estado.acumulador_atual:
            return AcaoAgente.LISTAR_ACUMULADORES

        # Se empresa ainda tem acumuladores
        if self.estado.empresa_atual.tem_acumuladores_pendentes:
            return AcaoAgente.INATIVAR_ACUMULADOR

        # Se empresa acabou
        return AcaoAgente.ABRIR_EMPRESA

    # ============================================================
    # EXECUÇÃO: Implementar cada ação
    # ============================================================

    @com_retry(tentativas=3, espera=2.0, descricao="Abrir empresa")
    def executar_abrir_empresa(self, codigo_empresa: int) -> bool:
        """Abre uma empresa no sistema."""
        self.logger.info(f"📂 Abrindo empresa {codigo_empresa}...")
        self.estado.registrar_empresa(codigo_empresa, f"Empresa {codigo_empresa}")

        # Simular abertura (seu código aqui)
        # self.automacao.press("f8")
        # self.automacao.write(str(codigo_empresa))
        # self.automacao.press("enter")

        self.estado.mudar_tela(TelaAtual.LISTAGEM_EMPRESAS)
        return True

    @com_retry(tentativas=2, espera=1.5, descricao="Listar acumuladores")
    def executar_listar_acumuladores(self) -> bool:
        """Lista acumuladores de uma empresa."""
        self.logger.info(f"📋 Listando acumuladores...")

        # Simular listagem
        # self.automacao.press("alt")
        # self.automacao.press("a")

        self.estado.mudar_tela(TelaAtual.LISTAGEM_ACUMULADORES)
        return True

    @com_retry(tentativas=3, espera=2.0, descricao="Inativar acumulador")
    def executar_inativar_acumulador(self, acumulador: Acumulador) -> bool:
        """Inativa um acumulador específico."""
        if not self.estado.acumulador_atual:
            return False

        self.logger.info(
            f"⚙️  Inativando acumulador {acumulador.codigo} "
            f"(empresa {acumulador.empresa_codigo})"
        )

        try:
            # Simular inativação
            # self.automacao.click(1310, 406)  # Clicar em acumulador
            # self.automacao.write(str(acumulador.codigo))
            # self.automacao.press("enter")

            # Se chegou aqui, sucesso
            self.estado.registrar_sucesso_acumulador("inativado")
            return True

        except Exception as e:
            self.estado.registrar_erro(str(e), tipo="inativacao")
            raise

    def executar_recuperar_erro(self) -> bool:
        """Tenta recuperar de um erro."""
        self.logger.warning(f"🔄 Recuperando de erro: {self.estado.erro_ultimo}")
        self.automacao.close_windows_with_esc(quantidade=6, intervalo=0.2)
        return True

    def executar_reiniciar_fluxo(self) -> bool:
        """Reinicia o fluxo de uma empresa."""
        self.logger.error(f"⚠️  Reiniciando fluxo...")
        self.automacao.close_windows_with_esc(quantidade=10, intervalo=0.3)
        self.estado.erro_ultimo = None
        self.estado.tentativas_acao_atual = 0
        return True

    # ============================================================
    # ORQUESTRAÇÃO PRINCIPAL
    # ============================================================

    def processar_empresa(self, empresa: Empresa) -> dict:
        """
        Processa uma empresa completamente.
        
        Isso é um exemplo simplificado. Seu código será muito mais complexo.
        O ponto é: agora você tem ESTADO + DECISÃO + AÇÃO estruturados.
        """
        self.logger.info(f"\n{'='*60}")
        self.logger.info(f"🏢 INICIANDO PROCESSAMENTO: {empresa.nome} ({empresa.codigo})")
        self.logger.info(f"{'='*60}")

        empresa.iniciar_processamento()
        self.estado.registrar_empresa(empresa.codigo, empresa.nome)

        try:
            # Abrir empresa
            self.executar_abrir_empresa(empresa.codigo)
            self.automacao.wait(1.0)

            # Listar acumuladores
            self.executar_listar_acumuladores()
            self.automacao.wait(1.0)

            # Processar cada acumulador
            acumuladores_para_inativar = [
                a for a in empresa.acumuladores if a.eh_para_inativar
            ]

            for i, acumulador in enumerate(acumuladores_para_inativar, 1):
                self.logger.info(
                    f"  [{i}/{len(acumuladores_para_inativar)}] "
                    f"Processando acumulador {acumulador.codigo}..."
                )

                self.estado.registrar_acumulador(acumulador.codigo)

                try:
                    self.executar_inativar_acumulador(acumulador)
                    self.logger.info(f"    ✓ Acumulador {acumulador.codigo} inativado")

                except Exception as e:
                    self.logger.error(f"    ✗ Falha: {e}")
                    acumulador.marcar_erro(str(e))

                self.automacao.wait(0.5)

            empresa.concluir_processamento()
            self.logger.info(f"✓ Empresa {empresa.codigo} concluída\n")

        except Exception as e:
            self.logger.error(f"✗ Erro crítico em empresa {empresa.codigo}: {e}")
            empresa.marcar_erro(str(e))

        # Retornar resumo
        return empresa.resumo()

    def processar_todas_empresas(self) -> None:
        """Processa todas as empresas."""
        self.logger.info("\n" + "="*60)
        self.logger.info("📊 INICIANDO PROCESSAMENTO GERAL")
        self.logger.info("="*60 + "\n")

        # Carregar dados
        empresas = carregar_empresas_com_acumuladores()

        if not empresas:
            self.logger.error("Nenhuma empresa carregada!")
            return

        self.logger.info(f"Carregadas {len(empresas)} empresas para processar\n")

        # Processar cada empresa
        for i, empresa in enumerate(empresas, 1):
            self.logger.info(f"\n[{i}/{len(empresas)}]")
            resultado = self.processar_empresa(empresa)

            # Mostrar resumo
            self.logger.info(f"  Resumo: {resultado}\n")

            # Verificar parada manual
            if self.config.arquivo_parada_manual.exists():
                self.logger.warning("⛔ Arquivo de parada manual detectado. Parando...")
                break

        # Relatório final
        self._imprimir_relatorio_final()

    def _imprimir_relatorio_final(self) -> None:
        """Imprime relatório final de execução."""
        resumo = self.estado.resumo_execucao()

        self.logger.info("\n" + "="*60)
        self.logger.info("📈 RELATÓRIO FINAL DE EXECUÇÃO")
        self.logger.info("="*60)
        self.logger.info(f"Empresas processadas: {resumo['empresas_totais']}")
        self.logger.info(f"Acumuladores inativados: {resumo['acumuladores_inativados']}")
        self.logger.info(f"Acumuladores com erro: {resumo['acumuladores_com_erro']}")
        self.logger.info(f"Tempo decorrido: {resumo['tempo_decorrido_minutos']:.1f} minutos")
        self.logger.info("="*60 + "\n")


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    try:
        agente = AgentePrincipal()
        agente.processar_todas_empresas()

    except KeyboardInterrupt:
        print("\n⛔ Interrompido pelo usuário")
    except Exception as e:
        print(f"❌ Erro fatal: {e}")
        import traceback
        traceback.print_exc()
