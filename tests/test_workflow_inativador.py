import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from agente_rpa.workflows.inativador import WorkflowInativacao


class WorkflowInativacaoTests(unittest.TestCase):
    def test_garante_relatorio_quando_ausente(self) -> None:
        sistema = MagicMock()
        workflow = WorkflowInativacao(sistema)

        with tempfile.TemporaryDirectory() as tmp_dir:
            caminho_relatorio = Path(tmp_dir) / "RELATORIO.xlsx"
            caminho_relatorio.write_bytes(b"fake")

            with patch("agente_rpa.workflows.inativador.verificar_relatorio_existente", return_value=None), \
                 patch("agente_rpa.workflows.inativador.aguardar_relatorio_empresa", return_value=caminho_relatorio), \
                 patch("agente_rpa.workflows.inativador.processar_relatorio_extraido", return_value=[10, 20, 10]):
                resultado = workflow._carregar_acumuladores_empresa_unicos(123, gerar_se_ausente=True)

            self.assertEqual([10, 20], resultado)
            sistema.gerar_relatorio_acumuladores.assert_called_once_with(123)

    def test_nao_considera_log_antigo_quando_relatorio_esta_ausente(self) -> None:
        workflow = WorkflowInativacao(MagicMock())

        with tempfile.TemporaryDirectory() as tmp_dir:
            pasta_temp = Path(tmp_dir)
            log_path = pasta_temp / "12.txt"
            log_path.write_text(
                "Status Geral da Empresa: OK\n"
                "Checklist de Acumuladores:\n"
                "Acumulador 6 - Inativado\n",
                encoding="utf-8",
            )

            cfg_mock = MagicMock()
            cfg_mock.pasta_inativacao = pasta_temp
            with patch("agente_rpa.workflows.inativador.CFG", cfg_mock), \
                 patch.object(WorkflowInativacao, "_buscar_caminho_relatorio", return_value=None):
                finalizada, _, _ = workflow.analisar_log_existente(12)

            self.assertFalse(finalizada)


if __name__ == "__main__":
    unittest.main()
