# ⚡ Quick Reference - Agente RPA v2.0

Tudo que você precisa saber em 2 minutos.

---

## Importações Mais Comuns

```python
# Config
from agente_rpa.config import cfg
config = cfg()

# Logger
from agente_rpa.core import log
logger = log()

# Automação
from agente_rpa.automacao.base import AutomacaoDominio
automacao = AutomacaoDominio()

# Domínio
from agente_rpa.dominio.modelos import Empresa, Acumulador
empresa = Empresa(codigo=1234, nome="XYZ")

# Exceções
from agente_rpa.core import SemAcumuladoresException, ParadaManualException

# Retry
from agente_rpa.utils.retry import com_retry

# Excel
from agente_rpa.utils.excel import carregar_empresas_com_acumuladores
```

---

## Ações Principais

### Automação
```python
automacao.click(x, y, tempo_espera=1.0)
automacao.write("texto")
automacao.press("enter")
automacao.hotkey("ctrl", "a")
automacao.clear_field()
automacao.find_image("imagem.png")
automacao.click_image_or_fallback("img.png", 100, 200)
automacao.ocr_region((0, 0, 1920, 1080))
automacao.take_screenshot("nome")
automacao.debug_screenshot("debug")
```

### Logger
```python
log().debug("Mensagem debug")
log().info("Informação")
log().warning("Aviso")
log().error("Erro")
log().critical("Crítico")
```

### Retry
```python
@com_retry(tentativas=3, espera=2.0, descricao="Descrição")
def minha_funcao():
    # Será executada com retry automático
    pass

minha_funcao()  # Sem lambda, sem try/except
```

### Estado
```python
estado = EstadoAgente()

# Registrar
estado.registrar_empresa(codigo, nome)
estado.registrar_acumulador(codigo)
estado.registrar_acao(AcaoAgente.INATIVAR_ACUMULADOR)
estado.registrar_erro("Erro aqui", tipo="tipo", recuperavel=True)
estado.registrar_sucesso_acumulador("inativado")

# Consultar
print(estado.tela_atual)  # TelaAtual.LISTAGEM_EMPRESAS
print(estado.empresa_atual.codigo)
print(estado.tentativas_acao_atual)
print(estado.resumo_execucao())
```

### Modelos
```python
# Empresa
empresa = Empresa(codigo=1234, nome="Empresa A")
empresa.acumuladores.append(Acumulador(codigo=42))

empresa.quantidade_acumuladores  # 1
empresa.quantidade_para_inativar  # Aplicou regras
empresa.percentual_processado  # 0.0
empresa.tem_acumuladores_pendentes  # True
empresa.obter_acumuladores_para_inativar()  # [...]

# Acumulador
acc = Acumulador(codigo=42, deve_inativar="SIM")
acc.eh_para_inativar  # True
acc.processado  # False
acc.marcar_sucesso("inativado")
acc.marcar_erro("Campo não encontrado")
```

---

## Fluxo Básico

```python
from agente_rpa.config import cfg
from agente_rpa.core import log, EstadoAgente
from agente_rpa.automacao.base import AutomacaoDominio
from agente_rpa.utils.excel import carregar_empresas_com_acumuladores

# 1. Inicializar
config = cfg()
logger = log()
automacao = AutomacaoDominio()
estado = EstadoAgente()

# 2. Carregar dados
empresas = carregar_empresas_com_acumuladores()

# 3. Processar
for empresa in empresas:
    estado.registrar_empresa(empresa.codigo, empresa.nome)
    
    for acumulador in empresa.obter_acumuladores_para_inativar():
        try:
            estado.registrar_acumulador(acumulador.codigo)
            
            # Sua lógica aqui
            automacao.click(100, 200)
            automacao.write(str(acumulador.codigo))
            automacao.press("enter")
            
            estado.registrar_sucesso_acumulador("inativado")
            logger.info(f"✓ Acumulador {acumulador.codigo} inativado")
            
        except Exception as e:
            estado.registrar_erro(str(e))
            logger.error(f"✗ Erro: {e}")

# 4. Relatório
logger.info(estado.resumo_execucao())
```

---

## Estrutura de Pastas

```
agente_rpa/
├── config/
│   ├── __init__.py
│   └── settings.py          # Edite aqui para customizar
│
├── core/
│   ├── __init__.py
│   ├── logger.py            # Não edite (singleton)
│   ├── estado.py            # Não edite (máquina de estados)
│   └── exceções.py          # Adicione exceções aqui
│
├── automacao/
│   ├── __init__.py
│   └── base.py              # Não edite (baixo nível)
│
├── dominio/
│   ├── __init__.py
│   └── modelos.py           # Estenda com seus modelos
│
├── utils/
│   ├── __init__.py
│   ├── retry.py             # Não edite
│   └── excel.py             # Customize leitura aqui
│
└── workflows/               # Crie seus workflows aqui
    └── __init__.py
```

---

## Testar Rápido

```bash
# Teste se tudo está importando
python -c "
from agente_rpa.config import cfg
from agente_rpa.core import log
from agente_rpa.automacao.base import AutomacaoDominio

print('✓ Config:', cfg().faixa_inativar_min)
print('✓ Logger:', log())
print('✓ Automação:', AutomacaoDominio())
"
```

---

## Erros Comuns

| Erro | Solução |
|------|---------|
| `ImportError: agente_rpa` | `sys.path.insert(0, str(Path(__file__).parent))` |
| `FileNotFoundError: config` | `cfg().criar_pastas()` |
| `Logger não cria arquivo` | Verifique `Relatorio Final/` existe |
| `Template não encontrado` | Verifique `imagens_rpa/` tem arquivo |
| `Tesseract not found` | Instale Tesseract ou desative OCR |

---

## Arquivos Documentação

| Arquivo | Para quê |
|---------|----------|
| `ARQUITETURA.md` | Entender design completo |
| `DIAGRAMAS.md` | Ver fluxos e estrutura |
| `MIGRACAO.md` | Migrar código antigo |
| `exemplo_uso.py` | Ver código executável |
| `QUICK_REFERENCE.md` | Este arquivo |

---

## Checklist Novo Projeto

- [ ] Config customizada em `agente_rpa/config/settings.py`
- [ ] Pasta `agente_rpa/imagens_rpa/` com screenshots
- [ ] Teste com `python -c "from agente_rpa import cfg; print(cfg())"`
- [ ] Crie seu `workflows/seu_workflow.py`
- [ ] Teste com `exemplo_uso.py`
- [ ] Migre seu código gradualmente

---

## Links Rápidos

- **Arquivo Principal**: `agente_rpa/config/settings.py`
- **Logger Global**: `agente_rpa/core/logger.py`
- **Estado Máquina**: `agente_rpa/core/estado.py`
- **Automação**: `agente_rpa/automacao/base.py`
- **Exemplo Executável**: `exemplo_uso.py`

---

## Última Dica

Você não está em "criar um agente". Você está em **transformar automação em agente profissional**.

Isso é muito mais valioso que 99% dos projetos "IA" por aí.

Continua assim. 🚀
