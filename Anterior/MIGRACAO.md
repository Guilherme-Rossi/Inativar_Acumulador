# 📚 Guia de Migração: v1.0 → v2.0

Como migrar seu código antigo para a nova arquitetura, passo a passo.

---

## Opção A: Migração Gradual (Recomendado)

Você não precisa reescrever tudo de uma vez. Faça incrementalmente.

### Passo 1: Preparar Ambiente

```bash
# 1. Criar pasta agente_rpa/ com a nova estrutura
# Já feito em: c:\...\ P1 - Inativar Acumulador\agente_rpa\

# 2. Instalar dependências (se necessário)
pip install pydantic pydantic-settings
```

### Passo 2: Usar Config Centralizada

**Antes (duplicação):**
```python
# main.py
class Config:
    base_dir = Path.home() / "Desktop" / "..."
    faixa_inativar_min = 1
    tempo_dominio = 10.5

# etapa1_validacao.py
class Config:
    base_dir = Path(r"C:\Users\...")
    faixa_min = 212
    tempo_dominio = 10.5

# etapa2_inativacao.py
class Config:
    base_dir = Path(r"C:\Users\...")
    tempo_dominio = 11.0
```

**Depois (centralizado):**
```python
# Em TODOS os arquivos:
from agente_rpa.config import cfg

config = cfg()
print(config.faixa_inativar_min)  # 1
print(config.tempo_dominio)        # 10.5
```

**Migração:**
```python
# Seu arquivo antigo:
from pathlib import Path

@dataclass(frozen=True)
class Config:
    base_dir: Path = Path.home() / "Desktop" / "..."
    faixa_inativar_min: int = 1
    # ... mais campos


# Novo arquivo:
from agente_rpa.config import cfg

# Use:
config = cfg()

# E remova sua classe Config local
# Se precisar customizar, edite agente_rpa/config/settings.py
```

### Passo 3: Usar Logger Centralizado

**Antes (duplicação):**
```python
# main.py
class Logger:
    def __init__(self):
        self.caminho = CFG.pasta_logs / "execucao_geral.log"
    def info(self, msg):
        print(msg)
        with open(self.caminho, "a") as f:
            f.write(msg + "\n")

LOG = Logger()

# etapa2_inativacao.py
class Logger:
    def __init__(self):
        self.log_path = CFG.pasta_logs / "etapa2_inativacao.log"
    def info(self, msg):
        print(msg)
        with open(self.log_path, "a") as f:
            f.write(msg + "\n")

LOG = Logger()
```

**Depois (centralizado):**
```python
# Em TODOS os arquivos:
from agente_rpa.core import log

logger = log()
logger.info("Mensagem importante")
```

**Migração:**
```python
# Seu código antigo usa LOG.info()
# Substitua todas as ocorrências:

# Antes:
LOG.info("Começando...")

# Depois:
from agente_rpa.core import log
log().info("Começando...")
```

### Passo 4: Usar Decorador @com_retry

**Antes (retry manual):**
```python
def executar_com_retry(func, tentativas=3, espera=2.0, descricao=""):
    ultimo_erro = None
    for tentativa in range(1, tentativas + 1):
        try:
            return func()
        except Exception as exc:
            ultimo_erro = exc
            LOG.info(f"[RETRY] {descricao} falhou: {exc}")
            if tentativa < tentativas:
                time.sleep(espera)
    raise ultimo_erro

# Uso
def buscar_acumulador(codigo):
    # ... seu código

resultado = executar_com_retry(
    lambda: buscar_acumulador(42),
    tentativas=3,
    espera=2.0,
    descricao="Buscar acumulador 42"
)
```

**Depois (decorador automático):**
```python
from agente_rpa.utils.retry import com_retry

@com_retry(tentativas=3, espera=2.0, descricao="Buscar acumulador")
def buscar_acumulador(codigo):
    # ... seu código

# Uso - sem lambda, sem try/except manual!
resultado = buscar_acumulador(42)
```

### Passo 5: Usar Automação Estruturada

**Antes:**
```python
from agente_dominio import AgenteConfig, AgenteDominio

agente_cfg = AgenteConfig(
    pasta_imagens=CFG.pasta_imagens,
    tempo_padrao=CFG.tempo_dominio,
    tempo_max=CFG.tempo_dominio_max,
)
agente = AgenteDominio(agente_cfg)

# Uso
agente.clicar_coordenada(100, 200, 1.0)
agente.escrever("texto")
agente.pressionar("enter")
```

**Depois:**
```python
from agente_rpa.automacao.base import AutomacaoDominio

automacao = AutomacaoDominio()  # Configuração automática!

# Uso
automacao.click(100, 200, tempo_espera=1.0)
automacao.write("texto")
automacao.press("enter")
```

**Migração:**
```python
# Seu arquivo antigo:
agente.clicar_coordenada(x, y, seg)
agente.escrever(texto, seg)
agente.pressionar(tecla, seg)
agente.localizar_imagem(imagem)

# Novo:
automacao.click(x, y, tempo_espera=seg)
automacao.write(texto, tempo_espera=seg)
automacao.press(tecla, tempo_espera=seg)
automacao.find_image(imagem)
```

### Passo 6: Usar Modelos de Domínio

**Antes (dicionários):**
```python
# Dados espalhados em listas de dicts
empresas = [
    {"CODIGO": 1234, "NOME": "Empresa A"},
    {"CODIGO": 5678, "NOME": "Empresa B"},
]

acumuladores = [
    {"ACUMULADOR": 42, "DEVE_INATIVAR": "SIM"},
    {"ACUMULADOR": 99, "DEVE_INATIVAR": "NAO"},
]

# Sem validação, sem métodos
for emp in empresas:
    codigo = int(emp["CODIGO"])  # Precisa converter sempre
```

**Depois (classes):**
```python
from agente_rpa.dominio.modelos import Empresa, Acumulador

empresas = [
    Empresa(codigo=1234, nome="Empresa A"),
    Empresa(codigo=5678, nome="Empresa B"),
]

acumuladores = [
    Acumulador(codigo=42, deve_inativar="SIM"),
    Acumulador(codigo=99, deve_inativar="NAO"),
]

# Com métodos úteis e validação automática
for emp in empresas:
    print(emp.codigo)  # Tipo int garantido
    print(emp.quantidade_acumuladores)  # Propriedade calculada
    if emp.tem_acumuladores_pendentes:  # Método inteligente
        processar(emp)
```

### Passo 7: Usar Estado Centralizado

**Antes (sem estado):**
```python
# Estado espalhado em variáveis globais
empresa_atual = None
acumulador_atual = None
tentativas = 0
ultimo_erro = None

def processar():
    global empresa_atual, tentativas
    
    empresa_atual = empresas[0]
    tentativas = 0
    
    # ... código
    
    if erro:
        tentativas += 1
        if tentativas > 3:
            # Qual era o erro mesmo?
            # Qual empresa estamos?
```

**Depois (estado centralizado):**
```python
from agente_rpa.core import EstadoAgente, TelaAtual, AcaoAgente

estado = EstadoAgente()

def processar():
    # Claro onde estamos
    estado.registrar_empresa(empresa.codigo, empresa.nome)
    estado.registrar_acumulador(acumulador.codigo)
    
    try:
        # ... código
    except Exception as e:
        estado.registrar_erro(str(e), tipo="inativacao", recuperavel=True)
        # Estado já sabe tudo: empresa, acumulador, tentativas, erro
        
    # Saber o que fazer agora
    proxima_acao = decidir_proxima_acao(estado)
```

### Passo 8: Integrar com Excel Novo

**Antes:**
```python
def carregar_empresas(caminho):
    df = pd.read_excel(caminho)
    # ... processamento manual
    return lista_dicts

# Seu código
empresas = carregar_empresas(CFG.planilha_empresas)
```

**Depois:**
```python
from agente_rpa.utils.excel import carregar_empresas_com_acumuladores

# Um comando carrega TUDO + aplica regras
empresas = carregar_empresas_com_acumuladores()

# Usar como objetos, não dicts
for empresa in empresas:
    print(empresa.nome)
    print(empresa.quantidade_para_inativar)  # Regra aplicada!
```

---

## Opção B: Rewrite Completo

Se preferir fazer tudo de uma vez (mais rápido, menos incremental):

### 1. Criar novo main.py

```python
# novo_main.py

from agente_rpa.config import cfg
from agente_rpa.core import log
from agente_rpa.automacao.base import AutomacaoDominio
from agente_rpa.dominio.modelos import Empresa, Acumulador
from agente_rpa.utils.excel import carregar_empresas_com_acumuladores
from exemplo_uso import AgentePrincipal

if __name__ == "__main__":
    # Inicializar
    config = cfg()
    logger = log()
    
    logger.info("Iniciando agente...")
    
    # Executar
    agente = AgentePrincipal()
    agente.processar_todas_empresas()
    
    logger.info("✓ Execução concluída")
```

### 2. Renomear arquivos antigos

```bash
# Backup
mv main.py main_old_v1.py
mv etapa1_validacao.py etapa1_validacao_old_v1.py
mv etapa2_inativacao.py etapa2_inativacao_old_v1.py
```

### 3. Usar novo_main.py como principal

```bash
python novo_main.py
```

---

## Testes

### Teste 1: Config

```python
from agente_rpa.config import cfg

c = cfg()
assert c.faixa_inativar_min == 1
assert c.tempo_dominio == 10.5
print("✓ Config OK")
```

### Teste 2: Logger

```python
from agente_rpa.core import log

l = log()
l.info("Teste de logger")
# Verifique arquivo em Relatorio Final/agente.log
print("✓ Logger OK")
```

### Teste 3: Automação (sem GUI)

```python
from agente_rpa.automacao.base import AutomacaoDominio

a = AutomacaoDominio()
# Não vai clicar em nada, mas deve inicializar
print("✓ Automação OK")
```

### Teste 4: Modelos

```python
from agente_rpa.dominio.modelos import Empresa, Acumulador

e = Empresa(codigo=1234, nome="Teste")
a = Acumulador(codigo=42, empresa_codigo=1234)
e.acumuladores.append(a)

assert e.quantidade_acumuladores == 1
print("✓ Modelos OK")
```

### Teste 5: Retry

```python
from agente_rpa.utils.retry import com_retry

@com_retry(tentativas=3, espera=0.1)
def funcao_que_falha():
    raise ValueError("Erro intencional")

try:
    funcao_que_falha()
except ValueError:
    print("✓ Retry OK (falhou depois de 3 tentativas)")
```

---

## Checklist de Migração

- [ ] Backup dos arquivos antigos (main_old_v1.py, etc)
- [ ] Pasta agente_rpa/ criada com estrutura
- [ ] Config migrada
- [ ] Logger migrado  
- [ ] Automação migrada
- [ ] Modelos implementados
- [ ] Utils/Excel integrado
- [ ] Estado implementado
- [ ] Testes passando
- [ ] Novo main.py executando
- [ ] Código antigo removido ou archived

---

## Troubleshooting

### ImportError: No module named 'agente_rpa'

**Solução:**
```python
# Adicione no início do arquivo
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

# Ou execute de dentro do diretório do projeto
```

### Config não encontra arquivos

**Solução:**
```python
from agente_rpa.config import cfg

c = cfg()
c.base_dir.mkdir(parents=True, exist_ok=True)
# A config já cria as pastas
```

### Logger não cria arquivo

**Solução:**
```python
from agente_rpa.core import log

l = log()
# Arquivo é criado em: base_dir/Relatorio Final/agente.log
# Verifique se a pasta existe
```

---

## Próximos Passos

1. ✅ Escolha uma opção (gradual ou rewrite)
2. ✅ Execute a migração
3. ✅ Teste cada passo
4. ✅ Verifique que o comportamento é igual
5. 🔄 Refine a arquitetura conforme necessário
6. 📈 Comece a adicionar workflows novos

---

**Sua migração está bem estruturada.**

Não é apenas "refatoração". É transformação de um script em **agente profissional**.
