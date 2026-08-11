# 📊 Diagramas da Arquitetura

## 1. Estrutura de Módulos

```
agente_rpa/
├── config/settings.py ─────────────────────────┐
│   (Pydantic, singleton global)               │
│                                               │
├── core/                                        │
│   ├── logger.py ──────────────────────────────┼──> LoggerAgente (singleton)
│   ├── estado.py ──────────────────────────────┼──> EstadoAgente (máquina)
│   └── exceções.py ────────────────────────────┼──> Erros estruturados
│                                               │
├── automacao/                                   │
│   └── base.py ────────────────────────────────┼──> AutomacaoDominio
│       (click, write, press, ocr, imagens)    │
│                                               │
├── dominio/                                     │
│   └── modelos.py ─────────────────────────────┼──> Empresa, Acumulador
│                                               │
├── utils/                                       │
│   ├── retry.py ───────────────────────────────┼──> @com_retry decorator
│   └── excel.py ───────────────────────────────┼──> Leitura de dados
│                                               │
└── workflows/ ─────────────────────────────────┼──> (Próximos: validação, inativação)
                                                │
Como tudo se integra:                           │
AgentePrincipal (exemplo_uso.py) ─────────────┘
```

## 2. Fluxo de Dados

```
┌─────────────────────────────────────────────────────────────┐
│ Carregar Dados (Excel)                                      │
│ ↓                                                            │
│ .carregar_empresas_com_acumuladores()                       │
│   └──> Empresa + Acumuladores + Regras de Negócio          │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│ LOOP para cada Empresa:                                     │
│                                                             │
│ ┌───────────────────────────────────────────────────────┐   │
│ │ 1. Registrar Estado                                  │   │
│ │    estado.registrar_empresa()                        │   │
│ │                                                       │   │
│ │ 2. Decidir Próxima Ação                              │   │
│ │    ação = decidir_proxima_acao()                     │   │
│ │    └──> Máquina de Estados                           │   │
│ │                                                       │   │
│ │ 3. Executar Ação com Retry Automático                │   │
│ │    @com_retry(tentativas=3)                          │   │
│ │    executar_abrir_empresa() / inativar / etc         │   │
│ │                                                       │   │
│ │ 4. Registrar Resultado                               │   │
│ │    estado.registrar_sucesso()                        │   │
│ │    ou                                                │   │
│ │    estado.registrar_erro() → Retry automático        │   │
│ │                                                       │   │
│ └───────────────────────────────────────────────────────┘   │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│ Finalizar Execução                                          │
│ ↓                                                            │
│ Imprimir Relatório Final                                    │
│ (empresas processadas, acumuladores inativados, erros)      │
└─────────────────────────────────────────────────────────────┘
```

## 3. Máquina de Estados

```
                    TELA_INICIAL
                         │
                         ▼
                   LISTAGEM_EMPRESAS
                         │
                    ┌────┴────┬──────────────────┐
                    ▼         │                  │
            LISTAGEM_         │ (sem acumuladores)
            ACUMULADORES      │
                    │         │                  │
                    ├─────────┴──────────────────┘
                    │
            LOOP para cada acumulador:
            ┌───────────┬──────────────────┐
            │           │                  │
            ▼           ▼                  ▼
    DETALHE_           MODAL_          POPUP_
    ACUMULADOR         ERRO            DESCONHECIDO
            │           │                  │
            └──────┬────┴──────┬───────────┘
                   │           │
                   ▼           ▼
            [SUCESSO]    [RECUPERAÇÃO]
                   │           │
                   └─────┬─────┘
                         │
                    (próximo acumulador)
                         │
                    LISTAGEM_ACUMULADORES
                         │
                    (próxima empresa)
                         │
                    LISTAGEM_EMPRESAS
```

## 4. Camadas da Aplicação

```
┌────────────────────────────────────────────────────────────┐
│ CAMADA DE NEGÓCIO (Workflows)                              │
│ ┌──────────────────────────────────────────────────────┐   │
│ │ WorkflowValidacao  WorkflowInativacao  etc          │   │
│ │ (Etapa1)           (Etapa2)                         │   │
│ └──────────────────────────────────────────────────────┘   │
│              ↑              ↑              ↑                │
├─────────────────────────────────────────────────────────────┤
│ CAMADA COGNITIVA (Decisão + Estado)                        │
│ ┌──────────────────────────────────────────────────────┐   │
│ │ AgentePrincipal (Orquestrador)                      │   │
│ │ + EstadoAgente (Máquina de Estados)                 │   │
│ │ + decidir_proxima_acao() (Motor de Decisão)        │   │
│ └──────────────────────────────────────────────────────┘   │
│              ↑              ↑              ↑                │
├─────────────────────────────────────────────────────────────┤
│ CAMADA OPERACIONAL (Ações)                                 │
│ ┌──────────────────────────────────────────────────────┐   │
│ │ AutomacaoDominio (Mouse, Teclado, Visão)           │   │
│ │  • click()  • write()  • ocr_region()               │   │
│ │  • find_image()  • take_screenshot()                │   │
│ └──────────────────────────────────────────────────────┘   │
│              ↑              ↑              ↑                │
├─────────────────────────────────────────────────────────────┤
│ CAMADA DE SUPORTE (Config + Logging + Utils)               │
│ ┌──────────────────────────────────────────────────────┐   │
│ │ Config  Logger  Exceções  Retry  Excel              │   │
│ └──────────────────────────────────────────────────────┘   │
│              ↑              ↑              ↑                │
└─────────────────────────────────────────────────────────────┘
```

## 5. Diferença v1 vs v2

```
┌─────────────────────────────────────────────────────────────┐
│ V1.0: SCRIPT PROCEDURAL (ANTES)                             │
│                                                             │
│ def main():                                                │
│     for empresa in empresas:                              │
│         abrir_empresa(empresa)                            │
│         for acumulador in acumuladores:                   │
│             try:                                          │
│                 inativar_acumulador(acumulador)           │
│             except Exception as e:                        │
│                 if retry < 3:                            │
│                     retry += 1                           │
│                     retry_inativar()                      │
│                     ...                                   │
│                                                             │
│ Problemas:                                                │
│ ❌ Lógica espalhada                                        │
│ ❌ Sem estado centralizado                                │
│ ❌ Retry manual em vários lugares                         │
│ ❌ Logs inconsistentes                                    │
│ ❌ Config duplicada (3 arquivos)                          │
│ ❌ Difícil de testar                                      │
│ ❌ Difícil de estender                                    │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│ V2.0: AGENTE ESTRUTURADO (DEPOIS)                           │
│                                                             │
│ class AgentePrincipal:                                     │
│     def decidir_proxima_acao(self):                        │
│         if self.estado.erro:                              │
│             return RECUPERAR_ERRO                          │
│         elif self.estado.empresa.acumuladores_pendentes:  │
│             return INATIVAR_ACUMULADOR                     │
│         else:                                              │
│             return ABRIR_EMPRESA                           │
│                                                             │
│     @com_retry(tentativas=3)                              │
│     def executar_inativar_acumulador(self, acum):         │
│         self.automacao.click(...)                          │
│         self.estado.registrar_sucesso()                    │
│                                                             │
│ Benefícios:                                                │
│ ✅ Lógica centralizada                                    │
│ ✅ Estado centralizado (máquina)                          │
│ ✅ Retry automático (@decorator)                          │
│ ✅ Logs estruturados                                      │
│ ✅ Config única centralizada                              │
│ ✅ Fácil de testar (injeção)                              │
│ ✅ Fácil de estender (workflows)                          │
│ ✅ Pronto para IA (LLM)                                   │
└─────────────────────────────────────────────────────────────┘
```

## 6. Integração com IA (Futuro)

```
┌──────────────────────────────────────────────────┐
│ Input: Estado Atual + Contexto                   │
│                                                  │
│ estado.tela_atual = "MODAL_ERRO"                │
│ estado.acumulador_atual.codigo = 42             │
│ estado.erro_ultimo = "Campo não encontrado"     │
│                                                  │
└──────────────────────────────────────────────────┘
                      │
                      ▼
        ┌─────────────────────────────┐
        │ LLM (OpenAI, Ollama, etc)   │
        │                             │
        │ "Qual é a melhor ação?"     │
        │                             │
        └─────────────────────────────┘
                      │
                      ▼
┌──────────────────────────────────────────────────┐
│ Output: Ação Recomendada                         │
│                                                  │
│ {                                                │
│   "acao": "RECUPERAR_ERRO",                     │
│   "passos": [                                    │
│     "fechar_popup()",                           │
│     "retry_campo()",                            │
│   ],                                             │
│   "confianca": 0.92                             │
│ }                                                │
│                                                  │
└──────────────────────────────────────────────────┘
                      │
                      ▼
        agente.executar_acao(acao)
```

---

**Visualização em código:**

Ver [exemplo_uso.py](./exemplo_uso.py) para código prático.
Ver [ARQUITETURA.md](./ARQUITETURA.md) para documentação detalhada.
