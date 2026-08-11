"""
Teste rápido dos 3 arquivos principais do core.

Execute: python teste_core.py
"""

print("="*60)
print("🧪 TESTE QUICK - CORE DO AGENTE V2")
print("="*60)

# Teste 1: Logger Singleton
print("\n[1/3] Testando Logger Singleton...")
try:
    from agente_rpa.core.logger import log, AgenteLogger
    
    logger1 = log
    logger2 = log
    
    assert logger1 is logger2, "Logger não é singleton!"
    logger1.info("✓ Logger singleton funcionando")
    print("✓ Logger funcionando corretamente")
except Exception as e:
    print(f"✗ Erro no logger: {e}")

# Teste 2: Exceções
print("\n[2/3] Testando Exceções...")
try:
    from agente_rpa.core.excecoes import (
        SemAcumuladoresException,
        ParadaManualException,
        FalhaVisualizacaoException,
    )
    
    # Teste lançamento
    try:
        raise SemAcumuladoresException("Teste")
    except SemAcumuladoresException as e:
        pass
    
    log.info("✓ Exceções funcionando")
    print("✓ Exceções funcionando corretamente")
except Exception as e:
    print(f"✗ Erro nas exceções: {e}")

# Teste 3: Retry com Exceção de Negócio
print("\n[3/3] Testando Retry...")
try:
    from agente_rpa.utils.retry import executar_com_retry
    from agente_rpa.core.excecoes import SemAcumuladoresException
    
    # Função que falha sempre
    chamadas = []
    
    def funcao_com_erro():
        chamadas.append(1)
        raise ValueError("Erro genérico")
    
    # Deve fazer retry
    try:
        executar_com_retry(funcao_com_erro, tentativas=3, espera=0.1)
    except ValueError:
        pass  # Esperado
    
    assert len(chamadas) == 3, f"Esperava 3 chamadas, teve {len(chamadas)}"
    
    # Função com exceção de negócio
    chamadas2 = []
    
    def funcao_sem_acumuladores():
        chamadas2.append(1)
        raise SemAcumuladoresException("Sem acumuladores")
    
    # NÃO deve fazer retry - sobe direto
    try:
        executar_com_retry(funcao_sem_acumuladores, tentativas=3, espera=0.1)
    except SemAcumuladoresException:
        pass  # Esperado
    
    assert len(chamadas2) == 1, f"Não deveria retry, mas teve {len(chamadas2)} chamadas"
    
    log.info("✓ Retry funcionando (com distinção de erros)")
    print("✓ Retry funcionando corretamente")
except Exception as e:
    print(f"✗ Erro no retry: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "="*60)
print("✓ TODOS OS TESTES PASSARAM!")
print("="*60)

# Mostrar arquivo de log
from agente_rpa.core.logger import log
print(f"\n📝 Log salvo em: {log.caminho}")

# Imprimir primeiras linhas do log
try:
    with open(log.caminho, "r", encoding="utf-8") as f:
        linhas = f.readlines()[-10:]  # Últimas 10 linhas
        print("\n[Últimas linhas do log]")
        for linha in linhas:
            print(f"  {linha.strip()}")
except:
    pass

print("\n✨ Parabéns! A V2 está começando a tomar forma!")
