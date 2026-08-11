# Agente RPA para inativação de acumuladores

Este projeto automatiza a inativação de acumuladores no sistema Domínio Escrita Fiscal com foco em execução em lote, validação de dados e tratamento de cenários visuais do sistema.

A solução foi organizada em duas etapas principais:

## 1. Extração e validação

Objetivo:
- entrar na empresa no Domínio;
- acessar a tela de acumuladores;
- gerar o relatório da empresa em Excel;
- ler o relatório e identificar quais acumuladores realmente existem naquela empresa;
- preparar o conjunto de dados que será usado na etapa de inativação.

## 2. Inativação

Objetivo:
- localizar o acumulador alvo;
- criar uma nova vigência com situação inativa;
- gravar a alteração;
- tratar popups e mensagens do sistema;
- registrar o resultado em log por empresa.

## Como funciona o fluxo

O agente opera com base em:
- teclado e mouse automatizados;
- reconhecimento visual por imagens;
- OCR para leitura de textos quando necessário;
- tratamento de popups como data, CFOP, vigência, impostos e Simples Nacional.

## Regras impostas pelo fluxo

O processo segue regras explícitas para reduzir risco e garantir consistência:
- a extração do relatório é tratada como etapa obrigatória antes da inativação;
- o relatório da empresa é usado como fonte de verdade para identificar os acumuladores válidos;
- a automação tenta criar uma nova vigência, em vez de alterar diretamente a vigência atual;
- cenários de imposto, vigência, CFOP e Simples Nacional recebem tratamento específico;
- o log da empresa é atualizado ao longo da execução para registrar o estado real do processo.

## Estrutura do projeto

- main.py: ponto de entrada principal.
- agente_rpa/
  - automacao/: interação com o Domínio e ações de mouse/teclado.
  - config/: configuração central, tempos, coordenadas e regras de execução.
  - core/: logger, exceções e utilidades básicas.
  - dominio/: modelos de negócio como Empresa e Acumulador.
  - utils/: leitura de Excel, relatórios e tratamento de dados.
  - workflows/: orquestração do fluxo em lote.

## Pastas e arquivos esperados

O projeto usa uma estrutura relativa à raiz do repositório, mas alguns itens são sensíveis e devem ser criados manualmente pelo usuário antes da execução.

Itens que devem existir na estrutura local:
- planilhas de entrada na raiz do projeto;
- pasta para imagens de reconhecimento visual;
- pasta para relatórios gerados;
- pasta para logs e execução consolidada.

Os nomes esperados são:
- EMPRESA PARA INATIVAR ACUMULADOR.xlsx
- RELAÇÃO DE ACUMULADORES.xlsx
- imagens_rpa/
- Relatorio Final/
- Relatório Inativação/

Esses arquivos e diretórios não são distribuídos com o repositório por segurança. O usuário deve criá-los conforme o ambiente em que o projeto será executado.

## Pré-requisitos

- Python 3.10 ou superior;
- bibliotecas: pyautogui, pandas, openpyxl, pytesseract;
- Tesseract OCR instalado e configurado;
- sistema Domínio acessível e com a interface preparada para automação.

## Instalação

Instale as dependências:

```bash
pip install pyautogui pandas openpyxl pytesseract
```

Verifique se o Tesseract está disponível no ambiente.

## Como usar

1. Crie manualmente os diretórios e arquivos necessários para o ambiente local, como a pasta de imagens, a pasta de relatórios e a pasta de logs.
2. Coloque as planilhas de entrada na raiz do projeto.
3. Organize as imagens usadas pelo agente na pasta imagens_rpa.
4. Ajuste as configurações, se necessário, no arquivo de configuração do projeto.
5. Execute:

```bash
python main.py
```

6. Acompanhe os logs gerados nas pastas de saída.

Importante: não versionar dados sensíveis, planilhas reais, imagens internas ou arquivos com informações privadas. O projeto foi preparado para funcionar com um ambiente local configurado pelo usuário.

## Tutorial de uso correto

Para que o fluxo funcione corretamente:
- mantenha a tela do Domínio ativa durante a execução;
- use uma resolução de tela compatível com o layout esperado;
- preserve a escala da interface em 100%, pois as coordenadas foram calibradas para 1920x1080;
- confira se o sistema está aberto e pronto antes de iniciar;
- não interrompa o processo em meio a uma etapa visual, pois isso pode comprometer o fluxo.

## Observação sobre coordenadas

As coordenadas do projeto foram calibradas para uma tela de 1920x1080 com escala de 100%.
Se a resolução ou a escala do monitor for diferente, pode ser necessário ajustar os valores no arquivo de configuração para que a automação encontre corretamente os elementos da interface.

## Limitações

- a automação depende muito do estado visual da interface do Domínio;
- mudanças de layout, zoom ou resolução podem exigir ajuste fino;
- a execução deve ser supervisionada para validar se os popups e templates estão sendo reconhecidos corretamente.

## Status atual

O projeto já possui uma implementação funcional para execução em lote de inativação de acumuladores, com etapa de extração e validação, fluxo de nova vigência, tratamento de popups e geração de logs por empresa.
