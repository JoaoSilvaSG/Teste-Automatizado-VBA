# Testes automatizados para ERP

Este repositório contém funções VBA para automatizar testes em um ERP específico, com o objetivo de verificar comportamentos como rastreamento de registros e abertura de cabeçalhos. As funções podem ser utilizadas como base para desenvolver diferentes cenários de teste automatizados.

## Estrutura do projeto

- **/core:** Contém as funções VBA reutilizáveis, como abertura de menus, preenchimento de campos, cliques em botões e coleta de rastros.
- **/exemplos:** Contém exemplos de uso das funções do módulo `/core`, prontos para execução.
- **Planilha de exemplo:** Uma planilha Excel que demonstra o uso das macros para automatizar testes reais no ERP.

## Como usar

1. Abra o arquivo `.xlsm` no Excel.
2. Ative as macros, se solicitado.
3. Execute uma macro de exemplo ou crie suas próprias macros reutilizando as funções em `/core`.

## Pré-requisitos

- Microsoft Excel com suporte a macros (habilitar VBA).
- Acesso ao ERP com permissões adequadas para testes.
- Configuração do ERP para aceitar interações automatizadas (se aplicável).

## Funcionalidades principais

- Abertura de menus e navegação por atalhos.
- Interação com formulários do ERP (preenchimento, clique em botões).
- Leitura e validação de informações da interface.
- Captura de rastros para validação pós-execução.
- Verificação de imagens
