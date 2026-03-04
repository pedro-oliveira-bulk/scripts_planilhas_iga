# Biblioteca de Presenças – Google Apps Script

Sistema modular para gerenciamento de presenças em Google Sheets com arquitetura baseada em biblioteca reutilizável, controle de concorrência, validação de duplicidade, atualização automática de abas, lançar presenças e upload seguro de imagens para Google Drive.

# Visão Geral

A Biblioteca de Presenças foi desenvolvida para centralizar e padronizar a lógica de controle de aulas e presenças em múltiplas planilhas Google Sheets.

Ela resolve problemas comuns como:

- Gravação duplicada de aula
- Execução simultânea
- Inconsistência de datas
- Upload indevido de imagens
- Atualização manual de colunas de presença
- Erro antigo do forms em ter presenças fieis ao lançamento delas
  
O sistema é dividido em:

- **Biblioteca** – Regras de negócio e controle
- **Projeto Cliente** – Interface e personalizações locais

---

# Arquitetura do Sistema

## 1. Biblioteca

Responsável por:

- Validação de regras
- Controle transacional
- Integração com Google Drive
- Atualização automática de abas
- Controle de aula duplicada

### Funções públicas expostas

```javascript
processamentoDiarioCompletoPresencas()
salvarPresencas()
atualizarPagina1ComPresencas()
```

2. Projeto Cliente

Responsável por:

-Interface da planilha
-WebApp de envio de imagens
-Configurações locais
-Chamadas à biblioteca

Exemplo de chamada:

```javascript
function executarProcessoCompleto() {
  PresencasLib.processamentoDiarioCompletoPresencas();
}
```

Funcionalidades:
- Controle de Concorrência

Utiliza:

```javascript
LockService.getScriptLock()
```

Evita:

-Execução simultânea
-Duplicação de registros
-Corrupção de dados
-Validação de Aula Duplicada

Fluxo:

-Normalização da data (remoção do horário)
-Verificação se já existe coluna com mesma data
-Permissão de duplicação apenas se D4 estiver marcado
-Controle de Aula Duplicada

Célula utilizada:

-Lancar_Presenca!D4

TRUE → Aula registrada em:

-Página 1

Correções Manuais

-FALSE → Aula registrada apenas uma vez
-Normalização de Datas

Para evitar erro de comparação:

-new Date(ano, mes, dia)
-Remove horário da data e permite comparação correta.
-Atualização Automática da Página 1

Processo:

-Lê data em B4
-Procura coluna correspondente
-Se não existir, cria nova

Marca:

-P (Presente)
-F (Falta)

-Aplica cores configuradas
-Upload Seguro de Imagens
-Executado apenas após validações
-Cancelado em caso de erro
-Utiliza pasta específica no Drive
-Executa como proprietário do script

Requisitos:

-Conta Google
-Google Sheets
-Google Apps Script
-Permissão para implantar WebApp

### Instalação da Biblioteca:
1. Criar Projeto Base

-Acesse: https://script.google.com
-Crie novo projeto
-Insira o código principal
-Salve

2. Criar Versão

-Clique em Implantar
-Selecione Gerenciar versões
-Clique em Nova versão
-Adicione descrição
-Salve

3. Obter ID do Script

-Vá em Configurações do Projeto
-Copie o ID do Script
-Formato:
  -ex: AKfycbxxxxxxxxxxxxxxxx

### Configuração no Projeto Cliente:
1. Abrir Apps Script da Planilha

No Google Sheets:

-Extensões → Apps Script

2. Adicionar Biblioteca

-Vá em Bibliotecas
-Clique em Adicionar biblioteca
-Cole o ID do Script
-Selecione versão

Defina identificador:

-PresencasLib
-Salvar

Configuração do WebApp:

O projeto cliente deve conter:

```javascript
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("upload");
}
```

Requisitos importantes:
-
-Nome do arquivo HTML deve ser exatamente igual
-Sensível a maiúsculas/minúsculas
-Implantação Recomendada
-Tipo: Aplicativo da Web

Configuração:

-Executar como: Você
-Quem tem acesso: Qualquer pessoa

Isso garante:

-Upload funcional
-Acesso externo permitido
-Controle centralizado

### Estrutura de Pastas:

Biblioteca:

Codigo.gs
WebApp.gs
upload.html

Serviços:
-Drive

Cliente:

-Fluxo Completo do Sistema
-Aba lançar presenças com link do WebApp para carregar imagens e enviar as presenças
-Opção de duplicar presenças

Sistema:

-Aplica Lock
-Valida duplicidade
-Salva dados
-Atualiza Página 1
-Atualiza Correções Manuais (se necessário)
-Realiza upload das imagens
-Finaliza execução

### Boas Práticas Implementadas:

-Controle transacional
-Separação de responsabilidades
-Arquitetura modular
-Reutilização via biblioteca
-Validação antes de efeitos colaterais
-Normalização de dados
-Execução idempotente controlada

### Garantias do Sistema:

-Não grava aula duplicada sem D4
-Não envia imagem em caso de erro
-Não executa simultaneamente
-Cria colunas de presença automaticamente
-Mantém integridade dos dados
