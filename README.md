# ShootMail — Sistema de Gestão e Rastreamento de Emails

O **ShootMail** é uma solução robusta integrada ao ecossistema Google (Sheets + Gmail) para automação, rastreamento e gestão de processos de comunicação via email. Idealizado para o contexto da Marinha do Brasil, o sistema foca em eficiência no contato com fornecedores e monitoramento em tempo real de disparos.

## 🚀 Principais Funcionalidades

- **Dashboard Inteligente**: Acompanhamento de métricas de envios, aberturas e respostas com gráficos de atividade.
- **Gestão de Fornecedores**: Cadastro detalhado incluindo múltiplos emails, tipos de material e códigos de item.
- **Modelos de Email (Templates)**: Criação de mensagens dinâmicas com uso de variáveis como `{nome}` e `{empresa}`.
- **Rastreamento de Leads**: Monitoramento de leitura via pixel de rastreio e registro automático de data/hora de abertura.
- **Sincronização de Respostas**: Vinculação inteligente de respostas do Gmail ao processo correspondente através de headers e IDs únicos.
- **Relatórios**: Geração de relatórios completos em formatos PDF, CSV e TXT para auditoria e controle.
- **Automação de Gmail**: Verificação automática de novas respostas a cada 15 minutos (auto-sync).

## ⚙️ Guia de Configuração

Para que o sistema funcione corretamente, siga os passos abaixo:

### 1. Planilha Google
- Certifique-se de que a planilha possui as abas necessárias: `Fornecedores`, `Processos`, `Destinatarios_Processo`, `Disparos`, `Respostas` e `Config`.
- O sistema possui um script de `setupPlanilha` que cria essas abas automaticamente na primeira execução.

### 2. Google Apps Script (Backend)
- No menu da Planilha, vá em **Extensões > Apps Script**.
- Cole o código do arquivo `google_apps_script.js`.
- Clique em **Implantar > Nova implantação**.
- Selecione o tipo de implantação como **App da Web**.
- **Executar como**: "Eu" (seu email).
- **Quem pode acessar**: "Qualquer pessoa" (necessário para que o pixel de rastreio e os hooks funcionem).
- Copie a **URL do App da Web** gerada.

### 3. Configuração no Frontend
- Abra o arquivo `remixed-9bded00e.html` no navegador.
- Vá na aba **Configurações (⚙️)**.
- Preencha:
    - **URL do Web App**: A URL copiada no passo anterior.
    - **ID da Planilha**: O código longo presente na URL da sua planilha.
    - **Email/Nome do Remetente**: Suas credenciais do Gmail.
- Clique em **Salvar** e teste a conexão.

## 🎖️ Créditos

Este sistema é um produto de inovação e dedicação técnica.

**Idealizado e desenvolvido por**: COpAb - Sobressalente  
**Versão**: 1.2026  
**Créditos Especiais**: 2ºSG Freitas 11.0316

---
© 2026 ShootMail — Eficiência em Comunicação Digital.
