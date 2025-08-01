# Sistema de Avaliação de Estágio Probatório

## Objetivo
Criar um formulário automatizado para que diretores escolares avaliem servidores em estágio probatório, com geração automática de documentos e envio por e-mail ao finalizar.

## Requisitos Técnicos
1. **Plataforma**: Google Workspace (Formulários + Google Apps Script)
2. **Entradas**:
   - Dados do servidor avaliado (nome, matrícula, cargo)
   - Critérios de avaliação (múltiplas seções com escalas numéricas)
   - Campo para observações e recomendações
3. **Saídas**:
   - Documento editável (.docx)
   - PDF de leitura
4. **Disparo**: Envio automático por e-mail ao finalizar o formulário

## Fluxo do Sistema
```mermaid
graph TD
    A[Formulário Google] -->|Submissão| B[Apps Script]
    B --> C{Geração de Documentos}
    C --> D[Template Google Docs]
    D --> E[Preenchimento Automático]
    E --> F[Conversão para PDF]
    F --> G[Envio por E-mail]
    G --> H[Destinatário: Diretor + RH]