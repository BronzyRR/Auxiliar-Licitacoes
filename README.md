# Automatizador de Documentos para Licitação

## 📝 Descrição

Este projeto é um conjunto de scripts em Python criados para automatizar a geração de documentos e planilhas essenciais para a participação em processos licitatórios de engenharia. A solução é dividida em duas frentes principais:

1.  **Geração de Planilhas para OrçaFascio**: Converte e formata planilhas de cotações, composições e orçamentos para o padrão de importação do software OrçaFascio.
2.  **Criação de Documentos Word**: Gera automaticamente as declarações e anexos necessários para a habilitação e proposta em licitações, utilizando dados pré-configurados.

A ferramenta foi projetada para minimizar o trabalho manual, aumentar a produtividade e reduzir erros na preparação da documentação.

## ✨ Funcionalidades Principais

-   **Importação para OrçaFascio**:
    -   Criação da planilha de **Insumos** a partir de uma lista de cotações.
    -   Criação da planilha de **Composições** formatada corretamente, tratando composições principais e auxiliares.
    -   Geração da planilha do **Orçamento** com a estrutura de itens hierárquica (Metas, Níveis).
-   **Geração de Documentos de Habilitação e Proposta**:
    -   Criação automática de mais de 15 documentos em Word, incluindo:
        -   Declaração de Aceitação do Responsável Técnico
        -   Declaração de Indicação da Equipe Técnica
        -   Declarações Diversas (Fato Impeditivo, Não Emprego de Menor, etc.)
        -   Carta Proposta
        -   Declaração de Elaboração Independente de Proposta
        -   Capas para CD e Envelopes.
-   **Padronização**: Garante que todos os documentos sigam um padrão de formatação consistente (estilos, fontes, cabeçalho e rodapé).

## 🛠️ Tecnologias Utilizadas

-   **Python 3**
-   **openpyxl**: Para ler e escrever arquivos Excel (`.xlsx`).
-   **xlrd**: Para ler arquivos do formato antigo do Excel (`.xls`).
-   **xlsxwriter**: Para criar e formatar arquivos Excel (`.xlsx`) com mais opções de estilo.
-   **python-docx**: Para criar e manipular documentos do Microsoft Word (`.docx`).
