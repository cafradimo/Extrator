# Extrator

Este projeto é uma ferramenta desenvolvida para o CREA-RJ que permite a extração e processamento de dados de arquivos PDF e planilhas Excel. Ele utiliza a biblioteca **Streamlit** para criar uma interface interativa e amigável para o usuário.

## Funcionalidades

### 1. Extrator PDF Consolidado
- Extrai automaticamente dados de arquivos PDF relacionados a:
  - **Vínculos e S.I**: Identifica vínculos e informações de S.I nos documentos.
  - **Ramos de Atividade**: Processa e consolida informações sobre ramos de atividade.
  - **Processos/Protocolos**: Extrai dados de processos e protocolos, incluindo fiscais e retornos.

- Gera relatórios em:
  - **PDF**: Relatórios individuais para cada tipo de dado.
  - **Excel Consolidado**: Um arquivo Excel com todas as informações extraídas.

### 2. Processador de Planilhas de Autuações
- Processa planilhas Excel para extração de dados consolidados.
- Exibe os dados carregados e permite o processamento adicional.

## Tecnologias Utilizadas

- **Python**: Linguagem principal do projeto.
- **Streamlit**: Para a interface web interativa.
- **pdfplumber**: Para extração de texto de arquivos PDF.
- **FPDF**: Para geração de relatórios em PDF.
- **Pandas**: Para manipulação e análise de dados.
- **PyPDF2**: Para manipulação de arquivos PDF.
- **Pillow**: Para manipulação de imagens.
- **OpenPyXL**: Para geração de arquivos Excel.

## Estrutura do Projeto
