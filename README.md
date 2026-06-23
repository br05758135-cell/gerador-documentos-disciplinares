# Gerador de Documentos Disciplinares

Sistema desenvolvido em Python para automatizar a geração de advertências e suspensões disciplinares.

## Funcionalidades

- Leitura automática de planilhas Excel
- Identificação de advertências e suspensões
- Preenchimento automático de modelos DOCX
- Conversão para PDF
- Interface gráfica em Streamlit

## Tecnologias utilizadas

- Python
- Pandas
- OpenPyXL
- Python-docx
- Streamlit
- LibreOffice

## Estrutura do Projeto

backend_gerador.py
→ Responsável pelo processamento dos dados

Gerador_de_arquivos_disciplinares_1.ipynb
→ Notebook principal que cria e executa a aplicação

modelos/
→ Modelos DOCX utilizados na geração

entrada/
→ Planilhas Excel de entrada

saida/
→ PDFs gerados pelo sistema

## Como executar

1. Abrir o notebook no Google Colab
2. Executar a célula de instalação das dependências
3. Montar o Google Drive
4. Configurar os diretórios do projeto
5. Executar o Streamlit
6. Enviar a planilha Excel e gerar os documentos

## Contexto do Projeto

Este sistema foi desenvolvido para uso em ambiente corporativo com fortes restrições de instalação de softwares.

Por esse motivo, todo o processamento foi implementado utilizando Google Colab e Google Drive, eliminando a necessidade de instalação local de Python ou bibliotecas adicionais.

O sistema é utilizado para automatizar a emissão de documentos disciplinares, reduzindo significativamente o tempo gasto em atividades manuais.

## Autor

Leonardo Alves Pereira
