# Relatório Gerador

Este é um projeto que gera relatórios a partir de uma planilha em Excel e um template em PowerPoint.

## Inicialização do projeto

Para inicializar o projeto, você precisará criar um ambiente virtual com o comando `python -m venv venv` e ativar este ambiente com o comando `source venv/bin/activate` (ou `./venv/Scripts/Activate.ps1` no Windows).

Em seguida, você precisará instalar as dependências do projeto com o comando `pip install -r requirements.txt`.

## Rodar o programa

Para rodar o programa, você precisará executar o comando `python main.py`.

## Argumentos do programa

O programa utiliza o seguinte argumentos:

* `-p` ou `--path`: Caminho para o arquivo Excel que contém a planilha. O padrão é `dados.xlsx`.
* `-t` ou `--template`: Caminho para o template do PowerPoint. O padrão é `template.pptx`.
* `-o` ou `--output`: Caminho para o arquivo do relatório gerado. O padrão é `relatorio_final.pptx`.

## Template do PowerPoint

O template do PowerPoint deve conter os seguintes placeholders:

* `{{media_ti}}`
* `{{media_rh}}`
* `{{media_juridico}}`
* `{{max_ti}}`
* `{{max_rh}}`
* `{{max_juridico}}`

Estes placeholders serão substituídos pelo script Python com os valores calculados a partir da planilha em Excel.
