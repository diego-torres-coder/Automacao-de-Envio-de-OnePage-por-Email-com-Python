# Automação de Envio de OnePages por E-mail com Python

## Entendendo o Desafio

Você trabalha em uma rede com 25 lojas espalhadas pelo Brasil. Diariamente, os Analistas de Dados calculam o **OnePage** de cada loja e o enviam para o gerente da respectiva loja. O **OnePage** é um resumo gerencial que permite analisar o desempenho de uma loja com base nos seguintes indicadores:

- faturamento do dia
- faturamento do ano
- diversidade de produtos do dia
- diversidade de produtos do ano
- ticket médio do dia
- ticket médio do ano

Estes analistas também enviam para a diretoria um e-mail contendo dois anexos: um com um ranking de faturamento do dia e outro com um ranking de faturamento do ano.

Além disso, este processo envolve a criação de arquivos de backup numa pasta intitulada **Backup Arquivos Lojas**, a qual contém 25 subpastas — cada uma com o nome de uma das 25 lojas. Cada arquivo contém os registros de vendas de uma loja no dia da análise e é nomeada da seguinte forma:

```text
dia-mes-Nome da Loja.xlsx
```

## Passo a Passo para Solução do Desafio

Resumidamente, estes foram os 3 passos adotados para criar uma automação para resolver este desafio:


1. Calcular o OnePage de cada loja

2. Enviar um e-mail para cada gerente com o OnePage de sua respectiva loja, anexando a planilha com os dados de venda da loja 

3. Enviar um e-mail para a diretoria com um ranking de faturamento das lojas, informando qual loja apresentou melhor desempenho e qual apresentou pior desempenho

## Bibliotecas Utilizadas

Estas foram as bibliotecas utilizadas neste projeto:

- os
- yagmail
- python-dotenv
- pandas

## Como Reproduzir este Projeto

Inicialmente, navegue para a pasta na qual deseja clonar o repositório deste projeto. Em seguida, clone este repositório com o seguinte comando:

```bash
git clone https://github.com/diego-torres-coder/Automacao-de-Envio-de-OnePage-por-Email-com-Python
```

Para criar o ambiente virtual, navegue até a pasta do projeto  e digite o seguinte comando no terminal:

```bash
conda -n rpa-onepages python=3.10
```

Em seguida, ative o ambiente:

```bash
conda activate rpa-onepages
```

Com o ambiente ativo, instale as dependências do projeto:

```bash
pip3 install pandas openpyxl numpy jupyter yagmail python-dotenv
```

Alternativamente, você pode instalar as dependências deste projeto a partir do arquivo `requirements.txt`:

```bash
pip3 install -r requirements.txt
```

Execute o Jupyter Notebook:

```bash
jupyter notebook
```

Abra o arquivo principal deste projeto e execute todas as células.

