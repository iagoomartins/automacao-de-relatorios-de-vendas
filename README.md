Este script Python automatiza o processo de geração de um relatório de vendas a partir de um arquivo Excel e o envio por email. O script utiliza as bibliotecas pandas, openpyxl e win32com.client para manipulação de dados e envio de emails.

Funcionalidades:
- Carrega dados de vendas de um arquivo Excel.
- Calcula o faturamento total por loja.
- Computa a quantidade total de produtos vendidos por loja.
- Determina o ticket médio por loja.
- Gera um email em HTML com o relatório de vendas e o envia para um destinatário especificado.

Como Funciona?
- O script lê os dados de vendas do arquivo Vendas.xlsx em um DataFrame do pandas.
- Agrega os dados de faturamento e quantidade por loja.
- Calcula o valor médio do ticket dividindo o faturamento total pela quantidade de produtos vendidos.
- Cria o corpo do email em HTML com o relatório de vendas.
- Envia o email utilizando o aplicativo Outlook.


Requisitos:
Python
pandas
openpyxl
pywin32
