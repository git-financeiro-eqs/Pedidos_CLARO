# Gerador de planilhas pedido
<br/>
Este é um script simples em Python que processa um arquivo PDF, extrai dados pertinentes ao faturamento, e gera uma planilha no formato .CSV.
<br/>
<br/>

### Funcionalidades

Lê o arquivo PDF com as informações do pedido.  
Elimina partes do texto que não são pertinentes.  
Extrai informações pertinentes ao faturamento, como:    
Número do item,
Código do item,
Valor,
Municipio.
Gera um arquivo CSV com esses dados, pronto para ser utilizado.
<br/>
<br/>

### Pré-requisitos:
Python 3.x;  
Biblioteca: PyMuPDF;  
<br/>

### Execução:
Execute o script Python no seu terminal ou ambiente de desenvolvimento, ou através do executável.  
O script irá pedir para selecionar o arquivo PDF e processará os dados automaticamente.  
O script irá gerar um arquivo CSV com os resultados na mesma pasta onde o script foi executado.
<br/>
<br/>

### Dependências
Este projeto depende de uma biblioteca externa Python. Você pode instalá-las usando pip:
```bash
pip install PyMuPDF  
```
As demais bibliotecas já fazem parte da instalação padrão do Python.
