import re
import csv
import sys
import fitz
from tkinter import messagebox
from tkinter.filedialog import askopenfilename



def limpar_texto_gigante(texto_gigante):
    """
    Remove todos os blocos indesejados do texto:
    1. Blocos do fornecedor (de FORNECEDOR até DATA)
    2. Cabeçalhos de tabela com informações de preços
    """

    texto_limpo = re.sub(
        r'FORNECEDOR.*?EQS ENGENHARIA S\.A\..*?DATA\s*\n',
        '',
        texto_gigante,
        flags=re.DOTALL
    )
    
    texto_limpo = re.sub(
        r'QTDE\s+UM\s+PREÇO UNIT\s+'
        r'ICMS/ISS\s+VALOR ST\s+IPI\s+TOTAL DO ITEM\s*'
        r'ITEM\s+DESCRIÇÃO.*?CENTRO\s+ENTREGA.*?\(BRL\).*?\(BRL\).*?\(BRL\).*?\(BRL\)\s*',
        '',
        texto_limpo,
        flags=re.DOTALL
    )

    texto_limpo = re.sub(
        r'^.*\(BRL\).*\n?',
        '',
        texto_limpo,
        flags=re.MULTILINE
    )

    texto_limpo = re.sub(r'\n{3,}', '\n\n', texto_limpo)
    return texto_limpo.strip()


def extrair_texto_como_string_gigante(caminho_pdf):
    """Extrai o PDF até encontrar a frase de corte"""
    texto_final = []
    frase_corte = "Das Cláusulas e Condições: O presente Pedido de Compra obriga as partes que o firmam e seus sucessores a qualquer título às seguintes cláusulas e condições"
    
    with fitz.open(caminho_pdf) as doc:
        for page in doc:
            texto_pagina = page.get_text()
            
            if frase_corte in texto_pagina:
                texto_pagina = texto_pagina.split(frase_corte)[0]
                texto_final.append(texto_pagina)
                break 
                
            texto_final.append(texto_pagina)
    
    return limpar_texto_gigante(''.join(texto_final))


def extrair_dados(texto):
    """
    Extrai dados estruturados do texto no formato do exemplo, incluindo:
    - Número do item (ex: 0001)
    - Código (ex: LC-116/03 07.02)
    - Valor unitário (ex: 344,40)
    - Município (ex: Três Coroas)
    """
    print("="*80)
    dados_extraidos = []
    
    i = 1

    padrao = fr'^\s*{i:04}\s*$'
    match = re.search(padrao, texto, re.MULTILINE)

    padrao_data = r'\b\d{2}\.\d{2}\.\d{4}\b'
    padrao_codigo = "LC-116/03 "


    try:
        while match:
            
            posicao_fim = match.end()
            posicao_codigo = texto[posicao_fim:].find(padrao_codigo)
            ponto_de_partida = posicao_fim + posicao_codigo
            codigo = texto[ponto_de_partida : ponto_de_partida + 15]
            match_data = re.search(padrao_data, texto[ponto_de_partida:], re.MULTILINE)
            match_dinheiro = re.search(r'\d+(?:\.\d{3})*,\d{2}', texto[ponto_de_partida + match_data.end():])
            fim_da_linha = texto[ponto_de_partida + match_data.end() + match_dinheiro.end():].find('\n')
            municipio = texto[ponto_de_partida + match_data.end() + match_dinheiro.end() + fim_da_linha + 1 :].split('\n')[0].strip()
            
            dados_extraidos.append({
                    "Item": f"{i:04}",
                    "Código": codigo,
                    "Municipio": municipio,
                    "Valor Unitario": match_dinheiro.group(),
                })
            
            i += 1
            padrao = fr'^\s*{i:04}\s*$'

            match = re.search(padrao, texto, re.MULTILINE)

    except Exception as e:
        print(f"Erro ao processar o item {i}: {e}")

    return dados_extraidos


def salvar_csv(dados, output_file="dados_extraidos.csv"):
    with open('dados_extraidos.csv', 'w', newline='', encoding='utf-8-sig') as f:  # Note utf-8-sig
        writer = csv.DictWriter(f, fieldnames=dados[0].keys(), delimiter=';')
        writer.writeheader()
        writer.writerows(dados)
    print(f"Dados salvos em {output_file}")



if __name__ == "__main__":

    caminho_pdf = askopenfilename(
        title="Selecione o PDF com o pedido",
        filetypes=[("PDF files", "*.pdf")],
    )

    if not caminho_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF selecionado.")
        sys.exit(1)
    else:
        texto_limpo = extrair_texto_como_string_gigante(caminho_pdf)

        dados_extraidos = extrair_dados(texto_limpo)

        salvar_csv(dados_extraidos)

        messagebox.showinfo("Sucesso!", "Os dados foram extraídos com sucesso!")
