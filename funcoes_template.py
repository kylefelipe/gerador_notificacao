""" Criado para trabalhar partes do template"""
from docx import Document
from docx import table


def popula_tabela (tabela, dados, documento):
    """ Popula a tabela do template com os itens da empresa presentes em dados_empresa_tabela"""

    tabela_d = tabela[0]

    for iten_d in dados:
        adiciona = tabela_d.add_row().cells
        adiciona[0].text = str(iten_d[0])
        adiciona[1].text = str(iten_d[1])
        adiciona[2].text = str(iten_d[2])
        adiciona[3].text = str(iten_d[3])
        adiciona[4].text = str(iten_d[4])
        adiciona[5].text = str(iten_d[5])

    # Mudando a formatação das células da tabela
    for linha_d in tabela_d.rows:
        for celular_d in linha_d.cells:
            celular_d.paragraphs[0].style = documento.styles['Normal2']  # Pode criar um outor estilo dentro do documento
                                                                         # do Template para ser aplicado às celulas da tabela
    return tabela_d
