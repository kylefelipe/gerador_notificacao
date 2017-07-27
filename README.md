Gerador de Notificação

Feito em Python 3.6, utilizando os módulos openpyxl e python-docx.
Por Kyle Felipe Vieira Roberto
Em Julho de 2017 para atender a demanda de emissão de diversas notificações que possuiam mais de um item por empresa.

Utilizando um arquivo DOCX e um XLSX como base de dados, permite gerar uma série de documentos, tendo o DOCX como base.
No DOCX tem um texto padrão, com algumas partes que podem ser modificadas, número da notificação, data, nome da empresa,
número do CPF/CNPJ, e uma tabela, onde entrará os itens que devem constar da empresa.

O script altera determinados textos em alguns parágrafos do DOCX padrão:

    No parágrafo 0 altera o texto '{0}' para o número da notificação.
    No parágrafo 3 altera o texto '{0}' para a data da emissão.
    No parágrafo 5 altera o texto '{0}' para o nome da empresa.
    No parágrafo 6 altera o texto '{0}' para o número do CPF ou CNPJ.

Há uma tabela inserida dentro do arquivo que irá receber os itens constantes na notificação. Os nomes das colunas
podem ser substituidos pelos nomes que o usuário precisar.
Pode haver mais de um item por empresa(CPF ou CNPJ), cada um desses itens será inserido dentro da notificação
da empresa, não é aconselável fazer mesclas no arquivo XLSX, os dados devem ser repetidos em todas as celulas e nem deve
conter fórmulas.

Foi criado, dentro do arquivo DOCX um estilo chamado "NORMAL2" que é aplicado a cada uma das células da tabela após
ela ser populada com os dados de cada item, pode-se modificar esse estilo no próprio DOCX ou criar um novo e mudar o
o nome dentro da função "popula_tabela" em "funcoes_template".

O nome do arquivo é salvo da seguinte maneira: NOTIFICAÇÃO Nº <NÚMERO DA NOTIFICAÇÃO> - <NOME DA EMPRESA>, nesse momento
é são removidos os segintes caracteres especiais "[-,./\()]"  evitando erro na geração do arquivo.
