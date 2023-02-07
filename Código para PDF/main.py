#biblioteca para manipular pdf
import PyPDF2

def extrair_pdf(nome_arquivo):
    #abre o pdf em formato binário e armazena em arquivo
    arquivo = open(nome_arquivo, 'rb')
    #le o conteudo do arquivo
    dados_do_arquivo = PyPDF2.PdfFileReader(arquivo)
    pagina = dados_do_arquivo.getPage(0)
    #Extrai texto do pdf
    texto_do_pdf = pagina.extractText()
    return texto_do_pdf

nome_arquivo = r'D:\UFS\Projeto Fapitec\Planejamentos Docentes\Planejamentos 2020\Escola Rosinha\01.01 - 1º Ano - Anos Iniciais - A - Mat - 07.09.2020 - 12.09.2020.pdf'
pdf = extrair_pdf(nome_arquivo)
print(pdf)