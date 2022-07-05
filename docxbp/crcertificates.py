from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.shared import RGBColor
from openpyxl import load_workbook


excel_file = (r"C:\Users\PC\Documents\Applications\scrapybots\docxbp\certificates.xlsx")#Configura o caminho do arquivo .xslx com a base de dados
open_excel_file = load_workbook(excel_file)#Abre o arquivo para escrita
open_sheet = open_excel_file["Dados"]#Seleciona a guia Dados


for i in range(2, len(open_sheet["A"]) + 1):#Loop de linhas

    word_file = Document(r"C:\Users\PC\Documents\Applications\scrapybots\docxbp\certificate.docx")#Abre o arquivo .docx com o modelo
    fiLe_style = word_file.styles["Normal"]#Configura o tamanho


    original_paragraph = "Frase frase frase frase frase @Coluna5 frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase @Coluna2 de @Coluna3 de @Coluna4."#Paragráfo original com valores a serem substituídos

    coluna1_value = open_sheet['A%s' % i].value#Subtítulo

    paragraph_one = "Frase frase frase frase frase "
    coluna5_value = open_sheet['E%s' % i].value#Valor para substituir @Coluna5
    
    paragraph_two = " frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase frase "

    coluna2_value = open_sheet['B%s' % i].value#Valor para substituir @Coluna2
    paragraph_three = " de "

    coluna3_value = open_sheet['C%s' % i].value#Valor para substituir @Coluna3
    paragraph_four = " de "

    coluna4_value = open_sheet['D%s' % i].value#Valor para substituir @Coluna4
    paragraph_five = "."

    coluna6_value = open_sheet['F%s' % i].value#Valor para substituir @Coluna6
    #O %s converte o valor de cada linha para string


    new_paragraph = f"{paragraph_one} {coluna5_value} {paragraph_two} {coluna2_value} {paragraph_three} {coluna3_value} {paragraph_four} {coluna4_value} {paragraph_five} {coluna6_value}"


    for pr in word_file.paragraphs:
        if "@Coluna1" in pr.text:
            pr.text = coluna1_value#Procura por @Coluna1 em cada parágrafo, quando encontra indexa o novo valor da linha

            fiLe_font = fiLe_style.font#Parâmetro de fonte
            fiLe_font.name = "Calibri (Corpo)"#Muda o tipo da fonte
            fiLe_font.size = Pt(24)#Muda o tamanho da fonte

        if "@Coluna2" in pr.text:
            pr.text = new_paragraph#Procura por @Coluna4 em cada parágrafo, quando encontra indexa o novo parágrafo concatenado

        if "@Coluna3" in pr.text:
            pr.text = new_paragraph

        if "@Coluna4" in pr.text:
            pr.text = new_paragraph
        
        if "@Coluna5" in pr.text:
            pr.text = new_paragraph

        if "@Coluna6" in pr.text:
            pr.text = new_paragraph


    new_file_path = "C:\\Users\\PC\\Documents\\Applications\\scrapybots\\docxbp\\nf\\" + coluna1_value + ".docx"

    word_file.save(new_file_path)#Salva como um novo arquivo
    print("Arquivos criados com sucesso!")