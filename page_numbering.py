from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION

# doc = Document('example.docx')
# output = open("OUTPUT.txt", "w+", encoding = "utf-8") 
# output.write("Проверка нумерации страниц:\n")

# 3.2 
def page_numbering(fpath = ""):
    doc = Document(fpath)
    output = "-- Проверка нумерации страниц --\n"
    string1 = str(doc.sections[0].footer._element.xml) 
    string2 = str(doc.sections[0].header._element.xml)
    cond1 = ('PAGE   \* MERGEFORMAT' in string1) 
    cond2 = ('PAGE   \* MERGEFORMAT' in string2)
    
    if (cond1 or cond2) and doc.sections[0].different_first_page_header_footer == False: 
        output = output + "-> Нумерация страниц НЕ проставляется на титульной странице\n"
    if (cond2) == True:
        output = output + "-> нумерация страниц должна быть в НИЖНЕМ КОЛОНТИТУЛЕ\n'"
    elif ('w:jc w:val="center"' in string1) == False: 
        output = output + "-> нумерация страниц должна быть УСТАНОВЛЕНА ПО ЦЕНТРУ\n"
    else:
        output = output + "-> OK\n"
    
    output = output + '\n'
    return output

