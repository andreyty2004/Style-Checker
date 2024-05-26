# Номер страницы проставляется в середине нижней части листа, без точки.
# Нумерация страниц начинается с титульной страницы, при этом номер на ней не
# проставляется. Нумерация сквозная по всему тексту, включая реферат, иллюстрации на
# отдельных страницах и приложения. 

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION

doc = Document('example.docx')
output = open("OUTPUT.txt", "w+", encoding = "utf-8") 
output.write("Проверка нумерации страниц:\n")

# 3.2 
string1 = str(doc.sections[0].footer._element.xml) 
string2 = str(doc.sections[0].header._element.xml)
cond1 = ('PAGE   \* MERGEFORMAT' in string1) 
cond2 = ('PAGE   \* MERGEFORMAT' in string2)
 
if cond1 and doc.sections[0].different_first_page_header_footer == False: 
    output.write('-- Нумерация страниц НЕ проставляется на титульной странице\n')

if (cond2) == True:
    output.write('-- нумерация страниц должна быть в НИЖНЕМ КОЛОНТИТУЛЕ\n')
elif (cond1) == False: 
    output.write('-- должна быть УСТАНОВЛЕНА НУМЕРАЦИЯ СТРАНИЦ (в нижнем колонтитуле)\n')
elif ('w:jc w:val="center"' in string1) == False: 
    output.write('-- нумерация страниц должна быть УСТАНОВЛЕНА ПО ЦЕНТРУ\n')
else:
    output.write('--OK--')

# check if the first page contains header or a footer
# if (doc.sections[0].header.paragraphs[0].text.isnumeric()):
#     output.write("\n -- Нумерация страниц НЕ проставляется на титульной странице")
#     output.write("\n -- Нумерация страниц проставляется в НИЖНЕЙ ЧАСТИ листа")
#     if(doc.sections[0].header.paragraphs[0].alignment != WD_ALIGN_PARAGRAPH.CENTER):
#         output.write("\n -- Нумерация страниц проставляется с ВЫРАВНИВАНИЕМ ПО ЦЕНТРУ")
# elif (doc.sections[0].footer.paragraphs[0].text.isnumeric()):
#     output.write("\n -- Нумерация страниц НЕ проставляется на титульной странице")
#     if(doc.sections[0].footer.paragraphs[0].alignment != WD_ALIGN_PARAGRAPH.CENTER):
#         output.write("\n -- Нумерация страниц проставляется с ВЫРАВНИВАНИЕМ ПО ЦЕНТРУ")
# else:
#     output.write("\n -- OK --")

# # check the style of page numbering of the first page
# print(doc.paragraphs[0].runs[0].text)

# print(doc.sections[0].header.paragraphs[0].text)
# # доступ к верхнему колонтитулу
# header = doc.sections[0].header.paragraphs[0]

# # доступ к нижнему колонтитулу
# footer = doc.sections[0].footer.paragraphs[0]

# # добавляем верхний колонтитул
# header.style.font.size = Pt(8)

# header.add_run('Верхний колонтитул')

# # выравниваем колонтитул по правому краю
# header.alignment = WD_ALIGN_PARAGRAPH.CENTER

# # добавляем нижний колонтитул
# footer.style.font.size = Pt(10)
# footer.add_run('Нижний колонтитул')

# # выравниваем колонтитул по правому краю
# footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

# # Добавим параграф
# doc.add_paragraph('Текст первой страницы.')
# doc.add_paragraph('Доступ к верхнему и нижнему колонтитулам.')

# # Добавим разрыв страницы
# doc.add_page_break()
# doc.add_paragraph('Текст второй страницы.')
# doc.add_paragraph('Доступ к верхнему и нижнему колонтитулам.')

# # теперь прочитаем колонтитулы
# text_header = doc.sections[0].header.paragraphs[0].text
# print(text_header)
# text_footer = doc.sections[0].footer.paragraphs[0].text
# print(text_footer)
doc.save("example.docx")

