from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH #чтобы смотреть выравнивания
import re
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import shutil
import os
from lxml import etree
# import pyunpack

doc = Document('file.docx')

output = open("OUTPUT.txt", "w+")

string = str(doc.sections[0].footer._element.xml)
cond = ('PAGE   \* MERGEFORMAT' in string)

if cond == False:
    output.write('должна быть установлена нумерация страниц (в нижней части страницы)')
elif ('w:jc w:val="center"' in string) == False:
    output.write('нумерация страниц должна быть установлена по центру')

if cond and doc.sections[0].different_first_page_header_footer == False:
    output.write('номер страницы на титульном листе не ставится')