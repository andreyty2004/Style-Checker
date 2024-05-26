from docx import Document
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from bs4 import BeautifulSoup


doc = Document('test2.2.docx')


for i in range(0, len(doc.paragraphs)):
    count1 = doc.paragraphs[i].text.count('фиг.')
    count2 = doc.paragraphs[i].text.count('Фиг.')
    if count1 + count2 > 0:
        print('В русскоязычной литературе не принято использовать обозначения «фигура 1», «Фиг.1».')
        break


def find_title(i, word):
    strs = [run.text.lower() for run in doc.paragraphs[i].runs]

    c = 1
    while c <= len(strs):
        for i in range(0, len(strs) - c + 1):
            string = ''
            for j in range(0, c):
                string += strs[i + j]
            if word in string:
                a = []
                for k in range(0, c):
                    a.append(i + k)        
                return a
        c += 1
    
    return []

def check_numbering_par(title_runs, number, c, name, is_appendix, page, par):

    title_text = ''.join([par.runs[i].text for i in range(title_runs[0], len(par.runs))])
    number_of_title = number + '.' + str(c)

    if number_of_title not in title_text:
        if is_appendix:
            print(f'Ошибка в подписи к рисунку "{name}" на странице {page}. Рисунки в приложениях нумеруются по схеме «номер приложения – точка – номер рисунка».')
        else:
            print(f'Ошибка в подписи к рисунку "{name}" на странице {page}. Рисунки в разделах нумеруются по схеме «номер раздела – точка – номер рисунка».')





def check_numbering_titles(stroka, number, c, name, is_appendix, page):
    soup = BeautifulSoup(stroka, "lxml")
    text = soup.find_all("w:t")
    title_text = ''.join([element.text for element in text])

    number_of_title = number + '.' + str(c)

    if number_of_title in title_text and title_text.find(number_of_title) == len(title_text) - len(number_of_title):
        pass
    else:
        if is_appendix:
            print(f'Ошибка в подписи к рисунку "{name}" на странице {page}. Рисунки в приложениях нумеруются по схеме «номер приложения – точка – номер рисунка».')
        else:
            print(f'Ошибка в подписи к рисунку "{name}" на странице {page}. Рисунки в разделах нумеруются по схеме «номер раздела – точка – номер рисунка».')


def page_increment(i, j):
    global isActive
    global page

    if j != len(doc.paragraphs[i].runs) - 1:
        if "w:br" in doc.paragraphs[i].runs[j]._r.xml and "w:lastRenderedPageBreak" in doc.paragraphs[i].runs[j + 1]._r.xml:
            page += 1
            isActive = False
            return 
        
    if i != len(doc.paragraphs) - 1 and len(doc.paragraphs[i + 1].runs) > 0 and j == len(doc.paragraphs[i].runs) - 1:
        if "w:br" in doc.paragraphs[i].runs[j]._r.xml and "w:lastRenderedPageBreak" in doc.paragraphs[i + 1].runs[0]._r.xml:
            page += 1
            isActive = False
            return

    if "w:br" in doc.paragraphs[i].runs[j]._r.xml:
        page += 1
        return

    if "w:lastRenderedPageBreak" in doc.paragraphs[i].runs[j]._r.xml and isActive:
        page += 1
        return

    if "w:lastRenderedPageBreak" in doc.paragraphs[i].runs[j]._r.xml:
        isActive = True  
        return



def get_picture_bottom_boarder_pt(run_xml):
    positionV = re.findall('(<wp:positionV(.|\n)*?wp:positionV>)', str(run_xml))[0]
    posOffset = re.findall('(<wp:posOffset.*?wp:posOffset>)', str(positionV))[0]
    top_boarder = re.findall('([-+]?\d+)', str(posOffset))[0]

    extent = re.findall('(<wp:extent(.|\n)*?>)', str(run_xml))[0]
    cy = re.findall('(cy=".*?")', str(extent))[0]
    height = re.findall('(\d+)', str(cy))[0]

    return (float(top_boarder) + float(height)) / 12700

def get_picture_name(run):
    string = str(run._r.xml)
    substr = re.findall('<wp:docPr.*?>', string)[0]
    substr1 = re.findall('descr=".*"', substr)[0]
    name = (re.findall('".*"', substr1)[0])[1:-1]
    return name

def get_title_top_margin(stroka):
    shape = re.findall('(<v:shape[^a-zA-Z].*?>)', stroka)[0]
    margin_top = re.findall('(margin-top:[-+]?\d+\.?\d+)', str(shape))[0]
    return float(re.findall('(\d+\.?\d+)', margin_top)[0])


def checking_str(stroka, name, bottom_boarder, page, in_one_par):

    if stroka.count('"center"') == 0:
        print(f'Подпись к рисунку "{name}" на странице {page} должна размещаться по центру')
    
    if stroka.count('<w:sz w:val="28"/>') == 0:
        print(f'Размер шрифра подписи к рисунку "{name}" на странице {page} должен совпадать с размером шрифта в основном тексте')

    if 'w:b w:val="0"' not in stroka:
        print(f'Шрифт подписи к рисунку "{name}" на странице {page} не должен быть написан жирным шрифтом')
    
    runs = BeautifulSoup(stroka, "lxml").find_all("w:r") # прогоны в подписи w:pict
    for tag in runs:
        if len(tag.find_all("w:i")) != 0:
            print(f'Шрифт подписи к рисунку "{name}" на странице {page} не должен быть написан курсивом')
            break
        

    for tag in runs:
        if len(tag.find_all("w:u")) != 0:
            print(f'Шрифт подписи к рисунку "{name}" на странице {page} не должен быть подчеркнут')
            break


    if 'w:hAnsi="Times New Roman"' not in stroka:
        print(f'Тип шрифра подписи к рисунку "{name}" на странице {page} должен совпадать с типом шрифта в основном тексте')

    if in_one_par:
        if abs(get_title_top_margin(stroka) - bottom_boarder) > 2.8:
            print(f'Подпись должна размещаться сразу под рисунком "{name}" на странице {page}')

    if stroka.count('stroked="f"') == 0:
        print(f'Подпись к рисунку {name} на странице {page} не должна выделяться рамкой')





def markup_options(run, name, page): #r - run     page - number of page

    sec = doc.sections[0]
    left_margin = sec.left_margin.pt
    right_margin = sec.right_margin.pt
    page_width = sec.page_width.pt

    if len(re.findall('(<wp:positionH(.|\n)*?wp:positionH>)', str(run._r.xml))) == 0:
        left_boarder = 0
    else:
        positionH = re.findall('(<wp:positionH(.|\n)*?wp:positionH>)', str(run._r.xml))[0]
        posOffset = re.findall('(<wp:posOffset.*?wp:posOffset>)', str(positionH))[0]
        left_boarder = re.findall('([-+]?\d+)', str(posOffset))[0]

    extent = re.findall('(<wp:extent(.|\n)*?>)', str(run._r.xml))[0]
    cx = re.findall('(cx=".*?")', str(extent))[0]
    picture_width = re.findall('(\d+)', str(cx))[0]

    cy = re.findall('(cy=".*?")', str(extent))[0]
    picture_heigth = re.findall('(\d+)', str(cy))[0]

    picture_area = (float(picture_width) / 12700) * (float(picture_heigth) / 12700)
    
    # 2.8 pt ~ 1 mm    достаточно близко
    if (('wp:wrapTopAndBottom' in str(run._r.xml)) or 
        ('wrapText="bothSides"' in str(run._r.xml) and abs(float(left_boarder) / 12700) < 2.8) and
        (abs((float(picture_width) / 12700 + float(left_boarder) / 12700 + left_margin) - (page_width - right_margin)) < 2.8)) == False:
        print(f'При подготовке иллюстрации "{name}" на странице {page} в редакторе MSWord следует использовать опции меню «Параметры разметки – Обтекание текстом – Сверху и снизу» или «Параметры разметки – Обтекание текстом – Вокруг рамки» при условии, что ширина рамки совпадает с шириной текста.')        
    
    return picture_area

def empty_string_separation(i, j, page): #i - index of drawing, j - index of title
    drawing_par = doc.paragraphs[i]
    if 'wp:wrapTopAndBottom' in str(drawing_par._p.xml) or 'wrapText="bothSides"' in str(drawing_par._p.xml):
        if len(doc.paragraphs[min(i, j) - 1].runs) != 0 and len(doc.paragraphs[max(i, j) + 1].runs) != 0:
            print(f'При размещении иллюстрации "{name}" на странице {page} в тексте следует отделять рисунок от текста пустой строкой и сверху, и снизу')


def process_picture_area(i, j, picture_area, page): #i - drawing      j - title
    sec = doc.sections[0]
    page_width = sec.page_width.pt
    page_heigth = sec.page_height.pt
    page_area = page_width * page_heigth

    if (picture_area < 0.5 * page_area) and doc.paragraphs[i].runs[0].contains_page_break:
        for k in range(1, 4):
            if len(doc.paragraphs[max(i, j) + k].runs) > 0:
                if (doc.paragraphs[max(i, j) + k].runs[0].contains_page_break):
                    print(f'Рисунок "{name}" на странице {page} слишком маленький, чтобы размещать его на отдельной странице')
                    return False

                if len(doc.paragraphs[max(i, j) + k].runs) > 1:
                    if ('w:br' in doc.paragraphs[max(i, j) + k].runs[0]._r.xml and doc.paragraphs[max(i, j) + k].runs[1].contains_page_break):
                        print(f'Рисунок "{name}" на странице {page} слишком маленький, чтобы размещать его на отдельной странице')
                        return False

                if len(doc.paragraphs[max(i, j) + k].runs) == 1 and len(doc.paragraphs[max(i, j) + k + 1].runs) > 0:
                    if ('w:br' in doc.paragraphs[max(i, j) + k].runs[0]._r.xml and doc.paragraphs[max(i, j) + k + 1].runs[0].contains_page_break):
                        print(f'Рисунок "{name}" на странице {page} слишком маленький, чтобы размещать его на отдельной странице')
                        return False
                return True

    if doc.paragraphs[i].runs[0].contains_page_break:
        return False

    return True

        
def checking_par(par, name, title_runs, page):

    if par.alignment != WD_ALIGN_PARAGRAPH.CENTER:
        print(f'Подпись к рисунку "{name}" на странице {page} должна размещаться по центру')

    size = par.style.font.size
    for j in title_runs:
        if par.runs[j].font.size != None:
            size = par.runs[j].font.size
        if size != None and size.pt != 14.0:
            print(f'Размер шрифра подписи к рисунку "{name}" на странице {page} должен совпадать с размером шрифта в основном тексте')
            break
    if size == None:
        print(f'Размер шрифра подписи к рисунку "{name}" на странице {page} должен совпадать с размером шрифта в основном тексте')

    F = par.style.font.bold #в случае если унаследовано от заголовка
    for j in title_runs:
        if par.runs[j].bold != None:
            F = par.runs[j].bold
        if (F == True):
            print(f'Шрифт подписи к рисунку "{name}" на странице {page} не должен быть написан жирным шрифтом')
            break

    for j in title_runs:
        if par.runs[j].font.italic == True:
            print(f'Шрифт подписи к рисунку "{name}" на странице {page} не должен быть написан курсивом')
            break


    for j in title_runs:
        if par.runs[j].font.name != 'Times New Roman':
            print(f'Тип шрифра подписи к рисунку "{name}" на странице {page} должен совпадать с типом шрифта в основном тексте')
            break

    for j in title_runs:
        if par.runs[j].font.underline == True:
            print(f'Шрифт подписи к рисунку "{name}" на странице {page} не должен быть подчеркнут')
            break

    
def check_numbering_above(i, name, number, c, page):
    par = doc.paragraphs[i]

    mention_runs = find_title(i, 'рисун')

    mention_text = ''.join([par.runs[i].text for i in range(mention_runs[0], len(par.runs))])
    number_of_title = number + '.' + str(c)

    if number_of_title not in mention_text:
        print(f'Ошибка в упоминании рисунка "{name}" на странице {page} в тексте перед рисунком. Рисунки в разделах нумеруются по схеме «номер раздела – точка – номер рисунка».')




        


def mentioned_above(i, name, number, c, page): # i - index of drawing

    par = doc.paragraphs[i]

    if 'Рис.' in par.text or 'рис.' in par.text:
        print(f'Ссылка на рисунок "{name}" на странице {page} должна быть дана без сокращений')
        return

    if 'Рисун' in par.text or 'рисун' in par.text:
        check_numbering_above(i, name, number, c, page)

        if 'w:instrText' in par._p.xml:
            return
        else:
            print(f'Ссылка на рисунок "{name}" на странице {page} в тексте должна являться перекрестной ссылкой (при нажатии на нее переносит на рисунок)')
            return 
    
    
    for k in range(1, 4):
        par_k = doc.paragraphs[i - k]
        if len(par_k.runs) != 0:
            if (len(par_k.runs) == 1 and 'w:br' in par_k.runs[0]._r.xml):
                continue


            c1 = par_k.text.count('Рис.')
            c2 = par_k.text.count('рис.')
            if c1 > 0 or c2 > 0:
                print(f'Ссылка на рисунок "{name}" на странице {page} должна быть дана без сокращений')
                break

            c1 = par_k.text.count('Рисун')
            c2 = par_k.text.count('рисун')
            if c1 + c2 == 0:
                print(f'Не найдено упоминание рисунка "{name}" на странице {page} в тексте перед рисунком')
                break
            
            if c1 + c2 != 0:
                check_numbering_above(i - k, name, number, c, page)
                if 'w:instrText' not in par_k._p.xml:
                    print(f'Ссылка на рисунок "{name}" на странице {page} в тексте должна являться перекрестной ссылкой (при нажатии на нее переносит на рисунок)')

            break




is_appendix = False #приложение
number_defined = False #нумерация заголовка определена

page = 1
isActive = True # можно ли увеличивать page


for i in range(0, len(doc.paragraphs)):

    par = doc.paragraphs[i]

    if 'Heading 1' in par.style.name and len(par.runs) > 0:
        number_defined = False
        head_str = par.text

        if 'приложение' in head_str.lower():
            is_appendix = True
            start = head_str.lower().find('приложение') + len('приложение')
            stroka = head_str[start:] #строка без 'приложение'
            numbers = re.findall('([А-Я])', stroka)
            if len(numbers) > 0:
                number = numbers[0]
                number_defined = True
                c = 0
        else:
            numbers = re.findall('(\d+)', head_str)
            if len(numbers) > 0:
                number = numbers[0]

                if head_str.find(number) == 0:
                    number_defined = True
                    c = 0



    for j in range(0, len(par.runs)):
        run = par.runs[j]

        page_increment(i, j)


        if 'w:drawing' in str(run._r.xml): 
            if number_defined:
                c += 1 #какой по счету рисунок в этом разделе     

            name = get_picture_name(run)
            picture_area = markup_options(run, name, page)

            if is_appendix == False:
                mentioned_above(i, name, number, c, page)

        
            if 'a:noFill' not in str(run._r.xml):
                print(f'Рисунок "{name}" на странице {page} не должен выделяться рамкой')

            search = re.findall('(w:pict(.|\n)*?w:pict)', str(par._p.xml))
            length = len(search)

            #нумерован надписью
            if length > 0:
                stroka = str(search[-1])
                for k in reversed(range(0, length)):
                    if 'w:instr' in str(search[k]):
                        stroka = str(search[k])
                        break

                bottom_boarder = get_picture_bottom_boarder_pt(run._r.xml)
                checking_str(stroka, name, bottom_boarder, page, True)
                if process_picture_area(i, i, picture_area, page):
                    empty_string_separation(i, i, page)
                if number_defined:
                    #          номер_раздела.номер_рисунка
                    check_numbering_titles(stroka, number, c, name, is_appendix, page)
                continue

            search = re.findall('(w:pict(.|\n)*?w:pict)', str(doc.paragraphs[i - 1]._p.xml))
            length = len(search)

            if length > 0:
                stroka = str(search[-1])
                for k in reversed(range(0, length)):
                    if 'w:instr' in str(search[k]):
                        stroka = str(search[k])
                        break

                bottom_boarder = get_picture_bottom_boarder_pt(run._r.xml)
                checking_str(stroka, name, bottom_boarder, page, False)
                if process_picture_area(i, i - 1, picture_area, page):
                    empty_string_separation(i, i - 1, page)
                if number_defined:
                    check_numbering_titles(stroka, number, c, name, is_appendix, page)
                continue


            search = re.findall('(w:pict(.|\n)*?w:pict)', str(doc.paragraphs[i + 1]._p.xml))
            length = len(search)
            if length > 0:
                stroka = str(search[-1])
                for k in reversed(range(0, length)):
                    if 'w:instr' in str(search[k]):
                        stroka = str(search[k])
                        break

                bottom_boarder = get_picture_bottom_boarder_pt(run._r.xml)
                checking_str(stroka, name, bottom_boarder, page, False)
                if process_picture_area(i, i + 1, picture_area, page):
                    empty_string_separation(i, i + 1, page)
                if number_defined:
                    check_numbering_titles(stroka, number, c, name, is_appendix, page)
                continue

            if doc.paragraphs[i + 1].text != '':
                if len(title_runs := find_title(i + 1, 'рисунок')) != 0:
                    checking_par(doc.paragraphs[i + 1], name, title_runs, page)
                    if process_picture_area(i, i + 1, picture_area, page):
                        empty_string_separation(i, i + 1, page)
                    if number_defined:
                        check_numbering_par(title_runs, number, c, name, is_appendix, page, doc.paragraphs[i + 1])
                    continue


            if doc.paragraphs[i + 2].text != '':
                if len(title_runs := find_title(i + 2, 'рисунок')) != 0:
                    checking_par(doc.paragraphs[i + 2], name, title_runs, page)
                    if process_picture_area(i, i + 2, picture_area, page):
                        empty_string_separation(i, i + 2, page)
                    if number_defined:
                        check_numbering_par(title_runs, number, c, name, is_appendix, page, doc.paragraphs[i + 2])
                    continue

            if doc.paragraphs[i].text != '':
                if len(title_runs := find_title(i, 'рисунок')) != 0:
                    checking_par(par, name, title_runs, page)
                    if process_picture_area(i, i, picture_area, page):
                        empty_string_separation(i, i, page)
                    if number_defined:
                        check_numbering_par(title_runs, number, c, name, is_appendix, page, doc.paragraphs[i])
                    continue

            print(f'Подпись к иллюстрации {name} на странице {page} не найдена')



        



#w:instr - если в подписи стоит номер или SEQ
#w:instrText - если в тексте нумерован правильно (стоит ссылка) иначе этой штуки нет

#w:bookmarkStart - в подписи если на него есть ссылка его имя можно поискать в поза-позапрошлом параграфе тк прошлый - картинка
#а позапрошлый - пустой 
#w:bookmarkStart - можно сверить на что ссылается ссылка из текста с помощью name REF
#w:instrText REF - есть если в тексте стоит именно ссылка

