import math2docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm
from docx.shared import Inches
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO

doc = Document("new.docx")

# Page №3

text = doc.add_paragraph("Для заданной схемы дано:")
text.runs[0].font.size = Pt(14)
text.paragraph_format.left_indent = Mm(15)
text = doc.add_paragraph("Таблица 1")
text.alignment = WD_ALIGN_PARAGRAPH.RIGHT

start_table1 = {'0': [100, 1, -120], '1': [300, 5.5, 90], '2': [280, 5, 60], '3': [260, 4.5, 45],
                '4': [240, 4, 30], '5': [220, 3.5, 0], '6': [200, 3, -30], '7': [175, 2.5, -45],
                '8': [150, 2, -60], '9': [125, 1.5, -90]}
start_table2 = {'0': [1000, 10, 0.02, 200], '1': [100, 100, 2, 200], '2': [150, 90, 1.2, 150],
                '3': [200, 80, 0.8, 125], '4': [250, 75, 0.6, 107], '5': [300, 60, 0.4, 111],
                '6': [400, 50, 0.25, 100], '7': [500, 40, 0.16, 100],
                '8': [600, 30, 0.1, 111], '9': [800, 24, 0.06, 200]}
# Получаем ввод от пользователя
input_row1 = int(input("Введите номер строки для таблицы 1: "))
input_row2 = int(input("Введите номер строки для таблицы 2: "))
# Вводим начальные условия
E = start_table1[str(input_row1)][0]
J = start_table1[str(input_row1)][1]
alpha = start_table1[str(input_row1)][2]
omega = start_table2[str(input_row2)][0]
R = start_table2[str(input_row2)][1]
L = start_table2[str(input_row2)][2]
C = start_table2[str(input_row2)][3]

table1_dictionary = {0: ["E", "J", "α", "ω", "R", "L", "C"],
                     1: ["B", "A", "град", "1/с", "Ом", "Гн", "мкФ"],
                     2: [str(E), str(J), str(alpha), str(omega), str(R), str(L), str(C)]}
table1 = doc.add_table(rows=3, cols=7, style='Table Grid')
for i in range(len(table1_dictionary)):
    for j in range(len(table1_dictionary[0])):
        table1.cell(i, j).text = str(table1_dictionary[i][j])
doc.add_picture('C:/Users/опасныйперчик/Desktop/MEGA/ТОЭ2/РГР4/Python/TOE4_1scheme/image/1.PNG')
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph("Рисунок 1.  Заданная схема электрической цепи").alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph()


# Функция создания жирного текста по центру
def bold_text(text):
    paragraph = doc.add_paragraph(str(text))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.runs[0]
    run.font.bold = True

bold_text("Цепь первого порядка")
doc.add_paragraph("При постоянном источнике тока e(t) = E после срабатывания ключа К1, "
                  "когда ключ К2 ещё не сработал, определяем напряжение i(t).").paragraph_format.first_line_indent = Mm(12.5)
doc.add_paragraph("1.1. Используем упрощённый классический метод, "
                  "когда дифференциальное уравнение для искомой функции  не составляется.").paragraph_format.first_line_indent = Mm(12.5)
text = doc.add_paragraph("1.1.1. Определяем независимые начальные условия (ННУ) при t=0: ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,"i_{L}(0-).")
run = text.add_run("Схема до коммутации: установившийся режим, постоянный источник, С – разрыв, L – закоротка.")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# Функция добавления картинки
def add_image(number, text):
    doc.add_picture(r'C:/Users/опасныйперчик/Desktop/MEGA/ТОЭ2/РГР4/Python/TOE4_1scheme/image/{}.PNG'.format(number))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(str(text)).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

add_image(2,"Рисунок 2.  ННУ")
math2docx.add_math(doc.add_paragraph(),  r"i_{L}(0-)=0\ A")

# Page №4

math2docx.add_math(doc.add_paragraph(),  r"i(0-)=\frac{E}{R}="+str(round((E/R),2))+r"\ A")
doc.add_paragraph("1.1.2. Определяем ЗНУ при t=0+. (Cхема после коммутации ключа К1). ").paragraph_format.first_line_indent = Mm(12.5)
add_image(3,"Рисунок 3.  ЗНУ")
doc.add_paragraph("По закону Ома:")
math2docx.add_math(doc.add_paragraph(),  r"i(0+)=\frac{E}{2R}=\frac{"+str(E)+"}{"+str(2*R)+"}="+str(round((E/(2*R)),2))+r"\ A")
text = doc.add_paragraph("1.1.3. Определяем принуждённую составляющую ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,"i_{пр}(t) ")
text.add_run("при t=∞  (схема после коммутации ключа К1, установившейся режим, постоянный источник, С – разрыв, L – закоротка)")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

add_image(3,"Рисунок 3.  ПС")
doc.add_paragraph("По закону Ома:")
math2docx.add_math(doc.add_paragraph(),  r"i_{пр}=\frac{E}{R+\frac{R}{2}}=\frac{"+str(E)+"}{"+str(R)+r"+\frac{"+str(R)+"}{2}}="+str(round((E/(R+R/2)),2))+r"\ A")
text = doc.add_paragraph("1.1.4. Определяем корень характеристического уравнения p. Используем метод сопротивления цепи после коммутации (C → 1/Cp; L → Lp),  "
                  "причем ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,r"R_{J}=∞,\ R_{E}=0.")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# Page №5

add_image(4,"Рисунок 4. Схема определения корня характеристического уравнения")
math2docx.add_math(doc.add_paragraph(),  r"Z(p)=Lp+\frac{3R}{2}=0")
math2docx.add_math(doc.add_paragraph(),  r"p=-\frac{3R}{2L}=\frac{"+str(3*R)+"}{"+str(2*L)+"}="+str(round((-(3*R)/(2*L)),2))+r"\ 1/c")
doc.add_paragraph("1.1.5. Определяем постоянную интегрирования B:")
math2docx.add_math(doc.add_paragraph(),  r"B=i(0+)-i_{пр}="+str(round((E/(2*R)),2))+"-"+str(round((E/(R+R/2)),2))+"="+str(round(((E/(2*R))-(E/(R+R/2))),2))+r"\ B")
doc.add_paragraph("1.1.6. Окончательный результат:")
math2docx.add_math(doc.add_paragraph(),  r"i(t)=i_{пр}+Be^{pt}="+str(round((E/(R+R/2)),2))+"+("+str(round(((E/(2*R))-(E/(R+R/2))),2))+r")e^{"+str(round((-(3*R)/(2*L)),2))+r"t}\ ,B")
math2docx.add_math(doc.add_paragraph(),  r"\tau =\frac{1}{|p|}=\frac{1}{|"+str(round((-(3*R)/(2*L)),2))+"|}="+str(round(1/(abs(-(3*R)/(2*L))),4))+r"\ c - постоянная\ времени.")

# Начало и конец изменения значения X
t = np.linspace(0,
                    0.1,
                    100)

#Функция создания и добавления графика
def create_plot(name,ylabel,text,limit_up_y,function_t):
    memfile = BytesIO()
    plt.title(str(name))
    plt.xlabel('t,c')
    plt.ylabel(str(ylabel))
    plt.xlim(0, 0.1)
    plt.ylim(0, int(limit_up_y))
    y = function_t
    plt.plot(t, y)
    plt.grid()
    plt.savefig(memfile)
    doc.add_picture(memfile, width=Inches(5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    memfile.close()
    doc.add_paragraph(str(text)).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

create_plot("Переходный процесс",
            "i, A",
"Рисунок 5. Классический метод",
            2,
            (E/(R+R/2))+((E/(2*R))-(E/(R+R/2)))*np.exp((-(3*R*t)/(2*L))))


doc.save("new.docx")

