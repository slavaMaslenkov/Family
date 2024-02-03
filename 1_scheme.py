import math2docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm
from docx.shared import Inches
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO
import sympy as sp
import re
import cmath

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
t = np.linspace(-0.01,
                    0.1,
                    100)

#Функция создания и добавления графика
def create_plot(name,ylabel,text,limit_up_y,function_t,limit_down_y = 0):
    plt.clf()
    memfile = BytesIO()
    plt.title(str(name))
    plt.xlabel('t,c')
    plt.ylabel(str(ylabel))
    plt.xlim(0, 0.1)
    plt.ylim(int(limit_down_y), int(limit_up_y))
    plt.grid(True)
    y = function_t
    plt.plot(t, y)
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

doc.add_paragraph("1.2. Используем операторный метод.").paragraph_format.first_line_indent = Mm(12.5)
doc.add_paragraph("1.2.1. Находим независимые начальные условия (п. 1.1.1)").paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(doc.add_paragraph(),  r"i_{L}(0-)=0\ A")
doc.add_paragraph("1.2.2. В операторной схеме после коммутации используем метод контурных токов:").paragraph_format.first_line_indent = Mm(12.5)
add_image(5,"Рисунок 6. Операторная схема")
math2docx.add_math(doc.add_paragraph(),r"\begin{displaymath}"
r"\left\{ \begin{array}{ll}"
r"I_{11}(p)\cdot 2R+I_{22}(p)\cdot R=\frac{E}{p}  \\"
r"I_{11}(p)\cdot R+I_{22}(p)\cdot (2R+Lp)=-L\cdot i_{L}(0-)=0 "
r"\end{array} \right."
r"\end{displaymath}")
doc.add_paragraph("Определяем операторное изображение искомого тока:").paragraph_format.first_line_indent = Mm(12.5)

#Решение уравнения
x, y,p = sp.symbols('x y p')
eq1 = sp.Eq(2*R*x + R*y, E/p)
eq2 = sp.Eq(R*x + y*(2*R+L*p), 0)
solution = sp.solve((eq1, eq2), (x, y))

math2docx.add_math(doc.add_paragraph(),r"i(p)=I_{11}(p)="+str(solution[x])+"")
doc.add_paragraph("По теореме разложения находим i(t):").paragraph_format.first_line_indent = Mm(12.5)

#Выделяем знаменатель для нахождения его нулей
denominator = re.findall(r'\((.*?)\)', str(solution[x]))[1]
match = re.search(r'(\d*\.?\d*)\*p\*\*2', denominator)
if match:
    number_in_square = match.group(1)
else:
    number_in_square = str(1)
match = re.findall(r'(\d*\.?\d*)\*p$', denominator)[0]
if match:
    number = match
else:
    number = str(1)

#Решатель уравнений
def solver(equation, value):
    p = value
    return eval(equation.replace("p", str(value)))

math2docx.add_math(doc.add_paragraph(),str(number_in_square)+"p^{2}+"+str(number)+"p=0")
numer, denom = sp.fraction(solution[x])
zeros_from_denom = sp.solve(denom)
derivative_of_denom = str(sp.Derivative(denom,p).doit())
math2docx.add_math(doc.add_paragraph(),"p_{0}="+str(round((zeros_from_denom[0]),2))+",p_{1}="+str(round((zeros_from_denom[1]),2))+r"\quad ,1/c")
math2docx.add_math(doc.add_paragraph(),r"i(t)=\frac{"+str(solver((re.findall(r'\((.*?)\)', str(solution[x]))[0]),round((zeros_from_denom[1]),2)))+"}{"+str(solver(derivative_of_denom,round((zeros_from_denom[1]),2)))+"}e^{"+str(round((zeros_from_denom[1]),1))+r"}+"
                                            r"\frac{"+str(solver((re.findall(r'\((.*?)\)', str(solution[x]))[0]),round((zeros_from_denom[0]),0)))+"}{"+str(solver(derivative_of_denom,round((zeros_from_denom[0]),0)))+"}e^{"+str(round((zeros_from_denom[0]),0))+r"\cdot t}="+str(round((E/(R+R/2)),2))+"+("+str(round(((E/(2*R))-(E/(R+R/2))),2))+r")e^{"+str(round((-(3*R)/(2*L)),2))+r"t}\ ,B")
create_plot("Переходный процесс",
            "i, A",
"Рисунок 7. Операторный метод",
            2,
            (E/(R+R/2))+((E/(2*R))-(E/(R+R/2)))*np.exp((-(3*R*t)/(2*L))))

text = doc.add_paragraph("2. При гармоническом источнике ЭДС ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,r"e(t)=\sqrt{2}E\sin (\omega t+\alpha)")
text.add_run(" определить ток i(t).")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
text = doc.add_paragraph("2.1. Используем упрощённый классический метод, когда дифференциальное уравнение для искомой функции i(t)  не составляется.").paragraph_format.first_line_indent = Mm(12.5)
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
text = doc.add_paragraph("2.1.1. ННУ. Определяем независимые начальные условия при  t=0-; ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,r"i_{L}(0-)")
text.add_run(" (схема до коммутации установившийся режим, гармонический источник, символический метод). ")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

add_image(6,"Рисунок 8. ННУ")

#Переопределение Е
E = E*(cmath.cos(np.radians(alpha))+1j*cmath.sin(np.radians(alpha)))

complex_number = E/R
doc.add_paragraph("Ток течет через закорокту. ").paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(doc.add_paragraph(),r"\underline{E}="+str(abs(E))+"e^{j"+str(alpha)+r"} ,B")
math2docx.add_math(doc.add_paragraph(),r"X_{L}=\omega L="+str(L*omega)+r" ,Ом")
math2docx.add_math(doc.add_paragraph(),r"\underline{I}=\frac{\underline{E}}{R}="+str(round(complex_number.real, 2) + round(complex_number.imag, 2) * 1j)+r"="+str(round((abs(E/R)),2))+"e^{"+str(round((np.degrees(cmath.phase(E/R))),1))+"j} ,A")
math2docx.add_math(doc.add_paragraph(),r"i(t)="+str(round((abs(E/R)),2))+r"\sqrt{2} \sin ("+str(omega)+"t+("+str(round((np.degrees(cmath.phase(E/R))),1))+")) ,A;")
math2docx.add_math(doc.add_paragraph(),r"i(0)="+str(round((abs(E/R)),2))+r"\sqrt{2} \sin ("+str(round((np.degrees(cmath.phase(E/R))),1))+")="+str(round((((abs(E/R))*np.sqrt(2)*np.sin(cmath.phase(E/R)))),2))+" ,A;i_{L}(0)=0.")
doc.add_paragraph("2.1.2. Определяем ЗНУ при t=0+  (схема после коммутации ключа К1):  ").paragraph_format.first_line_indent = Mm(12.5)
add_image(7,"Рисунок 9. ЗНУ")
doc.add_paragraph("По Закону Ома:  ").paragraph_format.first_line_indent = Mm(12.5)

math2docx.add_math(doc.add_paragraph(),r"e(t)="+str(round((abs(E)),2))+r"\sqrt{2} \sin ("+str(omega)+"t+("+str(alpha)+")) ,B;")
math2docx.add_math(doc.add_paragraph(),r"e(0)="+str(round((abs(E)),2))+r"\sqrt{2} \sin ("+str(alpha)+")="+str(round((((abs(E))*np.sqrt(2)*np.sin(np.radians(alpha)))),2))+" ,B;")
math2docx.add_math(doc.add_paragraph(),r"i(0+)=\frac{E(0)}{2R}="+str(round((((((abs(E))*np.sqrt(2)*np.sin(np.radians(alpha)))))/(2*R)),2))+",A")
text = doc.add_paragraph("2.1.3. Определяем принуждённую составляющую  ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,r"i_{пр}(0-)")
text.add_run(" при t=∞ (схема после коммутации ключа К1: установившейся режим, гармонический источник, символический метод):  ")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
add_image(8,"Рисунок 9. ПС")
doc.add_paragraph("По Закону Ома:  ").paragraph_format.first_line_indent = Mm(12.5)
complex_number = R+(R*(R+1j*(omega*L)))/(2*R+1j*(omega*L))
math2docx.add_math(doc.add_paragraph(),r"\underline{Z}=R+\frac{R\cdot (R+jX_{L})}{R+R+jX_{L}}="+str(round(complex_number.real, 2) + round(complex_number.imag, 2) * 1j)+r"="+str(round((abs(complex_number)),2))+"e^{"+str(round((np.degrees(cmath.phase(complex_number))),1))+"j} ,Ом")
complex_number = E/(R+(R*(R+1j*(omega*L)))/(2*R+1j*(omega*L)))
math2docx.add_math(doc.add_paragraph(),r"\underline{I_{пр}}=\frac{\underline{E}}{\underline{Z}}="+str(round(complex_number.real, 2) + round(complex_number.imag, 2) * 1j)+r"="+str(round((abs(complex_number)),2))+"e^{"+str(round((np.degrees(cmath.phase(complex_number))),1))+"j} ,A")
math2docx.add_math(doc.add_paragraph(),r"i_{пр}(t)="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(omega)+"t+("+str(round((np.degrees(cmath.phase(complex_number))),1))+")) ,A;")
math2docx.add_math(doc.add_paragraph(),r"i_{пр}(0)="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(round((cmath.phase(complex_number)),1))+")="+str(round((((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number)))),2))+" ,A;")
doc.add_paragraph("2.1.4. Определяем корень характеристического уравнения : Используем метод сопротивления цепи после коммутации. Аналогично п. 1.1.4 получаем p="+str(round((-(3*R)/(2*L)),2))+" ,1/c").paragraph_format.first_line_indent = Mm(12.5)
doc.add_paragraph("2.1.5. Определяем постоянную интегрирования B:").paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(doc.add_paragraph(),r"B=i(0+)-i_{пр}(0)="+str(round(((((((abs(E))*np.sqrt(2)*np.sin(np.radians(alpha)))))/(2*R))-(((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number))))),2))+",A;")
doc.add_paragraph("2.1.6. Окончательный результат").paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(doc.add_paragraph(),  r"i(t)=i_{пр}+Be^{pt}="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(omega)+"t+("+str(round((np.degrees(cmath.phase(complex_number))),1))+"))+("
                   +str(round(((((((abs(E))*np.sqrt(2)*np.sin(np.radians(alpha)))))/(2*R))-(((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number))))),2))+")e^{"+str(round((-(3*R)/(2*L)),2))+"t} ,A")
text = doc.add_paragraph("причем  ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,r"\tau = \frac{1}{|p|}="+str((round((1/(zeros_from_denom[0])),4)))+r",\frac{1}{c};t_{п}=5\tau="+str(round((5*(1/(-(3*R)/(2*L)))),2))+r",c;"
                                                                                                    r"T=\frac{2\tau}{\omega}="+str(round((((2*(-(3*R)/(2*L)))/(omega))),4))+"c,")
text.add_run(" -период принужденной составляющей.  ")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
create_plot("Переходный процесс",
            "i, A",
"Рисунок 9. Классичекий метод с гармоническим источником",
            2,
            (((abs(complex_number))*np.sqrt(2)*np.sin(omega*t+(cmath.phase(complex_number)))+((((abs(E))*np.sqrt(2)*np.sin(np.radians(alpha)))/(2*R))-((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number))))*np.exp((-(3*R*t)/(2*L))))),
                   -2)
doc.add_paragraph("2.2. Используем комбинированный операторно-классический метод для определения i(t). ").paragraph_format.first_line_indent = Mm(12.5)
doc.add_paragraph("2.2.1. Находим независимые начальные условия (п. 2.1.1): ").paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(doc.add_paragraph(),r"i_{L}(0)=0")
text = doc.add_paragraph("2.2.2. Определяем принуждённые составляющие  ")
text.paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(text,r"i_{пр}(t),i_{Lпр}(t)")
text.add_run(" при t=∞ (схема после коммутации ключа К1: установившийся режим, гармонический источник, символический метод.) ")
text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
add_image(10,"Рисунок 10. ПС")
doc.add_paragraph("По закону Ома:").paragraph_format.first_line_indent = Mm(12.5)
complex_number = R+(R*(R+1j*(omega*L)))/(2*R+1j*(omega*L))
math2docx.add_math(doc.add_paragraph(),r"\underline{Z}=R+\frac{R\cdot (R+jX_{L})}{R+R+jX_{L}}="+str(round(complex_number.real, 2) + round(complex_number.imag, 2) * 1j)+r"="+str(round((abs(complex_number)),2))+"e^{"+str(round((np.degrees(cmath.phase(complex_number))),1))+"j} ,Ом")
complex_number = E/(R+(R*(R+1j*(omega*L)))/(2*R+1j*(omega*L)))
math2docx.add_math(doc.add_paragraph(),r"\underline{I_{пр}}=\frac{\underline{E}}{\underline{Z}}="+str(round(complex_number.real, 2) + round(complex_number.imag, 2) * 1j)+r"="+str(round((abs(complex_number)),2))+"e^{"+str(round((np.degrees(cmath.phase(complex_number))),1))+"j} ,A")
math2docx.add_math(doc.add_paragraph(),r"i_{пр}(t)="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(omega)+"t+("+str(round((np.degrees(cmath.phase(complex_number))),1))+")) ,A;")
math2docx.add_math(doc.add_paragraph(),r"i_{пр}(0)="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(round((cmath.phase(complex_number)),1))+")="+str(round((((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number)))),2))+" ,A;")
doc.add_paragraph("По правилу разброса:").paragraph_format.first_line_indent = Mm(12.5)
complex_number = (E/(R+(R*(R+1j*(omega*L)))/(2*R+1j*(omega*L))))*(R/(2*R+1j*(omega*L)))
math2docx.add_math(doc.add_paragraph(),r"\underline{I_{Lпр}}=I_{пр}\cdot \frac{R}{2R+jX_{L}}="+str(round(complex_number.real, 2) + round(complex_number.imag, 2) * 1j)+r"="+str(round((abs(complex_number)),2))+"e^{"+str(round((np.degrees(cmath.phase(complex_number))),1))+"j} ,A")
math2docx.add_math(doc.add_paragraph(),r"i_{Lпр}(t)="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(omega)+"t+("+str(round((np.degrees(cmath.phase(complex_number))),1))+")) ,A;")
math2docx.add_math(doc.add_paragraph(),r"i_{Lпр}(0)="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(round((cmath.phase(complex_number)),1))+")="+str(round((((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number)))),2))+" ,A;")
doc.add_paragraph("2.2.3. Определяем начальное значение свободной составляющей тока на индуктивности").paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(doc.add_paragraph(),r"i_{Lсв}(0)=i_{L}(0)-i_{Lпр}(0)="+str(round((-((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number)))),3))+" ,A;")
doc.add_paragraph("2.2.4. Рассчитываем операторную схему замещения для свободных составляющих.").paragraph_format.first_line_indent = Mm(12.5)
add_image(11,"Рисунок 11. Операторная схема")
doc.add_paragraph("По закону Ома и по теореме разложения находим i(t):").paragraph_format.first_line_indent = Mm(12.5)
math2docx.add_math(doc.add_paragraph(),r"I(p)=\frac{L\cdot i_{Lсв}(0)}{pL+1.5R} \cdot \frac{R}{2R};")
p = sp.symbols('p')
math2docx.add_math(doc.add_paragraph(),r"I(p)="+str(sp.simplify((0.5*L*(round((-((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number)))),3)))/(1.5*R+L*p)))+r"=\frac{D(p)}{B(p)};")
complex_number = E/(R+(R*(R+1j*(omega*L)))/(2*R+1j*(omega*L)))
math2docx.add_math(doc.add_paragraph(),r"I(p)=i_{пр}(t)+i_{св}(t)="+str(round((abs(complex_number)),2))+r"\sqrt{2} \sin ("+str(omega)+"t+("+str(round((np.degrees(cmath.phase(complex_number))),1))+r"))+\frac{D_{k}(p_{k})}{B'_{k}(p_{k})}e^{p_{k}t}=\\"
                   +str(round((abs(complex_number)), 2)) + r"\sqrt{2} \sin (" + str(omega) + "t+(" + str(round((np.degrees(cmath.phase(complex_number))), 1)) + "))+("
                   + str(round(((((((abs(E)) * np.sqrt(2) * np.sin(np.radians(alpha))))) / (2 * R)) - (((abs(complex_number)) * np.sqrt(2) * np.sin(cmath.phase(complex_number))))), 2)) + ")e^{" + str(round((-(3 * R) / (2 * L)), 2)) +r"}t,A")

create_plot("Переходный процесс",
            "i, A",
"Рисунок 12. Комбинированный метод",
            2,
            (((abs(complex_number))*np.sqrt(2)*np.sin(omega*t+(cmath.phase(complex_number)))+((((abs(E))*np.sqrt(2)*np.sin(np.radians(alpha)))/(2*R))-((abs(complex_number))*np.sqrt(2)*np.sin(cmath.phase(complex_number))))*np.exp((-(3*R*t)/(2*L))))),
                   -2)
сделать перенос формулы!!!!!!!


doc.save("new.docx")

