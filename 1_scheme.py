import math2docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document("new.docx")

# Page №3
text = doc.add_paragraph("Для заданной схемы дано:")
text.runs[0].font.size = Pt(14)
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


doc.save("new.docx")