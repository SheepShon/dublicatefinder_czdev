import xlsxwriter
import os
import json
from colorama import Fore, Back, Style
from colorama import init

version = "0.3"

os.system("cls")
print(Style.BRIGHT + Fore.BLUE + "czDevelopment "+ Fore.WHITE + "| czDublicateFinder |" + Fore.YELLOW + " Version " + str(version) + Fore.WHITE + "\n")

file = open("что-искать.json" , "r")
m1 = json.load(file) # массив ориг артикулов "что искать"
file.close()

file = open("где-искать.json" , "r")
m2 = json.load(file) # массив ориг артикулов "где искать"
file.close()

m3 = [] # массив ориг артикулов дубликатов
m4 = {} # массив ориг артикулов не найденных
m5 = {} # массив норм артикулов "что искать"
m6 = {}	# массив норм артикулов "где искать"
last_art = ""

def normalize(a):
	b = a.replace(" ", "")
	c = b.replace("-", "")
	d = c.replace("/", "")
	e = d.replace(".", "")
	return e

for a in m1:
	e = normalize(a)
	m5[e] = a

for a in m2:
	e = normalize(a)
	m6[e] = a

m5_keys = list(m5.keys())
m6_keys = list(m6.keys())

for a in m5:
	if not a in m6_keys:
		m3.append(m5[a])
	for b in m2:
		if a == normalize(b):
			if last_art == a:
				if m5[a] in m4:
					m4[m5[a]] = m4[m5[a]] + 1
				else:
					m4[m5[a]] = 2
			last_art = a
		else:
			pass

workbook = xlsxwriter.Workbook('out.xlsx') # Создать файл
worksheet = workbook.add_worksheet('out')

bold_and_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
center = workbook.add_format({'align': 'center'})
worksheet.write('A1', 'czDevelopment | czDublicateFinder | Version ' + str(version) + "")
worksheet.write('A2', 'Не найдено ('+str(len(list(m3)))+" шт.)", bold_and_center)
worksheet.write('B2', 'Дубликаты ('+str(len(list(m4)))+" шт.)", bold_and_center)
worksheet.write('C2', 'шт.', bold_and_center)
worksheet.set_row(0, 20)
worksheet.set_column(0, 0, 20)
worksheet.set_column(1, 1, 20)
worksheet.set_column(2, 2, 5)

i=3
for a in m3:
	worksheet.write('A'+str(i), a, center)
	i = i+1
i=3
for a in m4:
	worksheet.write('B'+str(i), a, center)
	worksheet.write('C'+str(i), m4[a], center)
	i = i+1
workbook.close()

print(Fore.GREEN + "Обработка завершена. Данные сохранены в файле out.xlsx\n")
print(Fore.WHITE+"Нажмите любую клавишу для закрытия консоли..")
print(Style.NORMAL + Fore.BLACK+"")