Задание: автоматизация ABC анализа продаж при помощи python
    
import openpyxl

wb_obj = openpyxl.load_workbook("pyton.xlsx")
sheet_obj = wb_obj.active

i=3
o=3
p=3
l=3

n=(sheet_obj.max_row)
print(n)
virfer = 0
virmart = 0
virapr = 0
virmay = 0

while i < n:
    t= sheet_obj.cell(row=i, column = 3)
    print(t.value)
    if t.value!=None:
        virfer=virfer+t.value
        i= i + 1
    else:
        i = i + 1

while o < n:
    m= sheet_obj.cell(row=o, column = 5)
    print(m.value)
    if m.value!=None:
        virmart=virmart+m.value
        o= o + 1
    else:
        o= o + 1
       
while p < n:
    a= sheet_obj.cell(row=p, column = 7)
    print(a.value)
    if a.value!=None:
        virapr=virapr+a.value
        p= p + 1
    else:
        p = p + 1

while l < n:
    ma= sheet_obj.cell(row=l, column = 9)
    print(ma.value)
    if ma.value!=None:
        virmay=virmay+t.value
        l= l + 1
    else:
        l = l + 1

virvse = virfer + virmart + virapr + virmay
      
print('выручка за ферваль' + ' составила ' + str(virfer))
print('выручка за март' + ' составила ' + str(virmart))
print('выручка за апрель' + ' составила ' + str(virapr))
print('выручка за май' + ' составила ' + str(virmay))
print('всего' + 'выручка' + ' составила ' + str(virvse))
