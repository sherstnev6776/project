Задание: автоматизация ABC и XYZ анализа продаж при помощи python
#найти и прочитать эксель файл
import pandas as pd # pip install pandas

df = pd.read_excel(r'C:\Users\Алекс\Desktop\МАС\Кофе с собой\Продажи.xlsx', sheet_name='Лист1')

#присвоить класс (переменную) столбцам
class Tovar(Object):
    art = None
    count = None
    price = None

    part_of_all = None

    def __init__(self, art, count, price):
        self.art = art
        self.count = count
        self.price = price

def load_from_xml(xml_path):
    with open(xml_path, 'rb'):
        pass

    res = [[], [], []]
    res = [{'артикул': 123, 'кол-во': 23, 'объём': 45}]
    res = [Tovar(123, 23, 45)]



    res[0].part_of_all = a/b


    return res

def find_all_xmls('path'):

    return list_of_xmls_path
#найти долю продаж товара в объем объеме
def analyse(list_of_res):

#отсортировать строки по убыванию

#суммировать каждую строку с долей продаж к предыдущей строке в столбце

#присовить каждой строчке название (A <= 80%,80% < B <=95%, 95%< С <=100%)
    return report

#провести анализ XYZ на выяление стабильности спроса
