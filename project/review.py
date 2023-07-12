"""
Для подготовки текста к обзору и к описанию графиков
"""


import pandas as pd
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
from docx.text.run import Font, Run
import os


def value_change(name, current, previous):
    if name == "Цена":
        if current - previous > 0:
            return "увеличилась"
        elif current - previous == 0:
            return "не изменилась"
        else:
            return "уменьшилась"
    elif name == "Объем":
        if current - previous > 0:
            return "увеличился"
        elif current - previous == 0:
            return "не изменился"
        else:
            return "уменьшился"
    elif name == "Предложение":
        if current - previous > 0:
            return "увеличилось"
        elif current - previous == 0:
            return "не изменилось"
        else:
            return "уменьшилось"
    elif name == "Рост/падение":
        if current - previous > 0:
            return "наблюдался рост"
        elif current - previous == 0:
            return "не изменилось"
        else:
            return "наблюдалось падение"
    elif name == "Больше/меньше":
        if current - previous > 0:
            return "больше"
        elif current - previous == 0:
            return "не изменилось"
        else:
            return "меньше"
    else:
        return ""
    
def decl(number: int, titles = None):
    cases = [ 2, 0, 1, 1, 1, 2 ]
    if 4 < number % 100 < 20:
        idx = 2
    elif number % 10 < 5:
        idx = cases[number % 10]
    else:
        idx = cases[5]
    if titles == None:
        return f"""{'{:,}'.format(number).replace(',',' ')}"""
    else:
        return f"""{'{:,}'.format(number).replace(',',' ')} {titles[idx]}"""

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
    
def mean_without_outliers(array):
    mean = array[(array >= array.quantile(0.01)) & (array <= array.quantile(0.99))].mean()
    return mean
