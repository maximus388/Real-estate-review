# Скрипт для генерации аналитических обзоров вторичной недвижимости по городам сети компании Этажи



import os
from import_data import *
from clean import * 
from analysis import *
from review import *
import progressbar
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings("ignore")
from docx2pdf import convert



os.chdir(r"C:/Users/ws-tmn-an-15/Desktop/Харайкин М.А/Python документы/ПРОЕКТ/")

GOOGLE_SHEET_ID_MAIN = '1C6g54M6uOPuuIzqeNJfzT8kDR5VDVb5BrPeCuha8L_E'
SERVICE_FILE = 'C:\\Users\\ws-tmn-an-15\\Desktop\\Харайкин М.А\\Python документы\\python-automation-script-jupyter-notebook-266007-21fda3e2971a.json'

YEAR = 2023
MONTH = 6

locale.setlocale(locale.LC_ALL, '')

START = datetime.datetime(YEAR, MONTH, 1)
end = START + datedelta.MONTH - datetime.timedelta(seconds = 1)
END = end.strftime("%Y-%m-%d %H:%M:%S")
CURRENT_DATE = (START + datedelta.MONTH)
PREVIOUS_MONTH = START
PREVIOUS_YEAR = (START + datedelta.MONTH - datedelta.YEAR)

FILE_DATE = START.strftime("%Y-%m-%d")

MIN_OBJ_SALE_CNT = 100
MIN_OBJ_SOLD_CNT = 20

CITIES_LIST_FINAL = []
CITIES_FEW_SOLD = []

COLORS = ['#FEA19E', '#FF6F6A', '#FC403A', '#E6332A'] 

data = Import_data(SERVICE_FILE, GOOGLE_SHEET_ID_MAIN, CURRENT_DATE, PREVIOUS_MONTH, PREVIOUS_YEAR, FILE_DATE)



"""Импорт данных"""

print('Авторизация...')
gc = data.auth
print('Открытие основного гугл-документа...')
SH = data.open_doc(gc, GOOGLE_SHEET_ID_MAIN)

print('Получение списка городов...')
CITIES_LIST = data.get_cities_id_list(SH)

print('Подключение к БД RIES...')
RIES_CON = db_connect("RIES")

print('Выгрузка объединения районов...')
DF_DISTRICTS = data.get_district_db(RIES_CON, CITIES_LIST, SH)

print('Выгрузка районов вне города...')
DF_DISTRICTS_NOT_IN_CITY = data.districts_not_in_city(SH)

print('Выгрузка материалов стен...')
DF_WALLS = get_df_walls(SH)

print('Подключение к БД BI...')
BI_CON = db_connect("BI")

print('Выгрузка данных по предложению...')
DF_SALE = data.get_sale_data(BI_CON, CITIES_LIST, CURRENT_DATE, PREVIOUS_MONTH, PREVIOUS_YEAR, FILE_DATE)

print('Выгрузка данных по продажам...')
DF_SOLD = data.get_sold_data(RIES_CON, CITIES_LIST, START, END, FILE_DATE)

print('Выгрузка объектов с неопределенными материалами стен...')
DF_UNDEF_WALLS = undef_walls(DF_SALE, DF_WALLS)




"""Очистка данных"""

print('Очистка данных по предложению...')
DF_SALE_CLEANED = clean('SALE', DF_SALE, DF_WALLS, DF_DISTRICTS, CITIES_LIST)

print('Очистка данных по продажам...')
DF_SOLD_CLEANED = clean('SOLD', DF_SOLD, DF_WALLS, DF_DISTRICTS, CITIES_LIST)


"""Анализ"""

print('Получение карты городов, по которым нужно делать обзоры...')
DF_RESTR_MAP = get_cities_restriction_map(DF_SALE_CLEANED, DF_SOLD_CLEANED, MIN_OBJ_SALE_CNT, MIN_OBJ_SOLD_CNT, CITIES_LIST)

# CITIES_LIST_FINAL = list(DF_RESTR_MAP[DF_RESTR_MAP['need_review'] == 1]['city'][:])
CITIES_LIST_FINAL = ["Тюмень", "Тобольск", "Сургут"]
CITY_IMAGES = os.listdir(os.path.join(os.getcwd(), "Обложки"))
DF_CITY_IMAGES = pd.DataFrame({'city': [i[:-4] for i in CITY_IMAGES], 'image_dir': [os.path.join(os.getcwd(), "Обложки") + '\\' + i for i in CITY_IMAGES]})
print(f'Формирование обзоров...')
with progressbar.ProgressBar(max_value = len(CITIES_LIST_FINAL)) as bar:
    for city in CITIES_LIST_FINAL:
        data_review = City_analysis(city, gc, DF_SALE_CLEANED, DF_SOLD_CLEANED, COLORS, START, MONTH, YEAR)
            
        supply_cnt_city_info = data_review.get_supply_cnt_city_info(city)
        supply_cnt_by_city_info = data_review.get_supply_cnt_by_room_city_info(city)
        supply_cnt_by_district_city_info = data_review.get_supply_cnt_by_district_city_info(city)
        sale_mean_price_info = data_review.get_sale_mean_price_info(city)
        sale_mean_price_dynamic = data_review.get_sale_mean_price_dynamic(city, gc)
        
        sale_mean_price_first_graph_num = 2 if len(sale_mean_price_dynamic['DF_SALE_MEAN_PRICE_M2_1']) > 2 else 1        
    
        sale_mean_price_by_room_info = data_review.get_sale_mean_price_by_room_info(city, sale_mean_price_first_graph_num)
        sale_mean_price_by_district_info = data_review.get_sale_mean_price_by_district_info(city, sale_mean_price_first_graph_num)
        if city in DF_RESTR_MAP[DF_RESTR_MAP['need_sold_review'] == 1]['city'].to_list():
            df_sold_city = data_review.get_df_sold_city(city)
            sold_mean = data_review.get_sold_mean(city)
            sold_supply_cnt_by_room = data_review.get_sold_supply_cnt_by_room(city)    
        
        document = Document()
    
        style = document.styles['Normal']
        font = style.font
        font.name = 'Circe'
        font.size = Pt(12)
        
        sections = document.sections
        
        first_section = document.sections[0]
        first_section.top_margin = Cm(1.27)
        first_section.bottom_margin = Cm(1.27)
        first_section.left_margin = Cm(0)
        first_section.right_margin = Cm(0)
        
        
            
        def apply_paragraph_formatting(paragraph):
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        #
        # Логотип
        #
#         picture = document.add_picture(os.path.join(os.getcwd(), "Элементы дизайна", '1_Логотипы', '1_Логотипы', 'png', 'Лого в плашке + охранное поле.png'), width=Inches(2))
#         last_paragraph2 = document.paragraphs[-1] 
#         last_paragraph2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        #
        #
        #    
        # Название документа
        #
        #
        #
        head_main = document.add_heading()
        head_main_run = head_main.add_run(f'Краткий обзор вторичного рынка жилой недвижимости за {START.strftime("%B %Y").lower()} — {city}')
        head_main_run.font.name = 'Circe Extrabold'
        head_main_run.font.color.rgb = RGBColor(16,20,25)
        head_main_run.bold = True
        head_main_run.font.size = Pt(20)
        head_main.paragraph_format.space_before = Pt(0)
        head_main.paragraph_format.left_indent = Inches(0.5)
        head_main.paragraph_format.right_indent = Inches(0.5)
        head_main.alignment = 0
        #
        # Обложка
        #
        picture = document.add_picture(os.path.join(os.getcwd(), "Фото на обложки", city + '.jpg'), width=Inches(8.5))
        last_paragraph2 = document.paragraphs[-1] 
        last_paragraph2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        #
        #
        #
        # Заголовок
        #
        #
        # 
        head_0 = document.add_heading(level = 1)
        head_0_run = head_0.add_run('Основные выводы')
        head_0_format = head_0.paragraph_format
        head_0_format.first_line_indent = Cm(1.25)
        head_0_run.font.name = 'Circe'
        head_0_run.font.color.rgb = RGBColor(16,20,25)
        head_0_run.alignment = 0
        #
        #
        #  
        # Пункт 0.1
        #
        #
        #
        TOP_DISTRICTS_LIST = list(supply_cnt_by_district_city_info['df_sale_supply_cnt_by_district2']['district_new'])
        if len(TOP_DISTRICTS_LIST) >= 3:
            para_0_1 = document.add_paragraph(f"""
                За {START.strftime("%B %Y").lower()} в городе {city} наибольшим спросом у покупателей пользовались квартиры в следующих районах: {', '.join(TOP_DISTRICTS_LIST)};""".strip(), style = 'List Bullet')
            para_0_1.paragraph_format.space_before = Pt(12)
            para_0_1.paragraph_format.left_indent = Inches(0.5)
            para_0_1.paragraph_format.right_indent = Inches(0.5)
            para_0_1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            pass
        #
        #
        #  
        # Пункт 0.2
        #
        #
        #
        SALE_MEAN_PRICE_M2_CITY_CURRENT = int(round(float(1000 * sale_mean_price_info['df_sale_mean_price_m2_current']['price_m2']), 0))
        SALE_MEAN_PRICE_M2_CITY_MONTH_AGO = int(round(float(1000 * sale_mean_price_info['df_sale_mean_price_m2_month_ago']['price_m2']), 0))
        try:
            if len(sale_mean_price_info['df_sale_mean_price_m2_year_ago']) > 0:
                SALE_MEAN_PRICE_M2_CITY_YEAR_AGO = int(round(float(1000 * sale_mean_price_info['df_sale_mean_price_m2_year_ago']['price_m2']), 0))
            else:
                SALE_MEAN_PRICE_M2_CITY_YEAR_AGO = 0
        except TypeError:
            pass
        if DF_RESTR_MAP[DF_RESTR_MAP['city'] == city]['need_sold_review'].item() == 1:
            SOLD_MEAN_PRICE_M2_CITY_CURRENT = int(round(float(1000 * sold_mean['df_sold_mean_price_m2_current']['price_m2']), 0))
        
        if len(TOP_DISTRICTS_LIST) >= 3:
            if DF_RESTR_MAP[DF_RESTR_MAP['city'] == city]['need_sold_review'].item() == 1:
                if abs(SALE_MEAN_PRICE_M2_CITY_CURRENT/SALE_MEAN_PRICE_M2_CITY_MONTH_AGO - 1) >= 0.00005:
                    para_0_2_month_ago = " ({:.2%}".format((SALE_MEAN_PRICE_M2_CITY_CURRENT/SALE_MEAN_PRICE_M2_CITY_MONTH_AGO - 1)).replace('.', ',') + ' в сравнении с прошлым месяцем)'
                else:
                    para_0_2_month_ago = ''
                para_0_2 = document.add_paragraph(f"""
                    Удельная цена предложения составила {'{:,}'.format(SALE_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                    } руб. за кв. м.{para_0_2_month_ago}, а удельная цена продаж — {'{:,}'.format(SOLD_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')} руб. за кв. м.;""".strip(), style = 'List Bullet')
                    
            else:
                if abs(SALE_MEAN_PRICE_M2_CITY_CURRENT/SALE_MEAN_PRICE_M2_CITY_MONTH_AGO - 1) >= 0.00005:
                    para_0_2_month_ago = " ({:.2%}".format((SALE_MEAN_PRICE_M2_CITY_CURRENT/SALE_MEAN_PRICE_M2_CITY_MONTH_AGO - 1)).replace('.', ',') + ' в сравнении с прошлым месяцем)'
                else:
                    para_0_2_month_ago = ''
                para_0_2 = document.add_paragraph(f"""
                    Удельная цена предложения составила {'{:,}'.format(SALE_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                    } руб. за кв. м.{para_0_2_month_ago};""".strip(), style = 'List Bullet')
            para_0_2.paragraph_format.left_indent = Inches(0.5)
            para_0_2.paragraph_format.right_indent = Inches(0.5)
            para_0_2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY            
        else:
            if DF_RESTR_MAP[DF_RESTR_MAP['city'] == city]['need_sold_review'].item() == 1:
                para_0_2_alt_start = document.add_paragraph(f"""
                    За {START.strftime("%B %Y").lower()} в городе {city} удельная цена предложения составила {'{:,}'.format(SALE_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                    } руб. за кв. м., а удельная цена продаж — {'{:,}'.format(SOLD_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')} руб. за кв. м.;""".strip(), style = 'List Bullet')
            else:
                para_0_2_alt_start = document.add_paragraph(f"""
                    За {START.strftime("%B %Y").lower()} в городе {city} удельная цена предложения составила {'{:,}'.format(SALE_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                    } руб. за кв. м.;""".strip(), style = 'List Bullet')
            para_0_2_alt_start.paragraph_format.left_indent = Inches(0.5)
            para_0_2_alt_start.paragraph_format.right_indent = Inches(0.5)
            para_0_2_alt_start.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        #
        #
        #  
        # Пункт 0.3
        #
        #
        #
        para_0_3 = document.add_paragraph(f"""
            В сравнении с прошедшим месяцем удельная цена предложения {value_change("Цена", SALE_MEAN_PRICE_M2_CITY_CURRENT, SALE_MEAN_PRICE_M2_CITY_MONTH_AGO)
            } на {'{:,}'.format(abs(SALE_MEAN_PRICE_M2_CITY_CURRENT - SALE_MEAN_PRICE_M2_CITY_MONTH_AGO)).replace(',',' ')
            } руб. за 1 кв. м.;""".strip(), style = 'List Bullet')
        para_0_3.paragraph_format.left_indent = Inches(0.5)
        para_0_3.paragraph_format.right_indent = Inches(0.5)
        para_0_3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        #
        #
        #  
        # Пункт 0.4
        #
        #
        #
        SALE_SUPPLY_CNT_CITY_CURRENT = int(float(supply_cnt_city_info['df_sale_supply_cnt_current']['id']))
        SALE_SUPPLY_CNT_CITY_MONTH_AGO = int(float(supply_cnt_city_info['df_sale_supply_cnt_prev_month']['id']))
        try:
            if len(supply_cnt_city_info['df_sale_supply_cnt_prev_year']) > 0:
                SALE_SUPPLY_CNT_CITY_YEAR_AGO = int(float(supply_cnt_city_info['df_sale_supply_cnt_prev_year']['id']))
            else:
                SALE_SUPPLY_CNT_CITY_YEAR_AGO = 0
        except TypeError:
            pass
        
        para_0_4_1 = f"""За прошедший месяц объем предложения {value_change('Объем', SALE_SUPPLY_CNT_CITY_CURRENT, SALE_SUPPLY_CNT_CITY_MONTH_AGO)} """
        if value_change('Объем', SALE_SUPPLY_CNT_CITY_CURRENT, SALE_SUPPLY_CNT_CITY_MONTH_AGO) != 'не изменился':
            para_0_4_2 = f"""на {decl(abs(SALE_SUPPLY_CNT_CITY_CURRENT - SALE_SUPPLY_CNT_CITY_MONTH_AGO), ['квартиру', 'квартиры', 'квартир'])} """ 
        else:
            para_0_4_2 = ''
        if DF_RESTR_MAP[DF_RESTR_MAP['city'] == city]['need_sold_review'].item() == 1:
            para_0_4_3 = f"""и составил {decl(SALE_SUPPLY_CNT_CITY_CURRENT, ['объект', 'объекта', 'объектов'])};"""
        else:
            para_0_4_3 = f"""и составил {decl(SALE_SUPPLY_CNT_CITY_CURRENT, ['объект', 'объекта', 'объектов'])}."""
        
        para_0_4 = document.add_paragraph((para_0_4_1 + para_0_4_2 + para_0_4_3).strip(), style = 'List Bullet')
        para_0_4.paragraph_format.left_indent = Inches(0.5)
        para_0_4.paragraph_format.right_indent = Inches(0.5)
        para_0_4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        #
        #
        #  
        # Пункт 0.5
        #
        #
        #
        if DF_RESTR_MAP[DF_RESTR_MAP['city'] == city]['need_sold_review'].item() == 1:
            DF_SOLD_CLEANED_ONE_CITY_CURRENT_DATE = df_sold_city[(df_sold_city['date'].dt.month == MONTH) & (df_sold_city['date'].dt.year == YEAR)]
        
            DF_SOLD_CLEANED_ONE_CITY_MONTH_AGO = df_sold_city[(df_sold_city['date'].dt.month == (START - datedelta.MONTH).month) & 
                                                                          (df_sold_city['date'].dt.year == (START - datedelta.MONTH).year)] 
            DF_SOLD_CLEANED_ONE_CITY_YEAR_AGO = df_sold_city[(df_sold_city['date'].dt.month == (START - datedelta.YEAR).month) & 
                                                                         (df_sold_city['date'].dt.year == (START - datedelta.YEAR).year)]
            TORG_CURRENT = (DF_SOLD_CLEANED_ONE_CITY_CURRENT_DATE['price'].sum() - DF_SOLD_CLEANED_ONE_CITY_CURRENT_DATE['sold_price'].sum()) / DF_SOLD_CLEANED_ONE_CITY_CURRENT_DATE['price'].sum()
            if TORG_CURRENT > 0:
                para_0_5 = document.add_paragraph(f"""
                Средний предпродажный торг составил {"{:.2%}".format(TORG_CURRENT).replace('.', ',')};""".strip(), style = 'List Bullet')
                para_0_5.paragraph_format.left_indent = Inches(0.5)
                para_0_5.paragraph_format.right_indent = Inches(0.5)
                para_0_5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        #
        #
        #  
        # Пункт 0.6
        #
        #
        #
        if DF_RESTR_MAP[DF_RESTR_MAP['city'] == city]['need_sold_review'].item() == 1:
            EXPOS_CURRENT = mean_without_outliers(DF_SOLD_CLEANED_ONE_CITY_CURRENT_DATE['expos'])
            EXPOS_MONTH_AGO = mean_without_outliers(DF_SOLD_CLEANED_ONE_CITY_MONTH_AGO['expos'])
            EXPOS_YEAR_AGO = mean_without_outliers(DF_SOLD_CLEANED_ONE_CITY_YEAR_AGO['expos'])
            para_0_6 = document.add_paragraph(''.join(f"""Средний срок экспозиции проданных в прошедшем месяце квартир — {"{:.2f}".format(EXPOS_CURRENT).replace('.', ',')} мес.""".splitlines()), style = 'List Bullet')
            para_0_6.paragraph_format.left_indent = Inches(0.5)
            para_0_6.paragraph_format.right_indent = Inches(0.5)
            para_0_6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
        #
        #
        #
        #
        #
        #
        #
        #
        # Анализ предложения
        #
        #
        #
        #
        #
        #
        #
        #
        document.add_section()
        
        head_1 = document.add_heading(level = 1)
        head_1_run = head_1.add_run(f'1. Анализ предложения на вторичном рынке за {START.strftime("%B %Y").lower()} — {city}')
        head_1_format = head_1.paragraph_format
        head_1_format.first_line_indent = Cm(1.25)
        head_1_run.font.name = 'Circe'
        head_1_run.font.color.rgb = RGBColor(16,20,25)
        head_1_run.alignment = 1
        #
        #
        #  
        # Пункт 1.1
        #
        #
        #
        para_1_1 = document.add_paragraph(f"""За {START.strftime("%B %Y").lower()} предложение вторичного рынка жилой недвижимости составило {decl(SALE_SUPPLY_CNT_CITY_CURRENT, ['объект', 'объекта', 'объектов'])
            } за исключением квартир за чертой города и таких типов, как общежития, пансионаты, коммунальные квартиры, а также малоэтажное строительство.""".strip())
        para_1_1.paragraph_format.space_before = Pt(12)
        apply_paragraph_formatting(para_1_1)
        #
        #
        #  
        # Пункт 1.2
        #
        #
        #
        if SALE_SUPPLY_CNT_CITY_YEAR_AGO > 0:
            para_1_2 = document.add_paragraph(f"""Относительно прошлого месяца предложение {value_change("Предложение", SALE_SUPPLY_CNT_CITY_CURRENT, SALE_SUPPLY_CNT_CITY_MONTH_AGO)
                } на {decl(abs(SALE_SUPPLY_CNT_CITY_CURRENT - SALE_SUPPLY_CNT_CITY_MONTH_AGO), ['квартиру', 'квартиры', 'квартир'])
                } или на {"{:.2%}".format(abs((SALE_SUPPLY_CNT_CITY_CURRENT / SALE_SUPPLY_CNT_CITY_MONTH_AGO)-1)).replace('.', ',')
                }. В сравнении с прошлогодними данными предложение {value_change("Предложение", SALE_SUPPLY_CNT_CITY_CURRENT, SALE_SUPPLY_CNT_CITY_YEAR_AGO)
                } на {decl(abs(SALE_SUPPLY_CNT_CITY_CURRENT - SALE_SUPPLY_CNT_CITY_YEAR_AGO), ['квартиру', 'квартиры', 'квартир'])
                } ({"{:.2%}".format((SALE_SUPPLY_CNT_CITY_CURRENT / SALE_SUPPLY_CNT_CITY_YEAR_AGO)-1).replace('.', ',')}).""".strip())
        else:
            para_1_2 = document.add_paragraph(f"""Относительно прошлого месяца предложение {value_change("Предложение", SALE_SUPPLY_CNT_CITY_CURRENT, SALE_SUPPLY_CNT_CITY_MONTH_AGO)
                } на {decl(abs(SALE_SUPPLY_CNT_CITY_CURRENT - SALE_SUPPLY_CNT_CITY_MONTH_AGO), ['квартиру', 'квартиры', 'квартир'])
                } или на {"{:.2%}".format(abs((SALE_SUPPLY_CNT_CITY_CURRENT / SALE_SUPPLY_CNT_CITY_MONTH_AGO)-1)).replace('.', ',')}.""".strip())    
        apply_paragraph_formatting(para_1_2)
        #
        # Рисунок 1.1
        #
        document.add_picture(fr'{os.getcwd()}/Графики/{city}/GRAPH1_1.png', width=Inches(5))
        last_paragraph1 = document.paragraphs[-1] 
        last_paragraph1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        #
        #
        #
        # Пункт 1.3
        #
        #
        #
        df_para_1_3 = supply_cnt_by_city_info['df_sale_supply_cnt_by_room2']
        para_1_3 = document.add_paragraph(f"""На вторичном рынке г. {city} наибольшую долю составили {df_para_1_3['rooms_cnt_new'][df_para_1_3['id'].idxmax()]
            }-комнатные квартиры ({"{:.1%}".format(df_para_1_3['%'].max()).replace('.', ',')}) в количестве {decl(df_para_1_3['id'].max(), ['квартира', 'квартиры', 'квартир'])}. Затем идут {
            df_para_1_3['rooms_cnt_new'][1]}-комнатные квартиры, чей объем предложения составил {decl(df_para_1_3['id'][1], ['квартиру', 'квартиры', 'квартир'])} ({
            "{:.1%}".format(df_para_1_3['%'][1]).replace('.', ',')}). Количество {df_para_1_3['rooms_cnt_new'][2]}-комнатных квартир составило {
            decl(df_para_1_3['id'][2], ['квартира', 'квартиры', 'квартир'])} ({"{:.1%}".format(df_para_1_3['%'][2]).replace('.', ',')}).""".strip())
        apply_paragraph_formatting(para_1_3)
        #
        #
        #
        # Пункт 1.4
        #
        #
        #
        df_para_1_4 = supply_cnt_by_district_city_info['df_sale_supply_cnt_by_district2']
        if len(df_para_1_4) > 1:
            TOP_3_DISTRICTS_BY_SUPPLY = ', '.join(
                [f"""{df_para_1_4.loc[i][2]} — {decl(df_para_1_4.loc[i][3], ['объект', 'объекта', 'объектов'])} ({str(df_para_1_4.loc[i][4]).replace('.', ',')}%)""" for i in range(len(df_para_1_4) if len(df_para_1_4) < 3 else 3)])
            para_1_4 = document.add_paragraph(f"""Наибольшее предложение на вторичном рынке сосредоточено в следующих районах: {TOP_3_DISTRICTS_BY_SUPPLY}.""")
            apply_paragraph_formatting(para_1_4)
        #
        # Рисунок 1.2
        #
        COUNT_DISTRICT_IN_CITY_SUPPLY = len(supply_cnt_by_district_city_info['df_sale_supply_cnt_by_district1']['district_new'])
        if COUNT_DISTRICT_IN_CITY_SUPPLY > 1:
            if 2 >= COUNT_DISTRICT_IN_CITY_SUPPLY <= 5:
                graph_1_2_width = 7.2
            elif 5 > COUNT_DISTRICT_IN_CITY_SUPPLY <= 10:
                graph_1_2_width = 6.8
            else:
                graph_1_2_width = 6.5
            document.add_picture(fr'{os.getcwd()}/Графики/{city}/GRAPH1_2.png', width=Inches(graph_1_2_width))
            last_paragraph2 = document.paragraphs[-1] 
            last_paragraph2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        #
        #
        #
        #
        #
        #
        #
        #
        # Анализ ценовой ситуации
        #
        #
        #
        #
        #
        #
        #
        #
        document.add_section()
    
        head_2 = document.add_heading(level = 1)
        head_2_run = head_2.add_run(f'2. Анализ ценовой ситуации на вторичном рынке за {START.strftime("%B %Y").lower()} — {city}')
        head_2_format = head_2.paragraph_format
        head_2_format.first_line_indent = Cm(1.25)
        head_2_run.font.name = 'Circe'
        head_2_run.font.color.rgb = RGBColor(16,20,25)
        head_2_run.alignment = 1
        #
        #
        #
        # Пункт 2.1
        #
        #
        #
        if SALE_MEAN_PRICE_M2_CITY_YEAR_AGO > 0:
            para_2_1 = document.add_paragraph(f"""Удельная цена предложения за {START.strftime("%B %Y").lower()} в среднем составляла {'{:,}'.format(SALE_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                } руб./кв. м. В сравнении с прошлым месяцем она {value_change('Цена', SALE_MEAN_PRICE_M2_CITY_CURRENT, SALE_MEAN_PRICE_M2_CITY_MONTH_AGO)
                } на {decl(abs(SALE_MEAN_PRICE_M2_CITY_CURRENT - SALE_MEAN_PRICE_M2_CITY_MONTH_AGO), ['рубль', 'рубля', 'рублей'])
                } за 1 кв. м. ({"{:.2%}".format((SALE_MEAN_PRICE_M2_CITY_CURRENT/SALE_MEAN_PRICE_M2_CITY_MONTH_AGO - 1)).replace('.', ',')
                }). В сравнении с прошлым годом {value_change('Рост/падение', SALE_MEAN_PRICE_M2_CITY_CURRENT, SALE_MEAN_PRICE_M2_CITY_YEAR_AGO)
                } на {'{:,}'.format(abs(SALE_MEAN_PRICE_M2_CITY_CURRENT - SALE_MEAN_PRICE_M2_CITY_YEAR_AGO)).replace(',',' ')
                } руб./кв. м. ({"{:.2%}".format((SALE_MEAN_PRICE_M2_CITY_CURRENT/SALE_MEAN_PRICE_M2_CITY_YEAR_AGO - 1)).replace('.', ',')}).""".strip())
        else:
            para_2_1 = document.add_paragraph(f"""Удельная цена предложения за {START.strftime("%B %Y").lower()} в среднем составляла {'{:,}'.format(SALE_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                } руб./кв. м. В сравнении с прошлым месяцем она {value_change('Цена', SALE_MEAN_PRICE_M2_CITY_CURRENT, SALE_MEAN_PRICE_M2_CITY_MONTH_AGO)
                } на {decl(abs(SALE_MEAN_PRICE_M2_CITY_CURRENT - SALE_MEAN_PRICE_M2_CITY_MONTH_AGO), ['рубль', 'рубля', 'рублей'])
                } за 1 кв. м. ({"{:.2%}".format((SALE_MEAN_PRICE_M2_CITY_CURRENT/SALE_MEAN_PRICE_M2_CITY_MONTH_AGO - 1)).replace('.', ',')}).""".strip())
        para_2_1.paragraph_format.space_before = Pt(12)
        apply_paragraph_formatting(para_2_1)
        
        if len(sale_mean_price_dynamic['DF_SALE_MEAN_PRICE_M2_1']) > 2:
            #
            # Рисунок 2.1
            #
            document.add_picture(fr'{os.getcwd()}/Графики/{city}/GRAPH2_1.png', width=Inches(7.5))
            last_paragraph2 = document.paragraphs[-1] 
            last_paragraph2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        #
        #
        #
        # Пункт 2.2
        #
        #
        #
        df_para_2_2 = sale_mean_price_by_room_info['df_sale_mean_price_m2_by_room']
        try:
            MEAN_PRICE_ROOMS_CNT_1_CURRENT = df_para_2_2[(df_para_2_2['date'].dt.month == MONTH) & 
                                                         (df_para_2_2['date'].dt.year == YEAR) & 
                                                         (df_para_2_2['rooms_cnt_new'] == '1')]['price_m2_rounded'].item()
            sentence_2_2_1_1 = f"""Удельная цена 1-комнатных квартир оказалась на уровне {decl(MEAN_PRICE_ROOMS_CNT_1_CURRENT)} руб. за 1 кв. м."""
        except ValueError:
            sentence_2_2_1_1 = ''
            
        try:
            MEAN_PRICE_ROOMS_CNT_1_MONTH_AGO = df_para_2_2[(df_para_2_2['date'].dt.month == (START - datedelta.MONTH).month) & 
                                                           (df_para_2_2['date'].dt.year == (START - datedelta.MONTH).year) & 
                                                           (df_para_2_2['rooms_cnt_new'] == '1')]['price_m2_rounded'].item()
            sentence_2_2_1_2 = f""" Это {value_change("Больше/меньше", MEAN_PRICE_ROOMS_CNT_1_CURRENT, MEAN_PRICE_ROOMS_CNT_1_MONTH_AGO)
                } на {decl(abs(MEAN_PRICE_ROOMS_CNT_1_CURRENT - MEAN_PRICE_ROOMS_CNT_1_MONTH_AGO), ['рубль', 'рубля', 'рублей'])
                }, чем в прошлом месяце ({"{:.2%}".format((MEAN_PRICE_ROOMS_CNT_1_CURRENT/MEAN_PRICE_ROOMS_CNT_1_MONTH_AGO - 1)).replace('.', ',')
                })."""
        except ValueError:
            sentence_2_2_1_2 = ''
            
        try:   
            MEAN_PRICE_ROOMS_CNT_2_CURRENT = df_para_2_2[(df_para_2_2['date'].dt.month == MONTH) & 
                                                         (df_para_2_2['date'].dt.year == YEAR) & 
                                                         (df_para_2_2['rooms_cnt_new'] == '2')]['price_m2_rounded'].item()
            sentence_2_2_2_1 = f""" Удельная цена 2-комнатных квартир составила {decl(MEAN_PRICE_ROOMS_CNT_2_CURRENT)} руб./кв. м."""
        except ValueError:
            sentence_2_2_2_1 = ''
            
        try:     
            MEAN_PRICE_ROOMS_CNT_2_MONTH_AGO = df_para_2_2[(df_para_2_2['date'].dt.month == (START - datedelta.MONTH).month) & 
                                                           (df_para_2_2['date'].dt.year == (START - datedelta.MONTH).year) & 
                                                           (df_para_2_2['rooms_cnt_new'] == '2')]['price_m2_rounded'].item()
            sentence_2_2_2_2 = f""", став {value_change("Больше/меньше", MEAN_PRICE_ROOMS_CNT_2_CURRENT, MEAN_PRICE_ROOMS_CNT_2_MONTH_AGO)
                } на {decl(abs(MEAN_PRICE_ROOMS_CNT_2_CURRENT - MEAN_PRICE_ROOMS_CNT_2_MONTH_AGO))} руб. ({"{:.2%}".format((MEAN_PRICE_ROOMS_CNT_2_CURRENT/MEAN_PRICE_ROOMS_CNT_2_MONTH_AGO - 1)).replace('.', ',')
                })."""
        except ValueError:
            sentence_2_2_2_2 = ''
            
        try:    
            MEAN_PRICE_ROOMS_CNT_3_CURRENT = df_para_2_2[(df_para_2_2['date'].dt.month == MONTH) & 
                                                         (df_para_2_2['date'].dt.year == YEAR) & 
                                                         (df_para_2_2['rooms_cnt_new'] == '3')]['price_m2_rounded'].item()
            sentence_2_2_3_1 = f""" 3-комнатные квартиры имели удельную цену {decl(MEAN_PRICE_ROOMS_CNT_3_CURRENT)} руб./кв. м."""
        except ValueError:
            sentence_2_2_3_1 = ''
            
        try:    
            MEAN_PRICE_ROOMS_CNT_3_MONTH_AGO = df_para_2_2[(df_para_2_2['date'].dt.month == (START - datedelta.MONTH).month) & 
                                                           (df_para_2_2['date'].dt.year == (START - datedelta.MONTH).year) & 
                                                           (df_para_2_2['rooms_cnt_new'] == '3')]['price_m2_rounded'].item()
            sentence_2_2_3_2 = f""" и стали стоить {value_change("Больше/меньше", MEAN_PRICE_ROOMS_CNT_3_CURRENT, MEAN_PRICE_ROOMS_CNT_3_MONTH_AGO)
                } на {decl(abs(MEAN_PRICE_ROOMS_CNT_3_CURRENT - MEAN_PRICE_ROOMS_CNT_3_MONTH_AGO))} руб. ({"{:.2%}".format((MEAN_PRICE_ROOMS_CNT_3_CURRENT/MEAN_PRICE_ROOMS_CNT_3_MONTH_AGO - 1)).replace('.', ',')
                })."""
        except ValueError:
            sentence_2_2_3_2 = ''
            
        try:    
            MEAN_PRICE_ROOMS_CNT_MANY_CURRENT = df_para_2_2[(df_para_2_2['date'].dt.month == MONTH) & 
                                                            (df_para_2_2['date'].dt.year == YEAR) & 
                                                            (df_para_2_2['rooms_cnt_new'] == 'многокомн.')]['price_m2_rounded'].item()
            sentence_2_2_many_1 = f""" Цена за квадратный метр многокомнатных квартир — {decl(MEAN_PRICE_ROOMS_CNT_MANY_CURRENT)} руб."""
        except ValueError:
            sentence_2_2_many_1 = ''
        
        if sentence_2_2_many_1 != '':
            try:    
                MEAN_PRICE_ROOMS_CNT_MANY_MONTH_AGO = df_para_2_2[(df_para_2_2['date'].dt.month == (START - datedelta.MONTH).month) & 
                                                                  (df_para_2_2['date'].dt.year == (START - datedelta.MONTH).year) & 
                                                                  (df_para_2_2['rooms_cnt_new'] == 'многокомн.')]['price_m2_rounded'].item()
                sentence_2_2_many_2 = f""", став {value_change("Больше/меньше", MEAN_PRICE_ROOMS_CNT_MANY_CURRENT, MEAN_PRICE_ROOMS_CNT_MANY_MONTH_AGO)
                    } на {decl(abs(MEAN_PRICE_ROOMS_CNT_MANY_CURRENT - MEAN_PRICE_ROOMS_CNT_MANY_MONTH_AGO))} руб. ({"{:.2%}".format((MEAN_PRICE_ROOMS_CNT_MANY_CURRENT/MEAN_PRICE_ROOMS_CNT_MANY_MONTH_AGO - 1)).replace('.', ',')
                    })."""
            except ValueError:
                sentence_2_2_many_2 = ''
        else:
            sentence_2_2_many_2 = ''
            
        full_sentence_2_2 = sentence_2_2_1_1 + sentence_2_2_1_2 + sentence_2_2_2_1 + sentence_2_2_2_2 + sentence_2_2_3_1 + sentence_2_2_3_2 + sentence_2_2_many_1 + sentence_2_2_many_2
        
        para_2_2 = document.add_paragraph(full_sentence_2_2.strip())
        apply_paragraph_formatting(para_2_2)
        #
        # Рисунок 2.2
        #
        document.add_picture(fr'{os.getcwd()}/Графики/{city}/GRAPH2_2.png', width=Inches(6))
        last_paragraph2 = document.paragraphs[-1] 
        last_paragraph2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        #
        #
        #
        # Пункт 2.3
        #
        #
        #
        df_para_2_3 = sale_mean_price_by_district_info['df_sale_mean_price_m2_by_district'].sort_values(by = ['price_m2'], ascending = False).reset_index(drop = True).head(3)
        df_para_2_3['price_m2'] = (df_para_2_3['price_m2']*1000).round(0).astype('int')
        MOST_EXPENSIVE_DISTRICT_1 = df_para_2_3[:1]
        MOST_EXPENSIVE_DISTRICT_2 = df_para_2_3[1:2]
        if len(MOST_EXPENSIVE_DISTRICT_2) == 1:
            MOST_EXPENSIVE_DISTRICT_3 = df_para_2_3[2:3]
            para_2_3_1 = document.add_paragraph(f"""За {START.strftime("%B %Y").lower()} лидером по величине удельной цены за кв. м. на вторичном рынке в г. {city} стал район {MOST_EXPENSIVE_DISTRICT_1['district_new'].item()
                } — {decl(MOST_EXPENSIVE_DISTRICT_1['price_m2'].item(), ['рубль', 'рубля', 'рублей'])
                } за кв. м., объем предложения в данном районе составил {decl(MOST_EXPENSIVE_DISTRICT_1['id'].item(), ['квартиру', 'квартиры', 'квартир'])}.""".strip())
            apply_paragraph_formatting(para_2_3_1)
            if len(MOST_EXPENSIVE_DISTRICT_3) > 0:
                para_2_3_2 = document.add_paragraph(f"""На втором месте по величине удельной цены идет район {MOST_EXPENSIVE_DISTRICT_2['district_new'].item()
                    }, где в среднем удельная цена составила {decl(MOST_EXPENSIVE_DISTRICT_2['price_m2'].item(), ['рубль', 'рубля', 'рублей'])
                    } за кв. м. ({decl(MOST_EXPENSIVE_DISTRICT_2['id'].item(), ['квартира', 'квартиры', 'квартир'])} в предложении). На 3 месте — район {MOST_EXPENSIVE_DISTRICT_3['district_new'].item()
                    } с удельной ценой {decl(MOST_EXPENSIVE_DISTRICT_3['price_m2'].item(), ['рубль', 'рубля', 'рублей'])} за кв. м. ({decl(MOST_EXPENSIVE_DISTRICT_3['id'].item(), ['квартира', 'квартиры', 'квартир'])}).""".strip())
            else:
                para_2_3_2 = document.add_paragraph(f"""На втором месте по величине удельной цены идет район {MOST_EXPENSIVE_DISTRICT_2['district_new'].item()
                    }, где в среднем удельная цена составила {decl(MOST_EXPENSIVE_DISTRICT_2['price_m2'].item(), ['рубль', 'рубля', 'рублей'])
                    } за кв. м. ({decl(MOST_EXPENSIVE_DISTRICT_2['id'].item(), ['квартира', 'квартиры', 'квартир'])} в предложении).""".strip())
            apply_paragraph_formatting(para_2_3_2)
        #
        #
        #
        # Пункт 2.4
        #
        #
        #
        if len(MOST_EXPENSIVE_DISTRICT_2) == 1:
            df_para_2_4 = sale_mean_price_by_district_info['df_sale_mean_price_m2_by_district']
            df_para_2_4['more_than_mean'] = df_para_2_4['price_m2'] > df_para_2_4['price_m2'].mean()
            COUNT_DISTRICT_IN_CITY = len(df_para_2_4)
            COUNT_DISTRICT_IN_CITY_MEAN_ABOVE = len(df_para_2_4[df_para_2_4['more_than_mean'] == False])
            if COUNT_DISTRICT_IN_CITY > 2:
                para_2_4 = document.add_paragraph(f"""{COUNT_DISTRICT_IN_CITY_MEAN_ABOVE} из {COUNT_DISTRICT_IN_CITY} рассматриваемых районов имеют среднюю удельную цену квадратного метра ниже, чем в среднем по рынку.""".strip())
                apply_paragraph_formatting(para_2_4)
        #
        # Рисунок 2.3
        #
        if len(MOST_EXPENSIVE_DISTRICT_2) == 1:
            if 2 > COUNT_DISTRICT_IN_CITY <= 5:
                graph_2_3_width = 7.2
            elif 5 > COUNT_DISTRICT_IN_CITY <= 10:
                graph_2_3_width = 6.8
            else:
                graph_2_3_width = 6.5
            document.add_picture(fr'{os.getcwd()}/Графики/{city}/GRAPH2_3.png', width=Inches(graph_2_3_width))
            last_paragraph2 = document.paragraphs[-1] 
            last_paragraph2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        #
        #
        #
        #
        #
        #
        #
        #
        # Анализ продаж
        #
        #
        #
        #
        #
        #
        #
        #
        if DF_RESTR_MAP[DF_RESTR_MAP['city'] == city]['need_sold_review'].item() == 1:
            #
            #
            #
            # Пункт 3.1
            #
            #
            #
            document.add_section()
    
            head_3 = document.add_heading(level = 1)
            head_3_run = head_3.add_run(f'3. Анализ продаж на вторичном рынке за {START.strftime("%B %Y").lower()} — {city}')
            head_3_format = head_3.paragraph_format
            head_3_format.first_line_indent = Cm(1.25)
            head_3_run.font.name = 'Circe'
            head_3_run.font.color.rgb = RGBColor(16,20,25)
            head_3_run.alignment = 1
            
            df_para_3_1 = sold_supply_cnt_by_room.reset_index()
            try:
                rooms_naming1 = '-комнатных' if len(df_para_3_1['rooms_cnt_new'][0]) == 1 else ''
            except:
                pass
            try:
                rooms_naming2 = '-комнатных' if len(df_para_3_1['rooms_cnt_new'][1]) == 1 else ''
            except:
                pass
            try:
                rooms_naming3 = '-комнатных' if len(df_para_3_1['rooms_cnt_new'][2]) == 1 else ''
            except:
                pass
            try:
                para_3_1 = document.add_paragraph(f"""За {START.strftime("%B %Y").lower()} в структуре продаж {str(df_para_3_1['%'][0]).replace('.', ',')
                    }% покупателей сделали выбор в пользу {df_para_3_1['rooms_cnt_new'][0]
                    }{rooms_naming1} квартир. Доля продаж {df_para_3_1['rooms_cnt_new'][1]}{rooms_naming2} квартир составила {str(df_para_3_1['%'][1]).replace('.', ',')
                    }%, а {df_para_3_1['rooms_cnt_new'][2]}{rooms_naming3} — {str(df_para_3_1['%'][2]).replace('.', ',')
                    }%.""".strip())
            except:
                try:
                    para_3_1 = document.add_paragraph(f"""За {START.strftime("%B %Y").lower()} в структуре продаж {str(df_para_3_1['%'][0]).replace('.', ',')
                        }% покупателей сделали выбор в пользу {df_para_3_1['rooms_cnt_new'][0]
                        }{rooms_naming1} квартир. Доля продаж {df_para_3_1['rooms_cnt_new'][1]}{rooms_naming2} квартир составила {str(df_para_3_1['%'][1]).replace('.', ',')
                        }%.""".strip())
                except:
                    para_3_1 = document.add_paragraph(f"""За {START.strftime("%B %Y").lower()} в структуре продаж {str(df_para_3_1['%'][0]).replace('.', ',')
                            }% покупателей сделали выбор в пользу {df_para_3_1['rooms_cnt_new'][0]
                            }{rooms_naming1} квартир.""".strip())                
            para_3_1.paragraph_format.space_before = Pt(12)
            apply_paragraph_formatting(para_3_1)
            #
            #
            #
            # Пункт 3.2
            #
            #
            #
            df_para_3_2 = sold_mean['df_sold_mean_values']
            SOLD_MEAN_AREA_TOTAL_CURRENT = df_para_3_2[(df_para_3_2['date'].dt.month == MONTH) & 
                                                       (df_para_3_2['date'].dt.year == YEAR)]['area_total', 'mean'].round(2).item()
            try:
                SOLD_MEAN_AREA_TOTAL_YEAR_AGO = df_para_3_2[(df_para_3_2['date'].dt.month == (START - datedelta.YEAR).month) & 
                                                            (df_para_3_2['date'].dt.year == (START - datedelta.YEAR).year)]['area_total', 'mean'].round(2).item()
                para_3_2 = document.add_paragraph(f"""Удельная цена продажи вторичного жилья составила {'{:,}'.format(SOLD_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                    } руб. за квадратный метр, а средняя площадь закрепилась на уровне {str(SOLD_MEAN_AREA_TOTAL_CURRENT).replace('.', ',')
                    } кв. м. ({str(SOLD_MEAN_AREA_TOTAL_YEAR_AGO).replace('.', ',')} кв. м. за прошлый год).""".strip())
            except ValueError: 
                para_3_2 = document.add_paragraph(f"""Удельная цена продажи вторичного жилья составила {'{:,}'.format(SOLD_MEAN_PRICE_M2_CITY_CURRENT).replace(',',' ')
                    } руб. за квадратный метр, а средняя площадь закрепилась на уровне {str(SOLD_MEAN_AREA_TOTAL_CURRENT).replace('.', ',')
                    } кв. м.""".strip())
            apply_paragraph_formatting(para_3_2)
            #
            #
            #
            # Пункт 3.3
            #
            #
            #
            if str(EXPOS_YEAR_AGO) == '<NA>':
                para_3_3 = document.add_paragraph(f"""Средний срок экспозиции проданных квартир — {"{:.2f}".format(EXPOS_CURRENT).replace('.', ',')
                    } мес. Он {value_change('Объем', EXPOS_CURRENT, EXPOS_MONTH_AGO)} на {"{:.2f}".format(abs(EXPOS_CURRENT - EXPOS_MONTH_AGO)).replace('.', ',') 
                    } месяца по сравнению с прошлым периодом.""".strip())
            else:
                if str(EXPOS_MONTH_AGO) == '<NA>':
                    para_3_3 = document.add_paragraph(f"""Средний срок экспозиции проданных квартир — {"{:.2f}".format(EXPOS_CURRENT).replace('.', ',')
                        } мес. Он {value_change('Объем', EXPOS_CURRENT, EXPOS_YEAR_AGO)} на {"{:.2f}".format(abs(EXPOS_CURRENT - EXPOS_YEAR_AGO)).replace('.', ',')
                        } месяца по сравнению с уровнем прошлого года.""".strip())
                else:
                    para_3_3 = document.add_paragraph(f"""Средний срок экспозиции проданных квартир — {"{:.2f}".format(EXPOS_CURRENT).replace('.', ',')
                        } мес. Он {value_change('Объем', EXPOS_CURRENT, EXPOS_MONTH_AGO)} на {"{:.2f}".format(abs(EXPOS_CURRENT - EXPOS_MONTH_AGO)).replace('.', ',') 
                        } месяца по сравнению с прошлым периодом и {value_change('Объем', EXPOS_CURRENT, EXPOS_YEAR_AGO)} на {"{:.2f}".format(abs(EXPOS_CURRENT - EXPOS_YEAR_AGO)).replace('.', ',')
                        } месяца по сравнению с уровнем прошлого года.""".strip())
            apply_paragraph_formatting(para_3_3)
            #
            #
            #
            # Пункт 3.4
            #
            #
            #
            df_para_3_4 = sold_mean['df_sold_mean_values']
            SOLD_MEAN_PRICE_CITY_CURRENT = df_para_3_4[(df_para_3_4['date'].dt.month == MONTH) & 
                                                       (df_para_3_4['date'].dt.year == YEAR)]['price', 'mean'].round(-3).item()
            try:
                SOLD_MEAN_PRICE_CITY_MONTH_AGO = df_para_3_4[(df_para_3_4['date'].dt.month == (START - datedelta.MONTH).month) & 
                                                             (df_para_3_4['date'].dt.year == (START - datedelta.MONTH).year)]['price', 'mean'].round(-3).item()
                para_3_4 = document.add_paragraph(f"""Средняя полная цена продажи на вторичном рынке составила {'{:,}'.format(SOLD_MEAN_PRICE_CITY_CURRENT).replace(',',' ')
                    } рублей, став {value_change('Больше/меньше', SOLD_MEAN_PRICE_CITY_CURRENT, SOLD_MEAN_PRICE_CITY_MONTH_AGO)} на {'{:,}'.format(abs(SOLD_MEAN_PRICE_CITY_CURRENT - SOLD_MEAN_PRICE_CITY_MONTH_AGO)).replace(',',' ')
                    } рублей сравнении с прошлым месяцем.""".strip())
            except ValueError:
                 para_3_4 = document.add_paragraph(f"""Средняя полная цена продажи на вторичном рынке составила {'{:,}'.format(SOLD_MEAN_PRICE_CITY_CURRENT).replace(',',' ')
                    } рублей.""".strip())
            apply_paragraph_formatting(para_3_4)
            #
            # Рисунок 3.1
            #      
            document.add_picture(fr'{os.getcwd()}/Графики/{city}/GRAPH3_1.png', width=Inches(5))
            last_paragraph2 = document.paragraphs[-1] 
            last_paragraph2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            pass
        
        for section in sections[1:]:
            section.top_margin = Cm(1.27)
            section.bottom_margin = Cm(1.27)
            section.left_margin = Cm(1.27)
            section.right_margin = Cm(1.5)
        
        new_section = document.add_section()
        new_section.left_margin = Cm(0.0)
        new_section.right_margin = Cm(0.0)
        new_section.top_margin = Cm(0.0)
        new_section.bottom_margin = Cm(0.0)
        try:
            os.makedirs(fr'{os.getcwd()}/Обзоры/{START.strftime("%Y")}/{START.strftime("%m")}. {START.strftime("%B")}/')
        except FileExistsError:
            pass
        document.add_picture(fr'{os.getcwd()}/Обзоры/Обложка.png', width=Inches(8.5), height = Inches(11))
        
        document.save(os.path.join(os.getcwd(), "Обзоры", START.strftime("%Y"), f'{START.strftime("%m")}. {START.strftime("%B")}', f'{city} - обзор вторичного рынка за {START.strftime("%B %Y").lower()}.docx'))
        
        bar.update(CITIES_LIST_FINAL.index(city))
        
# for review in [i for i in os.listdir(os.path.join(os.getcwd(), "Обзоры", START.strftime("%Y"), f'{START.strftime("%m")}. {START.strftime("%B")}')) if '.docx' in i]:
#     convert(os.path.join(os.getcwd(), "Обзоры", START.strftime("%Y"), f'{START.strftime("%m")}. {START.strftime("%B")}', review), 
#     os.path.join(os.getcwd(), "Обзоры", START.strftime("%Y"), f'{START.strftime("%m")}. {START.strftime("%B")}', review[:-4]) + 'pdf')
#     os.remove(os.path.join(os.getcwd(), "Обзоры", START.strftime("%Y"), f'{START.strftime("%m")}. {START.strftime("%B")}', review))        

print('Готово!')
