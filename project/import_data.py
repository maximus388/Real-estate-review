#!/usr/bin/env python
# coding: utf-8

# In[9]:


"""
        Импорт необходимых для расчета данных из БД РИЭС и БД BI
"""

import os
import os.path
import datetime, datedelta
import locale
import pygsheets
import re, ast
from sqlalchemy import create_engine
import psycopg2
import pandas as pd



def db_connect(RIES_or_BI):
    
    """
    Подключение к БД
    """

    try:
        with open(r'*\connection str '+ RIES_or_BI + r'.txt', 'r') as f:
            CONNECT = f.read()
            CONNECT = ast.literal_eval(re.sub(r'\n','', CONNECT))
    except FileNotFoundError:
        with open(fr'*\connection str '+ RIES_or_BI + r'.txt', 'r') as f:
            CONNECT = f.read()
            CONNECT = ast.literal_eval(re.sub(r'\n','', CONNECT))
    if RIES_or_BI == "RIES":
        try:
            con = create_engine(f"mysql+pymysql://{CONNECT['user']}:{CONNECT['password']}@{CONNECT['host']}/{CONNECT['database']}")
            return con
        except:
            pass
    elif RIES_or_BI == "BI":
        try:
            con = psycopg2.connect(database = CONNECT['database'],user = CONNECT['user'],password = CONNECT['password'],host = CONNECT['host'],port = CONNECT['port'],options = CONNECT['options'])
            return con
        except:
            pass
    else:
        raise DatabaseError("Неверное название БД!")
    

def get_df_walls(sh):

    """
    Импорт списка замен материалов стен из гугл-документа
    """
    
    wks_walls = sh.worksheet_by_title("Материалы стен")
    df_walls = wks_walls.get_as_df(end = "b")
    return df_walls



class Import_data:
    
    """
    Импорт данных по объединениям районов городов и материалов стен, а также выборок по предложению
    и продажам из БД за месяц в рамках проекта автоматизации аналитических обзоров по городам сети"""
    
    def __init__(self, SERVICE_FILE, GOOGLE_SHEET_ID_MAIN, CURRENT_DATE, PREVIOUS_MONTH, PREVIOUS_YEAR, FILE_DATE):
        self.SERVICE_FILE = SERVICE_FILE
        self.GOOGLE_SHEET_ID_MAIN = GOOGLE_SHEET_ID_MAIN

        self.auth = pygsheets.authorize(service_file = SERVICE_FILE)
    
    def open_doc(self, gc, sh_id):

        """
        Открытие гугл-документа
        """
        
        sh = gc.open_by_key(sh_id)
        return sh

    def get_cities_id_list(self, sh):

        """
        Получение списка ID городов
        """
                
        wks = sh.worksheet_by_title("Города")
        df_cities = wks.get_as_df(start = 'a1', end = "l")
        df_cities = df_cities[df_cities['Страна'] == 'Россия']
        cities_list = tuple(df_cities['Город'])
        return cities_list             
    
    def get_district_db(self, con, cities_list, sh):

        """
        Выгрузка районов городах сети и их актуализация в гугл-документе
        """
        
        df_raw = pd.read_sql(f"""<QUERY>""", con = con)
        
        wks_districts_raw = sh.worksheet_by_title("Выгрузка районов из базы")
        wks_districts_raw.clear(start = 'a1')
        wks_districts_raw.set_dataframe(df_raw, 'a1', extend = True, nan = '')
        
        wks_districts = sh.worksheet_by_title("Районы")
        DF_DISTRICTS = wks_districts.get_as_df(end = "e")
        DF_DISTRICTS = DF_DISTRICTS[DF_DISTRICTS['outside_city'] == '']
        DF_DISTRICTS.drop(['outside_city'], axis = 1, inplace = True)
        return DF_DISTRICTS
    
    def districts_not_in_city(self, sh):

        """
        Получение списка исключений районов городов(районы за пределами города
        """
        
        wks_districts_not_in_city = sh.worksheet_by_title("Районы вне города")
        df_districts_not_in_city = wks_districts_not_in_city.get_as_df(end = "c1000000")   
        return df_districts_not_in_city

    
    def get_sale_data(self, con, cities_list, CURRENT_DATE, PREVIOUS_MONTH, PREVIOUS_YEAR, FILE_DATE):

        """
        Выгрузка данных по предложению в городах сети за выбранный месяц
        """

        if os.path.isfile(fr'{os.getcwd()}\Выборки\Предложение\Выборка — %s.parquet' % (FILE_DATE)) == False:
            df = pd.read_sql(f"""<QUERY>""", con = con) 
            df.to_parquet(fr'{os.getcwd()}\Выборки\Предложение\Выборка — %s.parquet' % (FILE_DATE), engine = 'pyarrow')
        df_sale_query = pd.read_parquet(fr'{os.getcwd()}\Выборки\Предложение\Выборка — %s.parquet' % (FILE_DATE), engine = 'pyarrow')
        df_sale_query['city'] = df_sale_query['city'].replace('Новый Волгоград', 'Волгоград')
        return df_sale_query
        
    
    def get_sold_data(self, con, cities_list, START, END, FILE_DATE):

        """
        Выгрузка данных по продажам в городах сети за выбранный месяц
        """
        
        if os.path.isfile(fr'{os.getcwd()}\Выборки\Продажи\Выборка (продажи) — %s.parquet' % (FILE_DATE)) == False:
            df = pd.read_sql(f"""<QUERY>""", con = con)
                
            df.to_parquet(fr'{os.getcwd()}\Выборки\Продажи\Выборка (продажи) — %s.parquet' % (FILE_DATE), engine = 'pyarrow')
        df_sold_query = pd.read_parquet(fr'{os.getcwd()}\Выборки\Продажи\Выборка (продажи) — %s.parquet' % (FILE_DATE), engine = 'pyarrow')
        return df_sold_query

