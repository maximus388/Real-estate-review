import numpy as np
import pandas as pd


"""
            Функции очистки импортированных данных
"""

def delete_outliers(df):

    """
    Очистка выбросов (значений до 1 и после 99 процентиля)
    """
    
    df = df[(df['price'] >= df['price'].quantile(0.05))
          & (df['price'] <= df['price'].quantile(0.95))
          & (df['area_total'] >= df['area_total'].quantile(0.05))
          & (df['area_total'] <= df['area_total'].quantile(0.95))
          & (df['price_m2'] >= df['price_m2'].quantile(0.05)) 
          & (df['price_m2'] <= df['price_m2'].quantile(0.95))
         ]
    return df


def undef_districts(SALE_or_SOLD, df, df_walls, df_districts, df_districts_not_in_city, FILE_DATE):                   # Объекты с нераспределенными районами

    """
    Получение объектов с неопределенными районами
    """

    df['rooms_cnt_new'] = np.where(df['rooms_cnt'] <=3, df['rooms_cnt'], 'многокомн.')
    df = df.merge(df_walls, left_on = ['wall'], right_on = ['wall_old'], how = 'left')
    df = df[(df['wall_new'] != 'дерево') & (~df['wall'].isnull())]
    df = df.merge(df_districts, left_on = ['city', 'district'], right_on = ['city', 'district_old'], how = 'left')
    if SALE_or_SOLD == 'SALE':
        df_undef_districts = pd.pivot_table(df[(df['district_in'].isnull()) & (df['date'] == FILE_DATE)], 
                                            values = ['id'], index = ['city', 'district'], aggfunc = 'count').sort_values(by = ['id'], ascending = False).reset_index()
        df_undef_districts = df_undef_districts.merge(df_districts_not_in_city, left_on = ['city', 'district'], right_on = ['city', 'district'], how = 'left')
        df_undef_districts = df_undef_districts[df_undef_districts['exclude'].isnull()]
        df_undef_districts.drop(['exclude'], axis = 1, inplace = True)
    else:
        df_undef_districts = pd.pivot_table(df[(df['district_in'].isnull())], values = ['id'], index = ['city', 'district'], aggfunc = 'count').sort_values(by = ['id'], ascending = False).reset_index() 
    return df_undef_districts


def undef_walls(df, df_walls):

    """
    Получение объектов с неопределенными материалами стен
    """
    
    df['rooms_cnt_new'] = np.where(df['rooms_cnt'] <=3, df['rooms_cnt'], 'многокомн.')
    df = df.merge(df_walls, left_on = ['wall'], right_on = ['wall_old'], how = 'left')
    df = df[df['wall_new'] != 'дерево'] 
    df_undef_walls = df[df['wall_new'].isnull()]
    return df_undef_walls
                    
                    
def clean(SALE_or_SOLD, df, df_walls, df_districts, cities_list):

    """
    Очистка выгруженной выборки данных по предложению или продажам
    """
    
    df['rooms_cnt_new'] = np.where(df['rooms_cnt'] <=3, df['rooms_cnt'], 'многокомн.')
    df = df.merge(df_walls, left_on = ['wall'], right_on = ['wall_old'], how = 'left')
    df = df[(df['wall_new'] != 'дерево') & (~df['wall'].isnull())]
    df = df.merge(df_districts, left_on = ['city', 'district'], right_on = ['city', 'district_old'], how = 'left')
    df = df.drop(columns=['district', 'district_in', 'wall'])
    df = df[~(df['price'].isnull()) & 
            ~(df['price_m2'].isnull()) &
            ~(df['area_total'].isnull()) & 
            ~(df['floors_cnt'].isnull()) & 
            ~(df['rooms_cnt'].isnull())]
    to_int_list = ['id', 'price', 'price_m2', 'area_total', 'sold_price', 'floors_cnt', 'rooms_cnt', 'expos']
    for i in to_int_list:
        try:
            df[i] = df[i].astype('float64').astype('Int64')
        except (TypeError, KeyError):
            pass
    df[['date', 'date_create']] = df[['date', 'date_create']].astype('datetime64')
    if SALE_or_SOLD == 'SOLD':
        df['sold_date'] = df['sold_date'].astype('datetime64')
        df['torg'] = df['torg'].astype('float64')
    df = df[(df['district_new'].notnull()) & (df['rooms_cnt'] != 0) & (df['rooms_cnt'] <= 7) & (df['floors_cnt'] > 1)]
    dates_list = tuple(df['date'].drop_duplicates().dt.strftime('%Y-%m-%d'))
    df_cleaned = []
    for city in cities_list:
        df_city = df[(df['city'] == city)]
        for date in dates_list:
            df_cleaned = pd.DataFrame(df_cleaned).append(delete_outliers(df_city[(df_city['date'] == date)].drop_duplicates()))
    return df_cleaned


def get_cities_restriction_map(df_sale, df_sold, min_obj_sale_cnt, min_obj_sold_cnt, cities_list):

    """
    Вывод итоговой таблицы городов с параметрами обзоров:

        1. need_review — нужен ли обзор (1 - да, 0 - нет)
        2. need_sold_review — нужен ли блок продаж в обзоре (1 - да, 0 - нет)

    """
    
    df_cities_restr = []
    for city in cities_list:
        df_sale_city = df_sale[df_sale['city'] == city]
        df_sold_city = df_sold[df_sold['city'] == city]
        obj_sale_cnt = len(df_sale_city[df_sale_city['date'] == df_sale_city['date'].max()])
        obj_sold_cnt = len(df_sold_city[df_sold_city['date'] == df_sold_city['date'].max()])
        cities_restr_dict = {'city': city, 
                             'need_review': 1 if obj_sale_cnt >= min_obj_sale_cnt else 0,
                             'need_sold_review': 1 if (obj_sale_cnt >= min_obj_sale_cnt) & (obj_sold_cnt >= min_obj_sold_cnt) else 0}
        df_cities_restr.append(pd.DataFrame(cities_restr_dict, index = [0]))
    df_cities_restr = pd.concat(df_cities_restr).reset_index(drop = True)
    return df_cities_restr
