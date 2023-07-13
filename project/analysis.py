"""
Отрисовка графиков для обзоров
"""


import os
import pandas as pd

import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
import matplotlib.ticker as ticker
import matplotlib.font_manager as font_manager
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import seaborn as sns

import datedelta
import numpy as np
from scipy import stats
import warnings
warnings.filterwarnings( "ignore", module = "matplotlib\..*" )

from import_data import *




"""
Получение шрифтов для текста обзора
"""

font_path  = r'*\Элементы дизайна\1_Circe\1_Circe\Circe Regular.ttf'
font_manager.fontManager.addfont(font_path)
prop = font_manager.FontProperties(fname=font_path)
plt.rcParams['font.family'] = 'Circe'
plt.rcParams['font.family'] = prop.get_name()


"""
Функции для разметки графиков
"""

def remove_borders(ax):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)

def autolabel_vertical(rects, ax):
    for rect in rects:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()/2., height,
                '%d' % int(height), fontsize=11,
                ha='center', va='bottom')
        
        
def autolabel_vertical_float(rects, ax):
    for rect in rects[1:]:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()/2,
                height if height >=0 else 0, 
                height, fontsize=10, ha='center', va='bottom')


"""
Описание логики анализа данных по городу
"""        
class City_analysis:
    
    def __init__(self, gc, city, df_sale_cleaned, df_sold_cleaned, COLORS, START, MONTH, YEAR):
        self.gc = gc
        self.city = city
        self.df_sale_cleaned = df_sale_cleaned
        self.df_sold_cleaned = df_sold_cleaned
        self.COLORS = COLORS
        self.START = START
        self.MONTH = MONTH
        self.YEAR = YEAR

    """
    Получение данных по активным объектам на рынке за месяц
    """
    def get_df_sale_city(self, city):
        return self.df_sale_cleaned[self.df_sale_cleaned['city'] == city]

    """
    Получение данных по проданным объектам на рынке за месяц
    """
    def get_df_sold_city(self, city):
        return self.df_sold_cleaned[self.df_sold_cleaned['city'] == city]
    
    """
    Получение объема предложения за месяц обзора, предыдущий месяц и за месяц в прошлом году
    """    
    def get_supply_cnt_city_info(self, city):
        df_sale_supply_cnt = pd.pivot_table(self.get_df_sale_city(city), values = ['id'], index = ['city', 'date'], aggfunc = {'id': 'count'}).reset_index()
        
        df_sale_supply_cnt_current = df_sale_supply_cnt[(df_sale_supply_cnt['date'].dt.month == self.MONTH) & 
                                                        (df_sale_supply_cnt['date'].dt.year == self.YEAR)]                                          # Объем предложения в текущем месяце
        df_sale_supply_cnt_prev_month = df_sale_supply_cnt[(df_sale_supply_cnt['date'].dt.month == (self.START - datedelta.MONTH).month) & 
                                                           (df_sale_supply_cnt['date'].dt.year == (self.START - datedelta.MONTH).year)]             # Объем предложения за предыдущий месяц
        df_sale_supply_cnt_prev_year = df_sale_supply_cnt[(df_sale_supply_cnt['date'].dt.month == (self.START - datedelta.YEAR).month) & 
                                                           (df_sale_supply_cnt['date'].dt.year == (self.START - datedelta.YEAR).year)]              # Объем предложения за предыдущий год
        
        df_sale_supply_cnt['date'] = df_sale_supply_cnt['date'].dt.strftime('%Y-%m')
        df_sale_supply_cnt1 = df_sale_supply_cnt.set_index('date')
        
        fig1, ax1 = plt.subplots(figsize=(4, 5))
        graph1 = ax1.bar(range(len(df_sale_supply_cnt1.index.values)),
               height = df_sale_supply_cnt1['id'],
               color='#994C4C',
               width = 0.6,
               edgecolor = 'black',
               label = 'data')
        ax1.set(xlabel="Дата",
                ylabel="Кол-во объектов")
        plt.title(f'Объем предложения — {self.city}', y=1.05)
        plt.xticks(range(len(df_sale_supply_cnt1.index.values)), df_sale_supply_cnt1.index.values)
        plt.ylim(0, 1.15 * df_sale_supply_cnt1['id'].max())
        xlocs, xlabs = plt.xticks()
        autolabel_vertical(graph1, ax1)
        plt.close()

        return {'df_sale_supply_cnt_current': df_sale_supply_cnt_current, 
                'df_sale_supply_cnt_prev_month': df_sale_supply_cnt_prev_month,
                'df_sale_supply_cnt_prev_year': df_sale_supply_cnt_prev_year,
                'df_sale_supply_cnt': df_sale_supply_cnt,
                'df_sale_supply_cnt1': df_sale_supply_cnt1}
    
    """
    Получение объема предложения в разрезе комнатности, построение графика
    """    
    def get_supply_cnt_by_room_city_info(self, city):
        df_sale_supply_cnt_by_room = pd.pivot_table(self.get_df_sale_city(city), values = ['id'], index = ['city', 'date', 'rooms_cnt_new'], aggfunc = {'id': 'count'}).reset_index()
        df_sale_supply_cnt_by_room['%'] = df_sale_supply_cnt_by_room['id'] / df_sale_supply_cnt_by_room.groupby(['city', 'date']).transform('sum')['id'] # Объем предложения в городе в разрезе комнатности
        df_sale_supply_cnt_by_room = df_sale_supply_cnt_by_room[(df_sale_supply_cnt_by_room['date'].dt.month == self.MONTH) & (df_sale_supply_cnt_by_room['date'].dt.year == self.YEAR)]
        df_sale_supply_cnt_by_room['date'] = df_sale_supply_cnt_by_room['date'].dt.strftime('%Y-%m')
        df_sale_supply_cnt_by_room1 = df_sale_supply_cnt_by_room.set_index('date').sort_values(by = ['id'], ascending = False)
        df_sale_supply_cnt_by_room2 = df_sale_supply_cnt_by_room1.reset_index()
        df_sale_supply_cnt_by_room2['%'] = df_sale_supply_cnt_by_room2['%'].astype('float')
        
        fig2, ax2 = plt.subplots(figsize=(7, 7), subplot_kw=dict(aspect="equal"))
        fig2.set_facecolor('white')
        
        img = plt.imread(r'*\Элементы дизайна\1_Логотипы\1_Логотипы\png\logo.png', format='png')
        imagebox = OffsetImage(img, zoom=0.08, alpha = 0.2)
        imagebox.image.axes = ax2
        ab = AnnotationBbox(imagebox, (0.5, 0.5), xycoords='axes fraction', frameon=False, bboxprops={'lw':0}, zorder=0)
        ax2.add_artist(ab)
        
        data_fig2 = list(df_sale_supply_cnt_by_room1['%'])
        linewidth = 0.7
        labels = []
        for i in df_sale_supply_cnt_by_room1['rooms_cnt_new']:
            if 'комн' in i:
                labels.append(str(i))
            else:
                labels.append(str(i) + '-комн.')
        patches, texts, autotexts = plt.pie(data_fig2, colors = self.COLORS[0:len(df_sale_supply_cnt_by_room1['rooms_cnt_new'])], startangle=-30, labels = labels,
                                            autopct = '%1.1f%%', pctdistance = 0.8,
                                            wedgeprops = {'width': 0.45, "edgecolor":"w",'linewidth': linewidth},
                                            textprops = {'font': 'Circe', 'fontsize': 15, 'verticalalignment': 'center'},
                )    
        plt.title(f'Рис. 1.1. Структура предложения в разрезе комнатности — {city}', y=-0.15, 
                  font = 'Circe', fontsize = 15
                 )
        try:
            os.makedirs(fr'{os.getcwd()}/Графики/{city}/')
        except FileExistsError:
            pass
        plt.savefig(fr'{os.getcwd()}/Графики/{city}/GRAPH1_1.png', bbox_inches='tight')
        plt.close()

        return {'df_sale_supply_cnt_by_room': df_sale_supply_cnt_by_room,
                'df_sale_supply_cnt_by_room1': df_sale_supply_cnt_by_room1,
                'df_sale_supply_cnt_by_room2': df_sale_supply_cnt_by_room2}
    
    """
    Получение объема предложения в разрезе районов города, построение графика
    """    
    def get_supply_cnt_by_district_city_info(self, city):
        df_sale_supply_cnt_by_district = pd.pivot_table(self.get_df_sale_city(city), values = ['id'], index = ['city', 'date', 'district_new'], aggfunc = 'count').reset_index()
        df_sale_supply_cnt_by_district['%'] = (100 * df_sale_supply_cnt_by_district['id'] / df_sale_supply_cnt_by_district.groupby(['city', 'date']).transform('sum')['id']).astype('float')
        df_sale_supply_cnt_by_district = df_sale_supply_cnt_by_district[(df_sale_supply_cnt_by_district['date'].dt.month == self.MONTH) & (df_sale_supply_cnt_by_district['date'].dt.year == self.YEAR)]
        df_sale_supply_cnt_by_district['date'] = df_sale_supply_cnt_by_district['date'].dt.strftime('%Y-%m')
        df_sale_supply_cnt_by_district = df_sale_supply_cnt_by_district[df_sale_supply_cnt_by_district['id'] > 5]
        df_sale_supply_cnt_by_district = df_sale_supply_cnt_by_district.sort_values(by = ['%'], ascending = True)
        df_sale_supply_cnt_by_district1 = df_sale_supply_cnt_by_district.set_index('date')
        df_sale_supply_cnt_by_district1['%'] = df_sale_supply_cnt_by_district1['%'].round(2)
        df_sale_supply_cnt_by_district2 = df_sale_supply_cnt_by_district1.sort_values(by = ['%'], ascending = False).reset_index().head(3)
        
        df_sale_supply_cnt_by_district1 = df_sale_supply_cnt_by_district1.sort_values(by = ['%'], ascending = False)
        if len(df_sale_supply_cnt_by_district1['district_new']) < 5:
            graph1_2_height = len(df_sale_supply_cnt_by_district1['district_new'])/1.8
        elif len(df_sale_supply_cnt_by_district1['district_new']) < 15:
            graph1_2_height = len(df_sale_supply_cnt_by_district1['district_new'])/2.5
        else:
            graph1_2_height = len(df_sale_supply_cnt_by_district1['district_new'])/3.2
        fig3, ax3 = plt.subplots(figsize=(8, graph1_2_height))
        fig3.set_facecolor('white')
        title_height = (-1.4) / (graph1_2_height)
                
        img = plt.imread(r'*\Элементы дизайна\1_Логотипы\1_Логотипы\png\logo.png', format='png')
        imagebox = OffsetImage(img, zoom=0.2, alpha = 0.2)
        imagebox.image.axes = ax3
        ab = AnnotationBbox(imagebox, (0.5, 0.5), xycoords='axes fraction', frameon=False, bboxprops={'lw':0}, zorder=0)
        ax3.add_artist(ab)
        
        ax3 = sns.barplot(x = df_sale_supply_cnt_by_district1['%']
                          , y = df_sale_supply_cnt_by_district1['district_new']
                          , palette = sns.set_palette(sns.color_palette('blend:#FC403A,#FF6F6A', n_colors = len(df_sale_supply_cnt_by_district1['district_new']))) 
                              if len(df_sale_supply_cnt_by_district1['district_new']) > 3 
                              else sns.set_palette(sns.color_palette('blend:#FC403A,#FC403A', n_colors = len(df_sale_supply_cnt_by_district1['district_new']))) 
                          , orient = 'h')  
        ax3.set(xlabel = None, ylabel = None, xticklabels=[])
        ax3.tick_params(left=False, bottom=False)
        for spine in ax3.spines:
            ax3.spines[spine].set_visible(False)
        for i in ax3.containers:
            ax3.bar_label(i, fmt = "   %.2f%%".format(i), size = 14, font = 'Circe')
        ax3 = plt.yticks(fontsize = 14, font = 'Circe', rotation=0, color='black')
        plt.title(f'Рис. 1.2. Структура предложения в разрезе районов — {city}', 
                  y = title_height, x=(fig3.subplotpars.right + fig3.subplotpars.left)/3, 
                  font = 'Circe', fontsize = 15 if len(df_sale_supply_cnt_by_district1['district_new']) <= 5 else 16)
        try:
            os.makedirs(fr'{os.getcwd()}/Графики/{city}/')
        except FileExistsError:
            pass
        plt.savefig(fr'{os.getcwd()}/Графики/{city}/GRAPH1_2.png', bbox_inches='tight')
        plt.close()
        
        return {'df_sale_supply_cnt_by_district': df_sale_supply_cnt_by_district,
                'df_sale_supply_cnt_by_district1': df_sale_supply_cnt_by_district1,
                'df_sale_supply_cnt_by_district2': df_sale_supply_cnt_by_district2}
    
    """
    Получение удельной цены предложения в городе за месяц обзора, предыдущий месяц и за месяц в прошлом году, построение графика
    """    
    def get_sale_mean_price_info(self, city):
        df_sale_mean_price_m2 = pd.pivot_table(self.get_df_sale_city(city), values = ['price', 'area_total'], index = ['city', 'date'], aggfunc = 'sum').reset_index()
        df_sale_mean_price_m2['price_m2'] = df_sale_mean_price_m2['price'] / df_sale_mean_price_m2['area_total']
        df_sale_mean_price_m2_current = df_sale_mean_price_m2[(df_sale_mean_price_m2['date'].dt.month == self.MONTH) & 
                                                              (df_sale_mean_price_m2['date'].dt.year == self.YEAR)]                                    # Удел цена за последний месяц
        df_sale_mean_price_m2_month_ago = df_sale_mean_price_m2[(df_sale_mean_price_m2['date'].dt.month == (self.START - datedelta.MONTH).month) & 
                                                                (df_sale_mean_price_m2['date'].dt.year == (self.START - datedelta.MONTH).year)]        # Удел цена за предыдущий месяц
        df_sale_mean_price_m2_year_ago = df_sale_mean_price_m2[(df_sale_mean_price_m2['date'].dt.month == (self.START - datedelta.YEAR).month) & 
                                                               (df_sale_mean_price_m2['date'].dt.year == (self.START - datedelta.YEAR).year)]          # Удел цена за предыдущий год
        df_sale_mean_price_m2['date'] = df_sale_mean_price_m2['date'].dt.strftime('%Y-%m')
        df_sale_mean_price_m2_1 = df_sale_mean_price_m2.set_index('date')
        df_sale_mean_price_m2_1['price_m2'] = (1000 * df_sale_mean_price_m2_1['price_m2']).round(0).astype('int64')
        
        fig4, ax4 = plt.subplots(figsize=(4,5))
        ax4.grid(visible = True, axis = 'y', zorder=0)
        graph4 = ax4.bar(range(len(df_sale_mean_price_m2_1.index.values)),
               height = df_sale_mean_price_m2_1['price_m2'],
               color='#994C4C',
               width = 0.6,
               edgecolor = 'black',
               label = 'data', zorder=3)
        ax4.set(xlabel="Дата",
                ylabel="Тыс. руб./м2")
        plt.ylim(0, 1.15 * df_sale_mean_price_m2_1['price_m2'].max())
        plt.title(f'Удельная цена квартир — {city}', y=-0.25)
        plt.xticks(range(len(df_sale_mean_price_m2_1.index.values)), df_sale_mean_price_m2_1.index.values)
        xlocs, xlabs = plt.xticks()
        autolabel_vertical(graph4, ax4)
        plt.close()
        
        return {'df_sale_mean_price_m2': df_sale_mean_price_m2,
                'df_sale_mean_price_m2_current': df_sale_mean_price_m2_current,
                'df_sale_mean_price_m2_month_ago': df_sale_mean_price_m2_month_ago,
                'df_sale_mean_price_m2_year_ago': df_sale_mean_price_m2_year_ago,
                'df_sale_mean_price_m2_1': df_sale_mean_price_m2_1}

    """
    Получение удельной цены предложения в городе за последний год помесячно, построение графика
    """    
    def get_sale_mean_price_dynamic(self, city, gc):
        sh_mean_price_m2 = gc.open_by_key('<google_sheet_id>') 
        wks_mean_price_m2 = sh_mean_price_m2.worksheet_by_title("Данные")
        DF_SALE_MEAN_PRICE_M2 = wks_mean_price_m2.get_as_df(start = 'a1')
        DF_SALE_MEAN_PRICE_M2_1 = DF_SALE_MEAN_PRICE_M2[(DF_SALE_MEAN_PRICE_M2['city'] == city) &
                                                        (DF_SALE_MEAN_PRICE_M2['rooms_cnt_new'] == 'Среднее')]
        DF_SALE_MEAN_PRICE_M2_1['price_m2'] = [float(str(val).replace(',', '.')) for val in DF_SALE_MEAN_PRICE_M2_1['price_m2']]
        DF_SALE_MEAN_PRICE_M2_1 = DF_SALE_MEAN_PRICE_M2_1.tail(13).reset_index(drop = True)
        DF_SALE_MEAN_PRICE_M2_1['date'] = pd.to_datetime(DF_SALE_MEAN_PRICE_M2_1['date']).dt.strftime('%Y-%m')
        DF_SALE_MEAN_PRICE_M2_1['%_growth'] = DF_SALE_MEAN_PRICE_M2_1['price_m2'].pct_change().mul(100).round(2)
        df_nan = DF_SALE_MEAN_PRICE_M2_1[DF_SALE_MEAN_PRICE_M2_1['%_growth'].isna()]
        df_not_nan = DF_SALE_MEAN_PRICE_M2_1[~DF_SALE_MEAN_PRICE_M2_1['%_growth'].isna()]
        df_not_nan = df_not_nan[(np.abs(stats.zscore(df_not_nan['%_growth'])) < 3)]
        DF_SALE_MEAN_PRICE_M2_1 = df_nan.append(df_not_nan).reset_index(drop = True)
        
        if len(DF_SALE_MEAN_PRICE_M2_1['%_growth']) > 2:
            fig7, ax7 = plt.subplots(figsize = (14, 8))
            fig7.set_facecolor('white')
            
            img = plt.imread(r'*\Элементы дизайна\1_Логотипы\1_Логотипы\png\logo.png', format='png')
            imagebox = OffsetImage(img, zoom=0.2, alpha = 0.2)
            imagebox.image.axes = ax7
            ab = AnnotationBbox(imagebox, (0.5, 0.5), xycoords='axes fraction', frameon=False, bboxprops={'lw':0}, zorder=0)
            ax7.add_artist(ab)
            
            plt.xticks(fontsize=15, rotation = 0 if len(DF_SALE_MEAN_PRICE_M2_1['date']) < 7 else 30) 
            col = []
            for val in DF_SALE_MEAN_PRICE_M2_1['%_growth']:
                if val <= 0:
                    col.append('#FC403A') # red 
                else:
                    col.append('#61BD4F') # green
            sns.barplot(x = DF_SALE_MEAN_PRICE_M2_1['date'], y = DF_SALE_MEAN_PRICE_M2_1['%_growth'], palette = col, ax = ax7)
            if DF_SALE_MEAN_PRICE_M2_1['%_growth'].min() >= 0:
                ax7_min_lim = 0
            elif -1.5 <= DF_SALE_MEAN_PRICE_M2_1['%_growth'].min() < 0:
                ax7_min_lim = -2
            else:
                ax7_min_lim = 1.3 * DF_SALE_MEAN_PRICE_M2_1['%_growth'].min()          
            plt.ylim(ax7_min_lim, 1.8 * DF_SALE_MEAN_PRICE_M2_1['%_growth'].max())
            ax8 = ax7.twinx()
            for i in [ax7, ax8]:
                i.spines['top'].set_visible(False)
                i.spines['left'].set_color('#888888')
                i.spines['right'].set_color('#888888')
                i.spines['bottom'].set_color('#888888')
                i.tick_params(axis='both', colors='#888888')
            ax7.set_xlabel(None)
            ax7.set_ylabel("Темп прироста, %", fontsize = 15, font = 'Circe')
            ax7.set_yticklabels(ax7.get_yticks(), size = 15, font = 'Circe', color='black')
            ax7.set_xticklabels(DF_SALE_MEAN_PRICE_M2_1['date'], size = 15, font = 'Circe', color='black')
            for container in ax7.containers:
                ax7.bar_label(container, fmt = "%.1f%%".format(container), size = 15, font = 'Circe')
                
            sns.lineplot(x = DF_SALE_MEAN_PRICE_M2_1['date'], y = DF_SALE_MEAN_PRICE_M2_1['price_m2'], color='#FC403A', marker='o', markersize = 12, linewidth = 3, ax = ax8)
            ax8.yaxis.set_label_position("right")
            
            ax8.set_ylabel("Цена 1 м², тыс. руб.", fontsize = 15, font = 'Circe')
            plt.ylim(0.3 * DF_SALE_MEAN_PRICE_M2_1['price_m2'].min(), 1.1 * DF_SALE_MEAN_PRICE_M2_1['price_m2'].max())
            ax8.set_yticklabels(ax8.get_yticks(), size = 15, font = 'Circe', color='black')
            for i, v in enumerate(DF_SALE_MEAN_PRICE_M2_1['price_m2'].round(1)):
                plt.annotate(v, (DF_SALE_MEAN_PRICE_M2_1.index[i] - 0.25, DF_SALE_MEAN_PRICE_M2_1['price_m2'][i] + 0.03 * DF_SALE_MEAN_PRICE_M2_1['price_m2'].mean()), size = 15, font = 'Circe')
#             sns.reset_defaults() ### Ресет настроек
            ax7 = plt.title(f'Рис. 2.1. Динамика удельной цены предложения — {city}', y = -0.3, fontsize = 18, font = 'Circe')
            try:
                os.makedirs(fr'{os.getcwd()}/Графики/{city}/')
            except FileExistsError:
                pass
            plt.savefig(fr'{os.getcwd()}/Графики/{city}/GRAPH2_1.png', bbox_inches='tight')
            plt.close()
            
        return {'DF_SALE_MEAN_PRICE_M2_1': DF_SALE_MEAN_PRICE_M2_1}

    """
    Получение удельной цены предложения в городе в разрезе комнатности, построение графика
    """    
    def get_sale_mean_price_by_room_info(self, city, sale_mean_price_first_graph_num):
        df_sale_mean_price_m2_by_room = pd.pivot_table(self.get_df_sale_city(city), values = ['price', 'area_total'], index = ['city', 'date', 'rooms_cnt_new'], aggfunc = 'sum').reset_index()
        df_sale_mean_price_m2_by_room['price_m2'] = df_sale_mean_price_m2_by_room['price'] / df_sale_mean_price_m2_by_room['area_total']
        df_sale_mean_price_m2_by_room['price_m2_rounded'] = (df_sale_mean_price_m2_by_room['price_m2'].round(3)*1000).astype('int')
        df_sale_mean_price_m2_by_room1 = df_sale_mean_price_m2_by_room[(df_sale_mean_price_m2_by_room['date'].dt.month == self.MONTH) & (df_sale_mean_price_m2_by_room['date'].dt.year == self.YEAR)]
        df_sale_mean_price_m2_by_room1['date'] = df_sale_mean_price_m2_by_room1['date'].dt.strftime('%Y-%m')
        df_sale_mean_price_m2_by_room1 = df_sale_mean_price_m2_by_room1.set_index('date')
        df_sale_mean_price_m2_by_room1['price_m2'] = df_sale_mean_price_m2_by_room1['price_m2'].round(2)
        df_sale_mean_price_m2_by_room2 = df_sale_mean_price_m2_by_room1.reset_index()
        df_sale_mean_price_m2_by_room2['price_m2'] = (df_sale_mean_price_m2_by_room2['price_m2']*1000).astype('int')
        
        fig5, ax5 = plt.subplots(figsize=(9,5))
        fig5.set_facecolor('white')

        img = plt.imread(r'*\Элементы дизайна\1_Логотипы\1_Логотипы\png\logo.png', format='png')
        imagebox = OffsetImage(img, zoom=0.15, alpha = 0.2)
        imagebox.image.axes = ax5
        ab = AnnotationBbox(imagebox, (0.5, 0.5), xycoords='axes fraction', frameon=False, bboxprops={'lw':0}, zorder=0)
        ax5.add_artist(ab)
        
        labels = []
        for i in df_sale_mean_price_m2_by_room1['rooms_cnt_new']:
            if 'комн' in i:
                labels.append(str(i))
            else:
                labels.append(str(i) + '-комн.')
        ax5 = sns.barplot(x = labels, y = df_sale_mean_price_m2_by_room1['price_m2_rounded'], 
                          palette = sns.set_palette(sns.color_palette('blend:#FC403A,#FF6F6A', n_colors = len(labels))))
        ax5.set_xlabel(None, size = 14, font = 'Circe')
        ax5.set_ylabel('Цена 1 м², руб.', size = 14, font = 'Circe')
        for spine in ax5.spines:
            if spine in ['top', 'right']:
                ax5.spines[spine].set_visible(False)
            ax5.spines[spine].set_color('#888888')
        ax5.tick_params(left=True, bottom=False, axis='both', colors='#888888')
        for i in ax5.containers:
            ax5.bar_label(i, size = 14, font = 'Circe')
        ax5 = plt.yticks(fontsize = 13, font = 'Circe', rotation=0, color='black')
        ax5 = plt.xticks(fontsize = 14, font = 'Circe', rotation=0, color='black')
        plt.ylim(0, 1.1 * df_sale_mean_price_m2_by_room1['price_m2_rounded'].max())
        ax5 = plt.title(f'Рис. 2.{sale_mean_price_first_graph_num}. Удельная цена квартир в разрезе комнатности — {city}', y=-0.3, x=(fig5.subplotpars.right + fig5.subplotpars.left)/2.3, fontsize = 14, font = 'Circe')
        try:
            os.makedirs(fr'{os.getcwd()}/Графики/{city}/')
        except FileExistsError:
            pass
        plt.savefig(fr'{os.getcwd()}/Графики/{city}/GRAPH2_2.png', bbox_inches='tight')
        plt.close()
            
        return {'df_sale_mean_price_m2_by_room': df_sale_mean_price_m2_by_room,
                'df_sale_mean_price_m2_by_room1': df_sale_mean_price_m2_by_room1,
                'df_sale_mean_price_m2_by_room2': df_sale_mean_price_m2_by_room2}
    
    """
    Получение удельной цены предложения в городе в разрезе районов города, построение графика
    """   
    def get_sale_mean_price_by_district_info(self, city, sale_mean_price_first_graph_num):
        df_sale_mean_price_m2_by_district = pd.pivot_table(self.get_df_sale_city(city), values = ['id', 'price', 'area_total'], index = ['city', 'date', 'district_new'], aggfunc ={'id': 'count', 'price': 'sum', 'area_total': 'sum'}).reset_index()
        df_sale_mean_price_m2_by_district['price_m2'] = df_sale_mean_price_m2_by_district['price'] / df_sale_mean_price_m2_by_district['area_total']
        df_sale_mean_price_m2_by_district = df_sale_mean_price_m2_by_district[df_sale_mean_price_m2_by_district['id'] > 5]
        df_sale_mean_price_m2_by_district = df_sale_mean_price_m2_by_district[(df_sale_mean_price_m2_by_district['date'].dt.month == self.MONTH) & (df_sale_mean_price_m2_by_district['date'].dt.year == self.YEAR)]
        df_sale_mean_price_m2_by_district['date'] = df_sale_mean_price_m2_by_district['date'].dt.strftime('%Y-%m')
        df_sale_mean_price_m2_by_district = df_sale_mean_price_m2_by_district.sort_values(by = ['price_m2'], ascending = True)
        df_sale_mean_price_m2_by_district1 = df_sale_mean_price_m2_by_district.set_index('date')
        df_sale_mean_price_m2_by_district1['price_m2'] = (1000*df_sale_mean_price_m2_by_district1['price_m2']).round(0).astype('int')

        df_sale_mean_price_m2_by_district1 = df_sale_mean_price_m2_by_district1.sort_values(by = ['price_m2'], ascending = False)
        if len(df_sale_mean_price_m2_by_district1['district_new']) < 5:
            graph2_3_height = len(df_sale_mean_price_m2_by_district1['district_new'])/1.8
        elif len(df_sale_mean_price_m2_by_district1['district_new']) < 15:
            graph2_3_height = len(df_sale_mean_price_m2_by_district1['district_new'])/2.5
        else:
            graph2_3_height = len(df_sale_mean_price_m2_by_district1['district_new'])/3.2
        fig6, ax6 = plt.subplots(figsize=(8, graph2_3_height))
        title_height = (-1.4) / (graph2_3_height)
        fig6.set_facecolor('white')

        img = plt.imread(r'*\Элементы дизайна\1_Логотипы\1_Логотипы\png\logo.png', format='png')
        imagebox = OffsetImage(img, zoom=0.2, alpha = 0.2)
        imagebox.image.axes = ax6
        ab = AnnotationBbox(imagebox, (0.5, 0.5), xycoords='axes fraction', frameon=False, bboxprops={'lw':0}, zorder=0)
        ax6.add_artist(ab)
        
        ax6 = sns.barplot(x = df_sale_mean_price_m2_by_district1['price_m2']
                          , y = df_sale_mean_price_m2_by_district1['district_new']
                          , palette = sns.set_palette(sns.color_palette('blend:#FC403A,#FF6F6A', n_colors = len(df_sale_mean_price_m2_by_district1['district_new']))) 
                              if len(df_sale_mean_price_m2_by_district1['district_new']) > 3 
                              else sns.set_palette(sns.color_palette('blend:#FC403A,#FC403A', n_colors = len(df_sale_mean_price_m2_by_district1['district_new'])))
                          , orient = 'h')
        ax6.set(xlabel = None, ylabel = None, xticklabels=[])
        ax6.tick_params(left=False, bottom=False)
        for spine in ax6.spines:
            ax6.spines[spine].set_visible(False)
        for i in ax6.containers:
            ax6.bar_label(i, fmt = "   %.0f".format(i), size = 14, font = 'Circe')
        ax6 = plt.yticks(fontsize = 14, font = 'Circe', rotation=0, color='black')
        plt.title(f'Рис. 2.{sale_mean_price_first_graph_num + 1}. Удельная цена квартир в разрезе районов — {city}', 
                  y = title_height, x=(fig6.subplotpars.right + fig6.subplotpars.left)/3, 
                  font = 'Circe', fontsize = 15 if len(df_sale_mean_price_m2_by_district1['district_new']) <= 5 else 16)
        try:
            os.makedirs(fr'{os.getcwd()}/Графики/{city}/')
        except FileExistsError:
            pass
        plt.savefig(fr'{os.getcwd()}/Графики/{city}/GRAPH2_3.png', bbox_inches='tight')
        plt.close()
        
        return {'df_sale_mean_price_m2_by_district': df_sale_mean_price_m2_by_district,
                'df_sale_mean_price_m2_by_district1': df_sale_mean_price_m2_by_district1}

    """
    Получение удельной цены продаж в городе
    """   
    def get_sold_mean(self, city):
        df_sold_mean_values = pd.pivot_table(self.get_df_sold_city(city), values = ['price_m2', 'area_total', 'price'], index = ['city', 'date'], aggfunc = {'area_total': ('mean', 'sum'), 'price': ('mean', 'sum')}).reset_index()
        df_sold_mean_values['price_m2'] = df_sold_mean_values['price', 'sum'] / df_sold_mean_values['area_total', 'sum']
        df_sold_mean_values['price', 'mean'] = (1000*df_sold_mean_values['price', 'mean']).astype('float').round(0).astype('int64')
        df_sold_mean_price_m2_current = df_sold_mean_values[(df_sold_mean_values['date'].dt.month == self.MONTH) & (df_sold_mean_values['date'].dt.year == self.YEAR)]
        return {'df_sold_mean_values': df_sold_mean_values,
                'df_sold_mean_price_m2_current': df_sold_mean_price_m2_current}
        
    """
    Получение удельной цены продаж в городе в разрезе комнатности, построение графика
    """       
    def get_sold_supply_cnt_by_room(self, city):
        df_sold_supply_cnt_by_room = pd.pivot_table(self.get_df_sold_city(city), values = ['id'], index = ['city', 'date', 'rooms_cnt_new'], aggfunc = 'count').reset_index()
        df_sold_supply_cnt_by_room = df_sold_supply_cnt_by_room[(df_sold_supply_cnt_by_room['date'].dt.month == self.MONTH) & (df_sold_supply_cnt_by_room['date'].dt.year == self.YEAR)]
        df_sold_supply_cnt_by_room['%'] = (df_sold_supply_cnt_by_room['id'] / df_sold_supply_cnt_by_room['id'].sum() * 100).round(2)
        df_sold_supply_cnt_by_room['date'] = df_sold_supply_cnt_by_room['date'].dt.strftime('%Y-%m')
        df_sold_supply_cnt_by_room = df_sold_supply_cnt_by_room.set_index('date').sort_values(by = ['%'], ascending = False)
        
        fig9, ax9 = plt.subplots(figsize=(7, 7), subplot_kw=dict(aspect="equal"))
        fig9.set_facecolor('white')

        img = plt.imread(r'*\Элементы дизайна\1_Логотипы\1_Логотипы\png\logo.png', format='png')
        imagebox = OffsetImage(img, zoom=0.08, alpha = 0.2)
        imagebox.image.axes = ax9
        ab = AnnotationBbox(imagebox, (0.5, 0.5), xycoords='axes fraction', frameon=False, bboxprops={'lw':0}, zorder=0)
        ax9.add_artist(ab)
        
        data_fig9 = list(df_sold_supply_cnt_by_room['%'])
        linewidth = 0.7
        labels = []
        for i in df_sold_supply_cnt_by_room['rooms_cnt_new']:
            if 'комн' in i:
                labels.append(str(i))
            else:
                labels.append(str(i) + '-комн.')
        wedges = ax9.pie(data_fig9, colors = self.COLORS[0:len(df_sold_supply_cnt_by_room['rooms_cnt_new'])], startangle=-30, labels = labels,
                                  autopct = '%1.1f%%', pctdistance = 0.8,
                                                wedgeprops = {'width': 0.45, "edgecolor":"w",'linewidth': linewidth},
                                                textprops = {'font': 'Circe', 'fontsize': 15, 'verticalalignment': 'center'},
                        )
        plt.title(f'Рис. 3.1. Структура продаж в разрезе комнатности — {city}', y=-0.15, font = 'Circe', fontsize = 15)
        try:
            os.makedirs(fr'{os.getcwd()}/Графики/{city}/')
        except FileExistsError:
            pass
        plt.savefig(fr'{os.getcwd()}/Графики/{city}/GRAPH3_1.png', bbox_inches='tight')
        plt.close()
    
        return df_sold_supply_cnt_by_room  
