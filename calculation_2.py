import pandas as pd
from pathlib import Path
from collections import defaultdict
import numpy as np
from pandas.io.excel import ExcelWriter
import locale
import time
import openpyxl
import matplotlib as mp
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from matplotlib.ticker import NullLocator


class Calculation():
    def __init__(self):
        #объявление переменных

        self.int_months = None
        self.url_mr = None
        self.order_find = None
        self.output_path = None
        
        
        self.all_ml = None
        self.sklad_ml = None
        self.multIndex = None
        self.excel_table = None
        self.BytesImg = BytesIO()
    
    
    @property
    def df_all(self):
        return self.all_ml
    
    @property
    def df_sklad(self):
        return self.sklad_ml
    
    @property
    def ex_table(self):
        return self.excel_table
    
    @property
    def count_months(self):
        return self.int_months
    
    @count_months.setter
    def count_months(self,value):
        if type(value) == int:
            self.int_months = value
        else:
            raise TypeError("Месяц должен быть числом")
    
    @property
    def path_out(self):
        return self.output_path
    
    @path_out.setter
    def path_out(self, value):
        self.output_path = Path(value)
        if not self.output_path.is_dir():
            self.output_path = None
            raise PermissionError("Нет доступа к папке выхода")
        
    @property
    def order_list(self):
        return self.order_find
    @order_list.setter
    def order_list(self,value):
        self.order_find = value
    
    @property
    def path_base(self):
        return self.url_mr
    
    @path_base.setter
    def path_base(self, value):
        self.url_mr = Path(value)
        if not self.url_mr.is_file():
            self.url_mr = None
            raise PermissionError("Нет доступа к базе")
    
    
    
    def start_work_table(self):
        #Начало выполнения программы
        self.get_data_base()
        self.create_table()
        self.create_plot()
        self.write_to_excel()


    
    def get_data_base(self):
        #Считывание базы данных
        self.all_ml = pd.read_csv(self.url_mr,sep = "^",dtype = {"Маршрутный лист": str}, encoding = "cp1251",parse_dates = ["Дата распечатки"],date_parser= lambda x: pd.to_datetime(x,format = r"%d/%m/%Y").to_period("M"))
        self.sklad_ml = self.all_ml[self.all_ml["Тип сдачи"] == "D"].copy()
        self.sklad_ml["Дата сдачи"] = pd.to_datetime(self.sklad_ml["Дата сдачи"], format = "%d/%m/%y").apply(lambda x: x.to_period("M"))
    
    def create_table(self):
        #Создание таблицы отчета
        def create_info_table(key):
            #Создание датафрейма с  нужными данными
            dt_mr = self.all_ml[self.all_ml["Заказ"] == key]
            sklad = self.sklad_ml[self.sklad_ml["Заказ"] == key]
            date = pd.Timestamp.today().to_period("M")
            return_dict = defaultdict(list)

            for n_month in range(0,self.int_months):
                data = date - n_month
                skl_variable = dict()

                if dt_mr.size:
                    ml_table_this_m = dt_mr[dt_mr["Дата распечатки"] == data]
                    ml_table_below_m = dt_mr[dt_mr["Дата распечатки"]  <= data]

                else:
                    return None

                if sklad.size:
                    sklad_this_m = sklad[sklad["Дата сдачи"] == data]
                    sklad_table_below_m = sklad[sklad["Дата сдачи"] <= data]
                else:
                    sklad_this_m = sklad_table_below_m = sklad

                #1. Расчет наименований дозапуска
                return_dict["Наименований, шт."].append(ml_table_this_m["№ детали"].unique().size)
                return_dict["Деталей, шт."].append(ml_table_this_m["Количество"].sum())
                return_dict["Время, нч."].append(ml_table_this_m["Норм/часы"].sum())

                #2. Расчет полного запуска

                skl_variable.update({"all_Name":ml_table_below_m["№ детали"].unique().size})
                skl_variable.update({"all_detail":ml_table_below_m["Количество"].sum()})
                skl_variable.update({"all_hource":ml_table_below_m["Норм/часы"].sum()})

                return_dict["Наименований, шт."].append(skl_variable["all_Name"])
                return_dict["Деталей, шт."].append(skl_variable["all_detail"])
                return_dict["Время, нч."].append(skl_variable["all_hource"])

                #3. Расчет сдачи за месяц

                return_dict["Наименований, шт."].append(sklad_this_m["№ детали"].unique().size)
                return_dict["Деталей, шт."].append(sklad_this_m["Количество"].sum())
                return_dict["Время, нч."].append(sklad_this_m["Норм/часы"].sum())

                #4. Расчет полной сдачи

                skl_variable.update({"sk_all_Name":sklad_table_below_m["№ детали"].unique().size})
                skl_variable.update({"sk_all_detail":sklad_table_below_m["Количество"].sum()})
                skl_variable.update({"sk_all_hource":sklad_table_below_m["Норм/часы"].sum()})

                return_dict["Наименований, шт."].append(skl_variable["sk_all_Name"])
                return_dict["Деталей, шт."].append(skl_variable["sk_all_detail"])
                return_dict["Время, нч."].append(skl_variable["sk_all_hource"])

                #5. Расчет процентного соотношения

                if skl_variable["all_Name"] == 0:
                    return_dict["Наименований, шт."].append(0)
                else:
                    return_dict["Наименований, шт."].append(skl_variable["sk_all_Name"]/skl_variable["all_Name"]*100)

                if skl_variable["all_detail"] == 0:
                    return_dict["Деталей, шт."].append(0)
                else:
                    return_dict["Деталей, шт."].append(skl_variable["sk_all_detail"]/skl_variable["all_detail"]*100)

                if skl_variable["all_hource"] == 0:
                    return_dict["Время, нч."].append(0)
                else:
                    return_dict["Время, нч."].append(round(skl_variable["sk_all_hource"]/skl_variable["all_hource"]*100,2))

            data_table = pd.DataFrame(return_dict,index = self.multIndex)
            data_table.columns = pd.MultiIndex.from_tuples(list(zip([key] * data_table.columns.size,data_table.columns)))
            return(data_table.T)
        

        def create_mult_ind():
            locale.setlocale(locale.LC_ALL,"ru")
            list_colomns = ["Дозапущено","Полный запуск","Сдано за месяц деталей","Сдано всего деталей","Процент готовности"]
            month_list = [x.strftime("%B %Y") for x in pd.date_range(start = pd.Timestamp.today() - pd.DateOffset(months = self.int_months-1),periods = self.int_months,freq = "M").sort_values(ascending= False).to_list()]
            month_list = [[x]*list_colomns.__len__() for x in month_list]
            month_list = list(np.reshape(month_list,-1))
            array = [month_list,list_colomns*(self.int_months)]
            return pd.MultiIndex.from_tuples(list(zip(*array)))

        self.multIndex = create_mult_ind()
        for key in self.order_find:
            tab = create_info_table(key)
            self.excel_table = pd.concat([self.excel_table,tab])
        self.excel_table = self.excel_table.round(2) 
    
    
    
    
    def old_create_plot(self):
        res_s = defaultdict(dict)
        dates = pd.Series([x for x,y in self.excel_table.columns.to_list()]).drop_duplicates().to_list()
        for order in self.excel_table.index.levels[0]:
            for date in dates:
                tab = self.excel_table.loc[order,date].loc["Время, нч.","Процент готовности"]
                res_s[order][pd.to_datetime(date, format = "%B %Y")] = tab

        res_s = pd.DataFrame(res_s)
        res_s.sort_index(ascending = True,inplace = True)
        fx = plt.figure(figsize = (8,22))
        ax = fx.add_subplot(3,1,1)
        for i in range(res_s.columns.size):
            ax.plot(res_s.index.strftime("%B %y"),res_s.loc[:,res_s.columns[i]], marker = ".", markersize = 12, linewidth = 2, markerfacecolor = (1,1,1),\
                   label = res_s.columns[i])
        ax.tick_params(axis = "x", labelrotation = 45,labelsize = 9)
 
        ax.set_aspect(0.1)
        ax.minorticks_on()
        ax.grid()
        ax.yaxis.grid(which = "major" , linewidth = 1, color = "black")

        ax.grid(which = "minor", linewidth = 0.6, color = "gray")
        ax.xaxis.set_minor_locator(NullLocator())
        ax.set_ylabel("Выполнено %")
        ax.legend(fontsize = 8)

        ax.set_title("Выполнение заказов")



        res_s = defaultdict(dict)
        dates = pd.Series([x for x,y in self.excel_table.columns.to_list()]).drop_duplicates().to_list()
        for order in self.excel_table.index.levels[0]:
            for date in dates:
                tab = self.excel_table.loc[order,date].loc["Время, нч.","Дозапущено"]
                res_s[order][pd.to_datetime(date, format = "%B %Y")] = tab

        res_s = pd.DataFrame(res_s)
        res_s.sort_index(ascending = True,inplace = True)
        # res_s
        res_s.index = res_s.index.strftime("%B %y")
        ax = fx.add_subplot(3,1,2)
        res_s.plot.bar(ax = ax,stacked = True)

        # for i in range(res_s.columns.size):
        #     ax.bar(res_s.index.strftime("%B %y"),res_s.loc[:,res_s.columns[i]],0.5,label = res_s.columns[i])
        ax.tick_params(axis = "x", labelrotation = 45,labelsize = 9)
        ax.minorticks_on()
        ax.grid()
        ax.yaxis.grid(which = "major" , linewidth = 1, color = "black")
        ax.grid(which = "minor", linewidth = 0.6, color = "gray")
        ax.xaxis.set_minor_locator(NullLocator())
        ax.set_ylabel("Норм часы")
        ax.semilogy()
        ax.legend(fontsize = 8)
        ax.set_title("Дозапуск")
        fx.set_facecolor([0.9,0.9,0.9])
        for i in ax.containers:
            ax.bar_label(i,rotation = 0,size = 8,padding = 0,label_type='center')


        res_s = defaultdict(dict)
        dates = pd.Series([x for x,y in self.excel_table.columns.to_list()]).drop_duplicates().to_list()
        for order in self.excel_table.index.levels[0]:
            for date in dates:
                tab = self.excel_table.loc[order,date].loc["Время, нч.","Сдано за месяц деталей"]
                res_s[order][pd.to_datetime(date, format = "%B %Y")] = tab

        res_s = pd.DataFrame(res_s)
        res_s.sort_index(ascending = True,inplace = True)
        # res_s
        res_s.index = res_s.index.strftime("%B %y")
        ax = fx.add_subplot(3,1,3)
        res_s.plot.bar(ax = ax,stacked = True)


        # for i in range(res_s.columns.size):
        #     ax.bar(res_s.index.strftime("%B %y"),res_s.loc[:,res_s.columns[i]],0.5,label = res_s.columns[i])
        ax.tick_params(axis = "x", labelrotation = 45,labelsize = 9)
        ax.minorticks_on()
        ax.grid()
        ax.yaxis.grid(which = "major" , linewidth = 1, color = "black")
        ax.grid(which = "minor", linewidth = 0.6, color = "gray")
        ax.xaxis.set_minor_locator(NullLocator())
        ax.set_ylabel("Норм часы")
        ax.semilogy()
        ax.legend(fontsize = 8)
        ax.set_title("Сдано за месяц")
        for i in ax.containers:
            ax.bar_label(i,rotation = 0,size = 8,padding = 0,label_type='center')
        fx.set_facecolor([0.9,0.9,0.9])
        fx.savefig(self.BytesImg, format="png", dpi = 150, bbox_inches = 'tight')
        fx.clf()
   
    def create_plot(self):
        res_s = defaultdict(dict)
        dates = pd.Series([x for x,y in self.excel_table.columns.to_list()]).drop_duplicates().to_list()
        for order in self.excel_table.index.levels[0]:
            for date in dates:
                tab = self.excel_table.loc[order,date].loc["Время, нч.","Процент готовности"]
                res_s[order][pd.to_datetime(date, format = "%B %Y")] = tab

        res_s = pd.DataFrame(res_s)
        res_s.sort_index(ascending = True,inplace = True)
        fx = plt.figure(figsize = (8,22))
        ax = fx.add_subplot(3,1,1)
        for i in range(res_s.columns.size):
            ax.plot(res_s.index.strftime("%B %y"),res_s.loc[:,res_s.columns[i]], marker = ".", markersize = 12, linewidth = 2, markerfacecolor = (1,1,1),\
                   label = res_s.columns[i])
        ax.tick_params(axis = "x", labelrotation = 45,labelsize = 9)
 
        ax.set_aspect(0.1)
        ax.minorticks_on()
        ax.grid()
        ax.yaxis.grid(which = "major" , linewidth = 1, color = "black")

        ax.grid(which = "minor", linewidth = 0.6, color = "gray")
        ax.xaxis.set_minor_locator(NullLocator())
        ax.set_ylabel("Выполнено %")
        ax.legend(fontsize = 8)

        ax.set_title("Выполнение заказов")



        res_s = defaultdict(dict)
        dates = pd.Series([x for x,y in self.excel_table.columns.to_list()]).drop_duplicates().to_list()
        for order in self.excel_table.index.levels[0]:
            for date in dates:
                tab = self.excel_table.loc[order,date].loc["Время, нч.","Дозапущено"]
                res_s[order][pd.to_datetime(date, format = "%B %Y")] = tab

        res_s = pd.DataFrame(res_s)
        res_s.sort_index(ascending = True,inplace = True)
        # res_s
        res_s.index = res_s.index.strftime("%B %y")
        ax = fx.add_subplot(3,1,2)
        res_s.plot.bar(ax = ax,stacked = True)

        # for i in range(res_s.columns.size):
        #     ax.bar(res_s.index.strftime("%B %y"),res_s.loc[:,res_s.columns[i]],0.5,label = res_s.columns[i])
        ax.tick_params(axis = "x", labelrotation = 45,labelsize = 9)
        ax.minorticks_on()
        ax.grid()
        ax.yaxis.grid(which = "major" , linewidth = 1, color = "black")
        ax.grid(which = "minor", linewidth = 0.6, color = "gray")
        ax.xaxis.set_minor_locator(NullLocator())
        ax.set_ylabel("Норм часы")
#         ax.semilogy()
        ax.legend(fontsize = 8)
        ax.set_title("Дозапуск")
        fx.set_facecolor([0.9,0.9,0.9])
        for i in ax.containers:
            ax.bar_label(i,rotation = 0,size = 8,padding = 0,label_type='center')


        res_s = defaultdict(dict)
        dates = pd.Series([x for x,y in self.excel_table.columns.to_list()]).drop_duplicates().to_list()
        for order in self.excel_table.index.levels[0]:
            for date in dates:
                tab = self.excel_table.loc[order,date].loc["Время, нч.","Сдано за месяц деталей"]
                res_s[order][pd.to_datetime(date, format = "%B %Y")] = tab

        res_s = pd.DataFrame(res_s)
        res_s.sort_index(ascending = True,inplace = True)
        # res_s
        res_s.index = res_s.index.strftime("%B %y")
        ax = fx.add_subplot(3,1,3)
        res_s.plot.bar(ax = ax,stacked = True)


        # for i in range(res_s.columns.size):
        #     ax.bar(res_s.index.strftime("%B %y"),res_s.loc[:,res_s.columns[i]],0.5,label = res_s.columns[i])
        ax.tick_params(axis = "x", labelrotation = 45,labelsize = 9)
        ax.minorticks_on()
        ax.grid()
        ax.yaxis.grid(which = "major" , linewidth = 1, color = "black")
        ax.grid(which = "minor", linewidth = 0.6, color = "gray")
        ax.xaxis.set_minor_locator(NullLocator())
        ax.set_ylabel("Норм часы")
#         ax.semilogy()
        ax.legend(fontsize = 8)
        ax.set_title("Сдано за месяц")
        for i in ax.containers:
            ax.bar_label(i,rotation = 0,size = 8,padding = 0,label_type='center')
        fx.set_facecolor([0.9,0.9,0.9])
        fx.savefig(self.BytesImg, format="png", dpi = 150, bbox_inches = 'tight')
        fx.clf()
        
        
    def write_to_excel(self):
        #Запись таблицы в эксель
        list_color = list()
        t = mp.cm.get_cmap("Pastel1",self.int_months)
        for i in range(t.N):
            rgba = t(i)
            list_color.append(mp.colors.to_hex(rgba))
        int_len_1 = len(self.excel_table)+2
        int_len_2 = len(self.excel_table.columns)+1
        self.excel_table = self.excel_table.style
        for i in self.excel_table.columns.levels[0]:
            self.excel_table = self.excel_table.set_properties(**{"background-color": list_color.pop()}, subset  = [i])    
        import xlsxwriter
        Byte_excel = BytesIO()
        with ExcelWriter(Byte_excel, engine = "xlsxwriter",) as writer:
            self.excel_table.to_excel(writer,sheet_name = "Состояние_МП",freeze_panes= (2,2))
            workbook = writer.book
            worksheet = writer.sheets["Состояние_МП"]
            cell_format = workbook.add_format()
            cell_format.set_border()
            worksheet.conditional_format(xlsxwriter.utility.xl_range(0,0,int_len_1,int_len_2),{"type" : "no_errors","format":cell_format})   

        wb = load_workbook(Byte_excel)
        ws = wb["Состояние_МП"]
        graph = Image(self.BytesImg)
        graph.height = 75*20
        graph.width = 75*8
        cell = "C" + str(self.excel_table.index.levels[0].size * 3 + 4)
        ws.add_image(graph,cell)
        Byte_excel.seek(0)
        wb.save(Byte_excel)
        with open(self.output_path/"Отчет_об_МП.xlsx", "wb") as file:
            file.write(Byte_excel.getbuffer())        

if __name__ == "__main__":
    V = Calculation()
    V.path_base = Path(r".\LC\Маршрутные листы.csv")
    V.order_list =  ["Г2109","32110"]
    V.path_out = Path(r'.\LC')
    V.count_months = 12
    V.start_work_table()
    
        