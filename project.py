import sys
import xlrd
import sqlite3
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QListWidgetItem, QMainWindow
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QInputDialog, QTableWidgetItem
import matplotlib.pyplot as plt


def clear_db():
    """Удаляю все таблицы из предыдущей базы данных"""
    con = sqlite3.connect('timepad.db')
    cur = con.cursor()
    table_names = [''.join(i)
                   for i in list(cur.execute("select name"
                                             " from sqlite_master"
                                             " where type='table'"))]
    for i in range(len(table_names)):
        cur.execute(f"drop table '{table_names[i]}'")
        con.commit()


class MainProjectsWindow(QMainWindow):
    """Главное окно"""

    def __init__(self):
        super().__init__()
        uic.loadUi('results_qt.ui', self)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        # self.window_button.clicked.connect(self.open_dialog_sorts)
        self.table_b.clicked.connect(self.open_table)
        self.stat_btn.clicked.connect(self.open_stat)
        self.window_button.clicked.connect(self.open_graph)

    def open_graph(self):
        self.gr = Infographics()
        self.gr.show()
        self.hide()

    def open_stat(self):
        self.dialog = StatsOptions()
        self.dialog.show()
        self.hide()

    def open_table(self):
        """Ввод excel таблицы"""
        table = QFileDialog.getOpenFileName(self, 'Выбрать картинку', '',
                                            "Таблица(*.xlsx)")[0]
        """Делаю через try, во избежание ошибки, если я ничего не ввел"""
        try:
            main_data = []
            data = []
            file = xlrd.open_workbook(table)  # открываю таблицу
            sheet = file.sheet_by_index(0)  # беру первый лист
            """В переменной vals храню всю информацию о таблице"""
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            type_of_ticket = vals[0].index('Тип билета')
            table_name = vals[1][0] + '. ' + vals[1][3][:-3]
            index_surname = vals[0].index('Фамилия')
            index_name = vals[0].index('Имя')
            index_email = vals[0].index('E-mail')
            index_summa = vals[0].index('Сумма')
            index_status = vals[0].index('Статус')
            index_source = vals[0].index('Источник трафика')
            index_date_of_buy = vals[0].index('Оплатил')
            for i in vals[1:]:
                arr = [vals.index(i)]
                for j in range(len(i)):
                    if j == index_surname:
                        arr.append(i[j])
                    if j == index_name:
                        arr.append(i[j])
                    if j == index_email:
                        arr.append(i[j])
                    if j == index_summa:
                        arr.append(i[j])
                    if j == index_status:
                        arr.append(i[j])
                    if j == index_source:
                        arr.append(i[j])
                    if j == index_date_of_buy:
                        arr.append(i[j])
                    if j == type_of_ticket:
                        arr.append(i[j])
                data.append(arr)
            for x in data:
                main_data.append([x[0], (x[6] + ' ' + x[7]).strip(), x[5], x[2], x[3], x[4], x[-1], x[1]])
            con = sqlite3.connect('timepad.db')
            cur = con.cursor()
            table_names = [''.join(i) for i in
                           list(cur.execute("select name from "
                                            "sqlite_master where type='table'"))]
            if table_name not in table_names:
                cur.execute(f"CREATE TABLE '{table_name}' (ID INTEGER, Full_name TEXT, "
                            f"E_mail TEXT, Sum_of_payment TEXT, Type_of_ticket TEXT, "
                            f"Status TEXT, Source TEXT, Date_of_payment TEXT, PRIMARY KEY (id));")
                for i in main_data:
                    cur.execute(f"insert into '{table_name}' "
                                f"(ID, Full_name, E_mail, "
                                f"Sum_of_payment, Type_of_ticket, Status, Source, Date_of_payment) values "
                                f"({i[0]}, '{i[1]}',"
                                f" '{i[2]}', '{i[3]}', '{i[4]}', '{i[5]}', '{i[6]}', '{i[7]}')")
                con.commit()
                self.success_dialog_add = QMessageBox.information(self, 'Успех',
                                                                  'Данные добавлены',
                                                                  buttons=QMessageBox.Ok)
            else:
                self.success_dialog_add = QMessageBox.critical(self, 'Ошибка',
                                                               'Данные уже существуют',
                                                               buttons=QMessageBox.Ok)
        except:
            pass


class Infographics(QMainWindow):
    """Окно с построением инфографики"""

    def __init__(self):
        super().__init__()
        self.con = sqlite3.connect('timepad.db')
        self.cur = self.con.cursor()
        uic.loadUi('graphics.ui', self)
        self.table_names = [i[0] for i in
                            list(self.cur.execute("select name from "
                                                  "sqlite_master where type='table'"))]
        for i in self.table_names:
            self.events.addItem(i)
        self.btn_exit.clicked.connect(self.exit)
        self.btn_show.clicked.connect(self.draw)

    def exit(self):
        """Переход в главное окно"""
        self.main_window = MainProjectsWindow()
        self.main_window.show()
        self.hide()

    def draw(self):
        """Анализ данных для создания инфографики"""
        #        try:
        self.data = []
        self.action = self.type.currentText()
        if self.action:
            if self.action == 'Количество участников':
                self.viewers = []
                self.unique = []
                if self.events.currentText() == 'По всем событиям':
                    table_names = self.table_names
                    for i in table_names:
                        self.data = self.cur.execute(f"select * from '{i}'").fetchall()
                        arr = []
                        for j in self.data:
                            if j[5] == 'оплачено':
                                arr.append(list(j)[2])
                        self.unique.append(len(set(arr)))
                        self.viewers.append(len(arr))
                else:
                    table_names = [self.events.currentText()]
                    self.data = self.cur.execute(f"select * from '{table_names[0]}'").fetchall()
                    arr = []
                    for j in self.data:
                        if j[5] == 'оплачено':
                            arr.append(list(j)[2])
                    self.unique.append(len(set(arr)))
                    self.viewers.append(len(arr))
                self.draw_histogram(table_names, [self.viewers, self.unique], name='Количество участников')
        else:
            self.success_dialog_add = QMessageBox.critical(self, 'Ошибка',
                                                           'Выберите тип диаграммы',
                                                           buttons=QMessageBox.Ok)

    #        except:
    #            self.success_dialog_add = QMessageBox.critical(self, 'Ошибка',
    #                                                           'Данные не выбраны',
    #                                                           buttons=QMessageBox.Ok)

    def autolabel(self, rects, labels=None, height_factor=1.01):
        """Функция, указывающие значения аргументов на гисторгамме"""
        for i, rect in enumerate(rects):
            height = rect.get_height()
            y = rect.get_y()
            if labels is not None:
                try:
                    label = labels[i][0]
                    print(label)
                except (TypeError, KeyError):
                    label = ' '
            else:
                label = '%d' % int(height)
            plt.text(rect.get_x() + rect.get_width() / 2., height_factor * (y + height),
                     '{}'.format(label),
                     ha='center', va='bottom')

    def draw_histogram(self, ox, oy, layering=False, width=0.4, name='Статистика'):
        """Создание гистограммы с помощью библиотеки matplotlib"""
        fig, axes = plt.subplots(nrows=1, ncols=1)
        ps = []
        if not layering:
            for n in oy:
                p = plt.bar(range(len(oy[0])), n, width)
                ps.append(p[0])
            plt.subplots_adjust(right=0.6)
            plt.xticks(range(len(ox)), [i[:10] + '...' for i in ox])
            self.autolabel(axes.patches)
            plt.legend(ps, ('Зрители', 'Уникальные'), prop={'size': 10}, loc="center left",
                       bbox_to_anchor=(1, 0, 0.5, 1))
            plt.show()

    def closeEvent(self, event):
        self.exit()


class StatsOptions(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('stats_options.ui', self)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        self.btn_back.clicked.connect(self.exit)
        self.btn_stats.clicked.connect(self.main_stat)

    def exit(self):
        """Открывает главное окно"""
        self.main_window = MainProjectsWindow()
        self.main_window.show()
        self.hide()

    def main_stat(self):
        data = []
        dictionary_sums = {}
        dictionary_of_visits = {'всего покупок': 0, 'уникальных покупок': 0,
                                'отказов': 0, 'просрочено': 0, 'бесплатно': 0}
        dictionary_types = {}
        main_sum = 0
        con = sqlite3.connect('timepad.db')
        cur = con.cursor()
        table_names = [''.join(i) for i in
                       list(cur.execute("select name from "
                                        "sqlite_master where type='table'"))]
        for i in table_names:
            data.append(list(cur.execute(f"select * from '{i}'")))
        dates = []
        for i in data:
            for j in list(i):
                if list(j)[-1]:
                    dates.append(list(j)[-1])
        self.period.setText(self.period.text() + ' с ' + str(dates[0]).split()[0] + ' по ' + str(dates[-1]).split()[0])
        for i in data:
            mip_sum = 0
            for j in i:
                if list(j)[3]:
                    mip_sum += int(list(j)[3][:-2])
            main_sum += mip_sum
            dictionary_sums[table_names[data.index(i)]] = mip_sum
        info_for_the_first_table = sorted(list(map(list, list(dictionary_sums.items()))), key=lambda x: x[1])[::-1]
        self.income_table.setColumnCount(2)
        self.income_table.setRowCount(len(info_for_the_first_table) + 1)
        self.income_table.setHorizontalHeaderItem(0, QTableWidgetItem('Событие'))
        self.income_table.setHorizontalHeaderItem(1, QTableWidgetItem('Доход'))
        for i in range(len(info_for_the_first_table)):
            self.income_table.setItem(i, 0, QTableWidgetItem(str(info_for_the_first_table[i][0])))
            self.income_table.setItem(i, 1, QTableWidgetItem(str(info_for_the_first_table[i][1])))
        self.income_table.setItem(len(info_for_the_first_table), 0, QTableWidgetItem('Итого:'))
        self.income_table.setItem(len(info_for_the_first_table), 1, QTableWidgetItem(str(main_sum)))
        for i in data:
            for j in list(i):
                if list(j)[5] == 'оплачено':
                    dictionary_of_visits['всего покупок'] += 1
                else:
                    dictionary_of_visits[list(j)[5]] += 1
        for i in data:
            for j in list(i):
                if list(j)[4] not in dictionary_types.keys():
                    dictionary_types[list(j)[4]] = 1
                elif list(j)[4] in dictionary_types.keys():
                    dictionary_types[list(j)[4]] += 1
        arr = []
        for i in data:
            for j in list(i):
                if list(j)[5] == 'оплачено':
                    arr.append(list(j)[2])
        dictionary_of_visits['уникальных покупок'] += len(set(arr))

        info_for_the_second_table = sorted(list(map(list, list(dictionary_of_visits.items()))),
                                           key=lambda x: x[1])[::-1]
        self.visits_table.setColumnCount(2)
        self.visits_table.setRowCount(len(info_for_the_second_table))
        self.visits_table.setHorizontalHeaderItem(0, QTableWidgetItem('Посещения'))
        self.visits_table.setHorizontalHeaderItem(1, QTableWidgetItem('Кол-во посещений'))
        for i in range(len(info_for_the_second_table)):
            self.visits_table.setItem(i, 0, QTableWidgetItem(str(info_for_the_second_table[i][0])))
            self.visits_table.setItem(i, 1, QTableWidgetItem(str(info_for_the_second_table[i][1])))
        info_for_the_third_table = sorted(list(map(list, list(dictionary_types.items()))), key=lambda x: x[1])[::-1]
        self.typies_table.setColumnCount(2)
        self.typies_table.setRowCount(len(info_for_the_third_table))
        self.typies_table.setHorizontalHeaderItem(0, QTableWidgetItem('Митап'))
        self.typies_table.setHorizontalHeaderItem(1, QTableWidgetItem('Кол-во посещений'))
        for i in range(len(info_for_the_third_table)):
            self.typies_table.setItem(i, 0, QTableWidgetItem(str(info_for_the_third_table[i][0])))
            self.typies_table.setItem(i, 1, QTableWidgetItem(str(info_for_the_third_table[i][1])))
        data_for_reg = []
        for i in data:
            dictionary_customers = {}
            for j in list(i):
                num = list(j)[3][:-2]
                if list(j)[1] not in dictionary_customers:
                    if num:
                        dictionary_customers[list(j)[1]] = [1, int(num)]
                    else:
                        dictionary_customers[list(j)[1]] = [1, 0]
                else:
                    if num:
                        dictionary_customers[list(j)[1]] = [dictionary_customers[list(j)[1]][0] + 1,
                                                            dictionary_customers[list(j)[1]][1] + int(num)]
                    else:
                        dictionary_customers[list(j)[1]] = [dictionary_customers[list(j)[1]][0] + 1,
                                                            dictionary_customers[list(j)[1]][1] + 0]
            data_for_reg.append(list(dictionary_customers.items()))
        for i in data_for_reg:
            for j in list(i):
                if j[1][0] >= 2:
                    j[1][0] = 1
        real_reg = {}
        for i in data_for_reg:
            for j in i:
                name = list(j)[0]
                count_events = list(j)[1][0]
                sums = list(j)[1][1]
                if name not in real_reg:
                    if sums:
                        real_reg[name] = [count_events, sums]
                    else:
                        real_reg[name] = [count_events, 0]
                else:
                    if sums:
                        real_reg[name] = [real_reg[name][0] + 1, real_reg[name][1] + sums]
                    else:
                        real_reg[name] = [real_reg[name][0] + 1, real_reg[name][1] + 0]
        info_for_the_forth_table = sorted(list(map(list, list(real_reg.items()))),
                                          key=lambda x: (x[1][0], x[0]))[::-1]
        add_info = list(filter(lambda x: x[1][0] == 1, info_for_the_forth_table))
        sum_of_rest_mon = 0
        for i in add_info:
            sum_of_rest_mon += i[1][1]
        info_for_the_forth_table = list(filter(lambda x: x[1][0] >= 2, info_for_the_forth_table))
        self.reg_table.setColumnCount(3)
        self.reg_table.setRowCount(len(info_for_the_forth_table) + 1)
        self.reg_table.setHorizontalHeaderItem(0, QTableWidgetItem('Посетитель'))
        self.reg_table.setHorizontalHeaderItem(1, QTableWidgetItem('Потрачено денег'))
        self.reg_table.setHorizontalHeaderItem(2, QTableWidgetItem('Кол-во посещений'))
        for i in range(len(info_for_the_forth_table)):
            self.reg_table.setItem(i, 0, QTableWidgetItem(str(info_for_the_forth_table[i][0])))
            self.reg_table.setItem(i, 1, QTableWidgetItem(str(info_for_the_forth_table[i][1][1])))
            self.reg_table.setItem(i, 2, QTableWidgetItem(str(info_for_the_forth_table[i][1][0])))
        self.reg_table.setItem(len(info_for_the_forth_table), 0, QTableWidgetItem('Остальные'))
        self.reg_table.setItem(len(info_for_the_forth_table), 1, QTableWidgetItem(str(sum_of_rest_mon)))
        self.reg_table.setItem(len(info_for_the_forth_table), 2,
                               QTableWidgetItem(str(dictionary_of_visits['всего покупок']
                                                    + dictionary_of_visits['бесплатно'])))


def my_excepthook(type, value, tback):
    a = QMessageBox.critical(windows, "Упс... Ошибка", str(value) +
                             "\nЗаскринь и отправь мне\nhttps://vk.com/mr_sadness",
                             QMessageBox.Cancel)

    sys.__excepthook__(type, value, tback)


sys.excepthook = my_excepthook

clear_db()
app = QApplication(sys.argv)
windows = MainProjectsWindow()
windows.show()
sys.exit(app.exec_())
