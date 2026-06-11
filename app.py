import sys
import numpy as np
import pandas as pd
import json

from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QColor, QBrush

from ExcelReplacer import *
from PdfReader import *


class Main(QDialog):
    def __init__(self):
        super().__init__()
        self.setFixedSize(700, 390)
        self.setWindowTitle('Adapt')
        self.setWindowIcon(QtGui.QIcon('logo.ico'))
        self.setWindowFlags(
            Qt.Window | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
        self.setStyleSheet(
            '''
                QDialog { 
                    background-image: url(pdfxls.jpg);
                    background-size: 100%;
                    background-position: center;
                    background-color: (228,228,228);
                    font: bold;
                }
                QLabel, QComboBox, QPushButton {
                    font-size: 12pt;
                }

                QPushButton {
                    background-color: rgb(248,248,248);
                    border-radius: 10px;
                    border: 2px solid black;
                    font: bold;

                }

                QPushButton::hover {
                    background-color: rgb(255,255,255);
                    color: rgb(30, 100, 30); 
                    border: 2px solid rgb(30, 100, 30);
                }                
            '''
        )
        app.setStyleSheet(
            '''
                QDialog{
                    background-color: (228,228,228);
                }
                QLabel, QComboBox, QPushButton {
                    font-size: 12pt;
                }
            '''
        )

        # Кнопка списком наименований из базы данных
        self.namesButton = QPushButton('Список эталонных названий', self)
        self.namesButton.setGeometry(QtCore.QRect(20, 10, 320, 50))
        self.namesButton.setObjectName("namesButton")
        self.namesButton.clicked.connect(self.openNamesDialog)

        # Кнопка загрузки базы данных
        self.DBButton = QPushButton('Загрузить базу данных', self)
        self.DBButton.setGeometry(QtCore.QRect(360, 10, 320, 50))
        self.DBButton.setObjectName("baseButton")
        self.DBButton.clicked.connect(self.loadDB)


        # Кнопка конвертации pdf в excel
        self.fileGetter = QPushButton('Перевод PDF в xls', self)
        self.fileGetter.setGeometry(QtCore.QRect(20, 330, 320, 50))
        self.fileGetter.setObjectName("load_button1")
        self.fileGetter.clicked.connect(self.openPdfReader)
        # Кнопка обработки файла
        self.replacerButton = QPushButton('Обработка', self)
        self.replacerButton.setGeometry(QtCore.QRect(360, 330, 320, 50))
        self.replacerButton.setObjectName("load_button2")
        self.replacerButton.clicked.connect(self.openMarker)

    def openNamesDialog(self):
        dialog = NamesDialog()
        dialog.exec()
        dialog.show()

    def loadDB(self):
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбор файла', '',
                                                  'Excel Files (*.xlsx *.xls)')
        if not filename:
            return  # Пользователь не выбрал файл

        conf = ConfirmDB()
        result = conf.showDialog(filename)
        if result == QMessageBox.Cancel:
            return

        try:
            dfs = pd.read_excel(filename, sheet_name=None)
            dfs = {key: df for key, df in dfs.items() if not df.empty}
            dfs = strip_strings_in_dataframes(dfs)
            with open("standards.json", "r", encoding="utf-8") as read_file:
                data = json.load(read_file)

            for df_name, df in dfs.items():
                for index, row in df.iterrows():
                    if pd.notna(row.iloc[1]):
                        standard_article = row.iloc[0]
                        standard_name = row.iloc[1]
                        name = row.iloc[2]
                        measure = row.iloc[3]
                        article = ""
                        if standard_name not in data:
                            if pd.isna(standard_article):
                                standard_article = ""
                            if pd.isna(name) and name != standard_name:
                                name = ""
                            if pd.isna(measure):
                                measure = ""

                            save_new_object(data, standard_name, standard_article, measure, name, article)
                        else:
                            if pd.notna(standard_article) and data[standard_name]['standard_article'] != standard_article:
                                save_non_standard_article(data, data[standard_name]['standard_article'], standard_name)
                                data[standard_name]['standard_article'] = standard_article

                            if pd.notna(name):
                                save_non_standard_name(data, name, standard_name)

                            if pd.notna(measure) and data[standard_name]['measure'] != measure:
                                data[standard_name]['measure'] = measure

                    elif pd.notna(row.iloc[2]) and row.iloc[2] not in data:
                        standard_article = row.iloc[0]
                        standard_name = row.iloc[2]
                        measure = row.iloc[3]
                        name = ""
                        article = ""

                        if standard_name not in data:
                            if pd.isna(standard_article):
                                standard_article = ""
                            if pd.isna(measure):
                                measure = ""
                            save_new_object(data, standard_name, standard_article, measure, name, article)

                        else:
                            if pd.notna(standard_article) and data[standard_name]['standard_article'] != standard_article:
                                save_non_standard_article(data, data[standard_name]['standard_article'], standard_name)
                                data[standard_name]['standard_article'] = standard_article

                            if pd.notna(measure) and self.data[standard_name]['measure'] != measure:
                                data[standard_name]['measure'] = measure

            with open("standards.json", "w", encoding="utf-8") as write_file:
                json.dump(data, write_file, ensure_ascii=False, indent=4)

            end_msg = QMessageBox()
            end_msg.setWindowIcon(QtGui.QIcon('logo.ico'))
            end_msg.setText(f"Данные из файла <b>{filename}</b> добавлены в базу данных приложения")
            end_msg.setWindowTitle("Готово")
            end_msg.exec_()

        except Exception as error:
            msg = QMessageBox()
            msg.setWindowIcon(QtGui.QIcon('logo.ico'))
            msg.setIcon(QMessageBox.Critical)
            msg.setText(f"Не удалось загрузить файл <b>{filename}</b>")
            msg.setInformativeText(f"{error}")
            msg.setWindowTitle("Ошибка")
            msg.exec_()

    def openMarker(self):
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбор файла', '',
                                                  'Excel Files (*.xlsx *.xls)')
        if filename:  # Пользователь выбрал файл
            try:
                dfs = pd.read_excel(filename, sheet_name=None)
                dfs = {key: df for key, df in dfs.items() if not df.empty}
                dfs = strip_strings_in_dataframes(dfs)
                check_column_names(dfs)

            except Exception as error:
                msg = QMessageBox()
                msg.setWindowIcon(QtGui.QIcon('logo.ico'))
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Ошибка данных")
                msg.setInformativeText(f"{error}")
                msg.setWindowTitle("Ошибка")
                msg.exec_()

            else:
                dialog = Marker(filename)  # Передача filename в конструктор класса Marker
                dialog.show()
                self.close()
                dialog.exec()

    def openPdfReader(self):
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбор файла',
                                                  '',
                                                  'PDF Files (*.pdf)')
        try:
            if filename:  # Пользователь выбрал файл
                tables = pdf_reader(filename)

                output, _ = QFileDialog.getSaveFileName(self, "Сохранение файла",
                                                        '',
                                                        "Excel Files (*.xlsx)")
                if output:
                    list_to_excel(tables, output)

        except Exception as error:
            msg = QMessageBox()
            msg.setWindowIcon(QtGui.QIcon('logo.ico'))
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Ошибка")
            msg.setText(f"Не удалось загрузить файл <b>{filename}</b>")
            msg.setInformativeText(f"{error}")
            msg.exec_()


class ConfirmDB(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowIcon(QtGui.QIcon('logo.ico'))

    def showDialog(self, filename):
        msg = QMessageBox(self)
        msg.setWindowTitle('Подтверждение')
        msg.setIcon(QMessageBox.Question)
        msg.setText(f"Добавить данные из файла <b>{filename}</b> в базу данных?")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

        msg.button(QMessageBox.Ok).setText("да")
        msg.button(QMessageBox.Cancel).setText("нет")

        retval = msg.exec_()
        return retval


class NamesDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setFixedSize(1200, 800)
        self.setWindowTitle('Имена')
        self.setWindowIcon(QtGui.QIcon('logo.ico'))

        self.setWindowFlags(
            Qt.Window | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)

        # names list #
        self.listLayout = QVBoxLayout()
        with open('standards.json', encoding='utf-8') as file:
            data = json.load(file)
            self.tableWidget = QTableWidget(len(data), 1)
            self.tableWidget.selectionModel().selectionChanged.connect(self.on_selectionChanged)
            self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget.setSelectionMode(QTableWidget.SingleSelection)
            self.tableWidget.setHorizontalHeaderLabels(["Наименования"])
            self.listLayout.addWidget(self.tableWidget)
            self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

            i = 0
            for name in data:
                self.tableWidget.setItem(i, 0, QTableWidgetItem(name))

                i += 1

        w = QWidget(self)
        w.setGeometry(QtCore.QRect(20, 10, 600, 720))
        w.setLayout(self.listLayout)
        w.show()

        # item view #

        self.viewListLayout = QVBoxLayout()
        self.labelName = QLabel("Наименование", self)
        self.labelName.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        self.labelAtr = QLabel("Артикул", self)
        self.labelAtr.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        self.labelMeasure = QLabel("Единица измерения", self)
        self.labelMeasure.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        self.nameEdit = QLineEdit(self)
        self.nameEdit.setFixedWidth(500)
        self.atrEdit = QLineEdit(self)
        self.atrEdit.setFixedWidth(500)
        self.measureEdit = QLineEdit(self)
        self.measureEdit.setFixedWidth(500)
        self.viewListLayout.addWidget(self.labelName)
        self.viewListLayout.addWidget(self.nameEdit)
        self.viewListLayout.addWidget(self.labelAtr)
        self.viewListLayout.addWidget(self.atrEdit)
        self.viewListLayout.addWidget(self.labelMeasure)
        self.viewListLayout.addWidget(self.measureEdit)

        w2 = QWidget(self)
        w2.setGeometry(QtCore.QRect(640, 10, 600, 600))
        w2.setLayout(self.viewListLayout)
        w2.show()

        # Ввод
        self.NameReader = QLineEdit(self)
        self.NameReader.setGeometry(QtCore.QRect(9, 750, 1011, 30))
        self.NameReader.setObjectName("NameReader")
        # Удалить кнопка
        self.NameDeleter = QPushButton(self)
        self.NameDeleter.setGeometry(QtCore.QRect(1040, 630, 140, 30))
        self.NameDeleter.setObjectName("NameDeleter")
        self.NameDeleter.setText('удалить')
        self.NameDeleter.clicked.connect(self.deleteName)
        # Изменить кнопка
        self.NameEditor = QPushButton(self)
        self.NameEditor.setGeometry(QtCore.QRect(900, 630, 140, 30))
        self.NameEditor.setObjectName("NameEditor")
        self.NameEditor.setText('изменить')
        self.NameEditor.clicked.connect(self.editName)
        # Кнопка добавить
        self.NameAdder = QPushButton(self)
        self.NameAdder.setGeometry(QtCore.QRect(1040, 750, 140, 30))
        self.NameAdder.setObjectName("NameAdder")
        self.NameAdder.setText('добавить')
        self.NameAdder.clicked.connect(self.addName)

    def on_selectionChanged(self, selected):
        for ix in selected.indexes():
            row = ix.row()
            with open('standards.json', encoding='utf-8') as file:
                data = json.load(file)
                self.nameEdit.setText(self.tableWidget.item(row, 0).text())
                self.atrEdit.setText(data[self.tableWidget.item(row, 0).text()]['standard_article'])
                self.measureEdit.setText(data[self.tableWidget.item(row, 0).text()]['measure'])

    def deleteName(self):
        row = self.tableWidget.currentRow()
        if row == -1:
            return
        msg = QMessageBox.question(
            self,
            "Внимание подтвердите удаление строки!",
            "Вы действительно хотите удалить "
            f"строку <b style='color: red;'>{row + 1}</b> ?",
            QMessageBox.Ok | QMessageBox.Cancel
        )
        if msg == QMessageBox.Cancel:
            return

        self.tableWidget.removeRow(row)
        # код удаляющий из файла
        with open('standards.json', encoding='utf-8') as file:
            data = json.load(file)
        self.nameEdit.setText('')
        self.atrEdit.setText('')
        self.measureEdit.setText('')
        try:
            j = 0
            for name in data:
                if j == row:
                    del data[name]

                    with open('standards.json', 'w', encoding="utf-8") as file:
                        json.dump(data, file, ensure_ascii=False, indent=4)

                    return
                j += 1
        except:
            pass

    def addName(self):
        name = self.NameReader.text()
        if name == "":
            return
        rowPosition = self.tableWidget.rowCount()
        self.tableWidget.insertRow(rowPosition)
        self.tableWidget.setItem(rowPosition, 0, QTableWidgetItem(name))

        with open('standards.json', encoding='utf-8') as file:
            data = json.load(file)
            data.update({name: {"standard_article": "", "measure": "", "names": ""}})
        with open('standards.json', 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False, indent=4)

    def editName(self):
        row = self.tableWidget.currentRow()
        if row == -1:
            return
        msg = QMessageBox.question(
            self,
            "Внимание подтвердите перезапись строки!",
            "Вы действительно хотите перезаписать "
            f"строку <b style='color: red;'>{row + 1}</b> ?",
            QMessageBox.Ok | QMessageBox.Cancel
        )
        if msg == QMessageBox.Cancel:
            return

        self.tableWidget.removeRow(row)
        with open('standards.json', encoding='utf-8') as file:
            data = json.load(file)
        itemname = self.nameEdit.text()
        art = self.atrEdit.text()
        measure = self.measureEdit.text()

        # список предыдущих использований
        # names = data[self.tableWidget.item(row, 0).text()]['names']
        # print(names)

        try:
            j = 0
            for name in data:
                if j == row:
                    del data[name]
                    break
                j += 1
        except:
            pass

        data.update({itemname: {"standard_article": art, "measure": measure, "names": "", "articles": ""}})
        with open('standards.json', 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False, indent=4)

        self.tableWidget.insertRow(row)
        self.tableWidget.setItem(row, 0, QTableWidgetItem(itemname))


class Marker(QDialog):
    def __init__(self, filename):  # Добавление filename в качестве аргумента конструктора
        super().__init__()
        # Лэйаут
        self.layout = None

        # Выпадающие списки
        self.combo_article = None
        self.combo_name = None
        self.combo_quantity = None
        self.combo_measure = None
        self.combo_add1 = None
        self.combo_add2 = None
        self.combo_add3 = None

        #  Кнопки
        self.okButton = None
        self.backButton = None

        self.setWindowFlags(
            Qt.Window | Qt.WindowMaximizeButtonHint | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)

        #  Данные
        self.dfs = None
        self.article_index = 3
        self.name_index = 1
        self.quantity_index = 6
        self.measure_index = 5
        self.add1_index = 0
        self.add2_index = 0
        self.add3_index = 0
        self.filename = filename
        self.initUI()

    def initUI(self):
        self.resize(1200, 800)
        self.setWindowTitle('Разметка')
        self.setWindowIcon(QtGui.QIcon('logo.ico'))
        self.setStyleSheet("QPushButton{font-size: 12pt;}")
        self.layout = QVBoxLayout(self)

        self.dfs = pd.read_excel(self.filename, sheet_name=None)
        self.dfs = {key: df for key, df in self.dfs.items() if not df.empty}
        self.dfs = strip_strings_in_dataframes(self.dfs)
        first_dataframe = list(self.dfs.values())[0]
        common_column_names = [""]
        common_column_names.extend(list(first_dataframe.columns))

        ######
        # Добавление лэйблов и выпадающих списков и подключение сигнала (7 раз одно и то же)

        row_layout = QHBoxLayout()
        # Добавление QLabel
        label = QLabel("Артикул", self)
        label.setAlignment(Qt.AlignCenter)
        row_layout.addWidget(label)
        # Добавление QComboBox
        self.combo_article = QComboBox(self)
        self.combo_article.addItems(common_column_names)
        self.combo_article.setCurrentIndex(self.article_index)
        row_layout.addWidget(self.combo_article)
        # Добавление горизонтального компоновщика в вертикальный компоновщик
        self.layout.addLayout(row_layout)

        # Подключение сигнала currentIndexChanged к слоту on_combo_index_changed
        self.combo_article.currentIndexChanged.connect(self.on_combo_article_changed)

        row_layout = QHBoxLayout()
        # Добавление QLabel
        label = QLabel("Товары (работы, услуги)", self)
        label.setAlignment(Qt.AlignCenter)
        row_layout.addWidget(label)
        # Добавление QComboBox
        self.combo_name = QComboBox(self)
        self.combo_name.addItems(common_column_names)
        self.combo_name.setCurrentIndex(self.name_index)
        row_layout.addWidget(self.combo_name)
        # Добавление горизонтального компоновщика в вертикальный компоновщик
        self.layout.addLayout(row_layout)

        # Подключение сигнала currentIndexChanged к слоту on_combo_index_changed
        self.combo_name.currentIndexChanged.connect(self.on_combo_name_changed)

        row_layout = QHBoxLayout()
        # Добавление QLabel
        label = QLabel("Кол-во", self)
        label.setAlignment(Qt.AlignCenter)
        row_layout.addWidget(label)
        # Добавление QComboBox
        self.combo_quantity = QComboBox(self)
        self.combo_quantity.addItems(common_column_names)
        self.combo_quantity.setCurrentIndex(self.quantity_index)
        row_layout.addWidget(self.combo_quantity)
        # Добавление горизонтального компоновщика в вертикальный компоновщик
        self.layout.addLayout(row_layout)

        # Подключение сигнала currentIndexChanged к слоту on_combo_index_changed
        self.combo_quantity.currentIndexChanged.connect(self.on_combo_quantity_changed)

        row_layout = QHBoxLayout()
        # Добавление QLabel
        label = QLabel("Ед.", self)
        label.setAlignment(Qt.AlignCenter)
        row_layout.addWidget(label)
        # Добавление QComboBox
        self.combo_measure = QComboBox(self)
        self.combo_measure.addItems(common_column_names)
        self.combo_measure.setCurrentIndex(self.measure_index)
        row_layout.addWidget(self.combo_measure)
        # Добавление горизонтального компоновщика в вертикальный компоновщик
        self.layout.addLayout(row_layout)

        # Подключение сигнала currentIndexChanged к слоту on_combo_index_changed
        self.combo_measure.currentIndexChanged.connect(self.on_combo_measure_changed)

        row_layout = QHBoxLayout()
        # Добавление QLabel
        label = QLabel("Доп. информация\n(опционально)", self)
        label.setAlignment(Qt.AlignCenter)
        row_layout.addWidget(label)

        # Добавление QComboBox
        self.combo_add1 = QComboBox(self)
        self.combo_add1.addItems(common_column_names)
        self.combo_add1.setCurrentIndex(self.add1_index)
        row_layout.addWidget(self.combo_add1)
        # Подключение сигнала currentIndexChanged к слоту on_combo_index_changed
        self.combo_add1.currentIndexChanged.connect(self.on_combo_add1_changed)

        self.combo_add2 = QComboBox(self)
        self.combo_add2.addItems(common_column_names)
        self.combo_add2.setCurrentIndex(self.add2_index)
        row_layout.addWidget(self.combo_add2)
        # Подключение сигнала currentIndexChanged к слоту on_combo_index_changed
        self.combo_add2.currentIndexChanged.connect(self.on_combo_add2_changed)

        self.combo_add3 = QComboBox(self)
        self.combo_add3.addItems(common_column_names)
        self.combo_add3.setCurrentIndex(self.add1_index)
        row_layout.addWidget(self.combo_add3)
        # Подключение сигнала currentIndexChanged к слоту on_combo_index_changed
        self.combo_add3.currentIndexChanged.connect(self.on_combo_add3_changed)

        # Добавление горизонтального компоновщика в вертикальный компоновщик
        self.layout.addLayout(row_layout)

        ######

        row_layout = QHBoxLayout()

        # Кнопка ок
        self.okButton = QPushButton('ок', self)
        self.okButton.clicked.connect(self.okAction)
        row_layout.addWidget(self.okButton)

        # Кнопка отмена
        self.backButton = QPushButton('назад', self)
        self.backButton.clicked.connect(self.backAction)
        row_layout.addWidget(self.backButton)

        self.layout.addLayout(row_layout)

        # Расположение вертикального компоновщика
        self.setLayout(self.layout)

    def okAction(self):
        main_columns = [self.article_index, self.name_index, self.quantity_index, self.measure_index]
        add_info_columns = [self.add1_index, self.add2_index, self.add3_index]
        add_info_columns = list(filter(lambda x: x != 0, add_info_columns))  # удаляем нули из списка
        if self.article_index != 0 and self.name_index != 0 and self.quantity_index != 0 and self.measure_index != 0:
            if len(set(add_info_columns + main_columns)) == len(add_info_columns + main_columns):
                for key, df in self.dfs.items():
                    if not pd.api.types.is_numeric_dtype(df.iloc[:, self.quantity_index - 1]):
                        msg = QMessageBox()
                        msg.setWindowIcon(QtGui.QIcon('logo.ico'))
                        msg.setWindowTitle("Предупреждение")
                        msg.setIcon(QMessageBox.Warning)
                        msg.setText(f"В поле количества присутствуют не числовые данные")
                        msg.exec_()
                        return
                dialog = Replacer(self.dfs, main_columns, add_info_columns, self.filename)
                dialog.show()
                self.close()
                dialog.exec()

    def backAction(self):
        dialog = Main()
        dialog.show()
        self.close()
        dialog.exec()

    # Слоты для обработки изменения выбранного элемента в QComboBox'ах (7 раз одно и тоже опять)
    def on_combo_article_changed(self, index):
        self.article_index = index

    def on_combo_name_changed(self, index):
        self.name_index = index

    def on_combo_quantity_changed(self, index):
        self.quantity_index = index

    def on_combo_measure_changed(self, index):
        self.measure_index = index

    def on_combo_add1_changed(self, index):
        self.add1_index = index

    def on_combo_add2_changed(self, index):
        self.add2_index = index

    def on_combo_add3_changed(self, index):
        self.add3_index = index

class Replacer(QDialog):
    def __init__(self, dfs, main_columns, add_info_columns, filename):
        super().__init__()
        # GUI элементы
        ## Лэйауты
        self.layout = None
        self.listLayout = None
        self.suggsLayout = None
        self.buttons1Layout = None
        self.buttons2Layout = None

        ## Виджеты-обёртки
        self.table = None
        self.suggs = None
        self.buttons1 = None
        self.buttons2 = None

        ## Кнопки
        self.addRowButton = None
        self.deleteRowButton = None
        self.resetButton = None
        self.replaceButton = None
        self.backButton = None
        self.okButton = None

        ## Таблицы
        self.tableWidget = None
        self.currentItemTableWidget = None
        self.originalItemTableWidget = None
        self.sugTableWidget = None

        self.checkboxSaveReplacement = None
        self.checkboxMergeTabeles = None

        ## Текстовые поля
        self.orig_comment = None
        self.sug_comment = None

        # Данные и таблицы
        self.data = None

        self.df = None
        self.df_with_add_info = None

        self.columns_with_add_info = None

        self.replacements_dict = None
        self.warn_dict = None
        self.suggestions_list = None
        self.suggestions = None

        self.current_row = None

        # Прочие поля
        self.oldPos = None

        self.setWindowFlags(
            Qt.Window | Qt.WindowMaximizeButtonHint | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)

        self.filename = filename
        self.initUI(dfs, main_columns, add_info_columns)

    def initUI(self, dfs, main_columns, add_info_columns):
        try:
            loading_msg = QMessageBox()
            loading_msg.setWindowIcon(QtGui.QIcon('logo.ico'))
            loading_msg.setText(f"Пожалуйста подождите")
            loading_msg.setWindowTitle("Загрузка")
            loading_msg.show()
            self.setStyleSheet("QLabel, QPushButton, QCheckBox{font-size: 12pt;} QTextBrowser{font-size: 10pt;}")
            # self.showMaximized()
            self.setWindowTitle('Замена')
            self.setWindowIcon(QtGui.QIcon('logo.ico'))
            self.resize(1200, 800)

            self.layout = QHBoxLayout(self)

            for i, val in enumerate(main_columns):
                main_columns[i] = val - 1

            if add_info_columns:
                for i, val in enumerate(add_info_columns):
                    add_info_columns[i] = val - 1

            articles_column = main_columns[0]
            names_column = main_columns[1]
            quantity_column = main_columns[2]
            measure_column = main_columns[3]

            self.suggestions_list, self.warn_dict = get_suggestions(dfs, main_columns, add_info_columns)

            loading_msg.close()

            with open("standards.json", "r", encoding="utf-8") as read_file:
                self.data = json.load(read_file)

            self.replacements_dict = {}

            # Создание пустого списка для хранения DataFrame
            frames_list = []
            # Итерация по словарю DataFrame и добавление столбца 'source'
            for key, df in dfs.items():
                df_with_key = df.assign(source=key)
                frames_list.append(df_with_key)

            # Объедините их с помощью pd.concat()
            self.df = pd.concat(frames_list)

            self.df = self.df.fillna('')
            key_column = len(self.df.columns) - 1

            result_columns = [articles_column, names_column, quantity_column, measure_column, key_column]

            self.columns_with_add_info = [articles_column, names_column]
            for val in add_info_columns:
                self.columns_with_add_info.append(val)
            self.columns_with_add_info.append(quantity_column)
            self.columns_with_add_info.append(measure_column)
            self.df_with_add_info = self.df.iloc[:, self.columns_with_add_info]
            self.df = self.df.iloc[:, result_columns]
            new_columns = ['Артикул', 'Товары (работы, услуги) ', 'Кол-во', 'Ед.', 'Лист']
            self.df.columns = new_columns
            for index, row in self.df.iterrows():
                measure_value = row.iloc[3]
                new_value = measure_check(measure_value)
                if new_value is not None:
                    self.df.iloc[index, 3] = new_value

            self.df_table = self.df.copy()
            self.df_table.insert(1, 'Артикул (исходные)', self.df.iloc[:, 0])
            self.df_table.insert(3, 'Наименование (исходные)', self.df.iloc[:, 1])
            self.listLayout = QVBoxLayout()

            label = QLabel("Текущая таблица", self)
            label.setAlignment(Qt.AlignCenter)
            self.listLayout.addWidget(label)

            # Создаем QTableWidget
            self.tableWidget = QTableWidget()
            self.populate_table(self.tableWidget, self.df_table)
            self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self.tableWidget.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            self.tableWidget.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            self.tableWidget.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            self.tableWidget.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
            self.tableWidget.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)

            self.tableWidget.selectionModel().selectionChanged.connect(self.on_selectionChangedMain)
            self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget.setSelectionBehavior(QTableWidget.SelectRows)
            self.tableWidget.setSelectionMode(QTableWidget.SingleSelection)

            self.color_columns_grey([1, 3], self.tableWidget)
            self.set_tooltips(self.tableWidget)

            self.listLayout.addWidget(self.tableWidget)

            #self.buttons1Layout = QHBoxLayout()
            #self.addRowButton = QPushButton("добавить строку", self)
            ## self.addRowButton.clicked.connect(self.)
            #self.buttons1Layout.addWidget(self.addRowButton)

            #self.deleteRowButton = QPushButton("удалить строку", self)
            #self.deleteRowButton.clicked.connect(self.deleteAction)
            #self.buttons1Layout.addWidget(self.deleteRowButton)

            #self.buttons1 = QWidget(self)
            #self.buttons1.setLayout(self.buttons1Layout)
            #self.listLayout.addWidget(self.buttons1)

            self.backButton = QPushButton("назад", self)
            self.backButton.clicked.connect(self.backAction)
            self.listLayout.addWidget(self.backButton)

            self.table = QWidget(self)
            self.table.setLayout(self.listLayout)
            self.layout.addWidget(self.table)

            self.suggsLayout = QVBoxLayout()

            self.orig_comment = QTextBrowser()
            self.suggsLayout.addWidget(self.orig_comment)

            label = QLabel("Изначальные данные", self)
            label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            self.suggsLayout.addWidget(label)

            # Создаем QTableWidget
            self.originalItemTableWidget = QTableWidget()
            self.originalItemTableWidget.setColumnCount(len(self.columns_with_add_info))
            self.originalItemTableWidget.setRowCount(1)
            self.originalItemTableWidget.verticalHeader().setVisible(False)
            self.originalItemTableWidget.setHorizontalHeaderLabels(self.df_with_add_info.columns)

            header = self.originalItemTableWidget.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.Stretch)
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            for i in range(len(add_info_columns)):
                header.setSectionResizeMode(2 + i, QHeaderView.Stretch)
            header.setSectionResizeMode(2 + len(add_info_columns), QHeaderView.ResizeToContents)
            header.setSectionResizeMode(3 + len(add_info_columns), QHeaderView.ResizeToContents)
            header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)

            self.adjust_table_height(self.originalItemTableWidget)
            self.originalItemTableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.originalItemTableWidget.setSelectionMode(QTableWidget.NoSelection)

            self.suggsLayout.addWidget(self.originalItemTableWidget)

            label = QLabel("Текущая строка", self)
            label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            self.suggsLayout.addWidget(label)

            # Создаем QTableWidget
            self.currentItemTableWidget = QTableWidget()
            self.currentItemTableWidget.setColumnCount(4)
            self.currentItemTableWidget.setRowCount(1)
            self.currentItemTableWidget.verticalHeader().setVisible(False)
            self.currentItemTableWidget.setHorizontalHeaderLabels(self.df.columns[:4])

            self.currentItemTableWidget.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            self.currentItemTableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self.currentItemTableWidget.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
            self.currentItemTableWidget.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)

            self.currentItemTableWidget.itemChanged.connect(self.on_item_changed)

            self.adjust_table_height(self.currentItemTableWidget)
            self.suggsLayout.addWidget(self.currentItemTableWidget)

            self.buttons2Layout = QHBoxLayout()
            self.replaceButton = QPushButton("заменить", self)
            self.replaceButton.clicked.connect(self.replaceAction)
            self.buttons2Layout.addWidget(self.replaceButton)

            self.resetButton = QPushButton("сбросить", self)
            self.resetButton.clicked.connect(self.resetAction)
            self.buttons2Layout.addWidget(self.resetButton)

            self.buttons2 = QWidget(self)
            self.buttons2.setLayout(self.buttons2Layout)
            self.suggsLayout.addWidget(self.buttons2)

            self.checkboxSaveReplacement = QCheckBox('Сохранять замену в базу данных')
            self.checkboxSaveReplacement.setChecked(True)
            self.suggsLayout.addWidget(self.checkboxSaveReplacement)

            label = QLabel("Предложения", self)
            label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            self.suggsLayout.addWidget(label)

            # Создаем QTableWidget
            self.sugTableWidget = QTableWidget()
            self.sugTableWidget.setColumnCount(3)
            self.sugTableWidget.setRowCount(7)

            self.sugTableWidget.setHorizontalHeaderLabels(self.df.iloc[:, [0, 1, 3]])

            self.sugTableWidget.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            self.sugTableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self.sugTableWidget.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
            self.adjust_table_height(self.sugTableWidget)

            self.sugTableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.sugTableWidget.setSelectionBehavior(QTableWidget.SelectRows)
            self.sugTableWidget.setSelectionMode(QTableWidget.SingleSelection)
            self.sugTableWidget.setRowCount(0)
            self.sugTableWidget.selectionModel().selectionChanged.connect(self.on_selectionChangedSuggs)

            self.suggsLayout.addWidget(self.sugTableWidget)

            self.sug_comment = QTextBrowser()
            self.suggsLayout.addWidget(self.sug_comment)

            self.checkboxMergeTabeles = QCheckBox('Объединить листы')
            self.checkboxMergeTabeles.setChecked(True)
            self.suggsLayout.addWidget(self.checkboxMergeTabeles)

            self.okButton = QPushButton("сохранить", self)
            self.okButton.clicked.connect(self.saveAction)
            self.suggsLayout.addWidget(self.okButton)

            self.suggs = QWidget(self)
            self.suggs.setLayout(self.suggsLayout)
            self.layout.addWidget(self.suggs)

            self.layout.setStretch(0, 4)  # Устанавливаем относительный размер левой пустой области
            self.layout.setStretch(1, 3)

        except Exception as error:
            msg = QMessageBox()
            msg.setWindowIcon(QtGui.QIcon('logo.ico'))
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Ошибка")
            msg.setText("Неизвестная ошибка: дальнейшая работа программы может быть некорректной")
            msg.setInformativeText(f"{error}")
            msg.exec_()

    def on_selectionChangedMain(self, selected, deselected):
        if selected:
            self.sugTableWidget.clearSelection()
            self.sug_comment.clear()
            values = None
            for index in selected.indexes():
                self.current_row = index.row()
                values = self.df.iloc[self.current_row]
            self.suggestions = self.suggestions_list[self.current_row]

            for i in range(4):
                self.currentItemTableWidget.setItem(0, i, (QTableWidgetItem(str(values.iloc[i]))))

            warns = ""
            if self.warn_dict[self.current_row]:
                warns = warns + self.warn_dict[self.current_row] + "\n"
            if measure_check(values.iloc[3]) is None:
                warns = warns + f"Незнакомая единица измерения: <b>{values.iloc[3]}</b>"

            self.orig_comment.setText(warns)

            for i in range(len(self.columns_with_add_info)):
                self.originalItemTableWidget.setItem(0, i, QTableWidgetItem(
                    str(self.df_with_add_info.iloc[self.current_row, i])))
                self.color_row_grey(self.originalItemTableWidget, 0)

            self.set_tooltips(self.originalItemTableWidget)

            if self.suggestions[0]:
                self.sugTableWidget.setRowCount(len(self.suggestions))
                for i, suggestion in enumerate(self.suggestions):
                    self.sugTableWidget.setItem(i, 0, (QTableWidgetItem(suggestion[1])))
                    self.sugTableWidget.setItem(i, 1, (QTableWidgetItem(suggestion[0])))
                    self.sugTableWidget.setItem(i, 2, (QTableWidgetItem(suggestion[2])))

                    if suggestion[3] < 5:
                        self.color_row_green(self.sugTableWidget, i)
                    elif suggestion[3] < 8:
                        self.color_row_blue(self.sugTableWidget, i)
                    else:
                        self.color_row_yellow(self.sugTableWidget, i)
                self.set_tooltips(self.sugTableWidget)

            else:
                self.sugTableWidget.clearContents()
                self.sugTableWidget.setRowCount(0)

    def on_selectionChangedSuggs(self, selected, deselected):
        if selected:
            current_sug = None
            values = None
            for index in selected.indexes():
                current_sug = index.row()
                values = [self.sugTableWidget.item(current_sug, col).text() for col in
                          range(self.sugTableWidget.columnCount())]

            self.sug_comment.setText(self.suggestions[current_sug][4])

            self.currentItemTableWidget.setItem(0, 0, (QTableWidgetItem(values[0])))
            self.currentItemTableWidget.setItem(0, 1, (QTableWidgetItem(values[1])))
            self.currentItemTableWidget.setItem(0, 3, (QTableWidgetItem(values[2])))

    def populate_table(self, tableWidget, frame):
        tableWidget.setRowCount(frame.shape[0])
        tableWidget.setColumnCount(6)
        tableWidget.setHorizontalHeaderLabels(['Артикул', 'Исходные', 'Товары (работы, услуги) ', 'Исходные', 'Кол-во', 'Ед.', 'Файл'])

        for i in range(frame.shape[0]):
            for j in range(6):
                if j == 1 or j == 3 or j == 4 or j == 5:
                    tableWidget.setItem(i, j, QTableWidgetItem(str(frame.iloc[i, j])))

    def adjust_table_height(self, tableWidget):
        tableWidget.resizeRowsToContents()
        total_height = tableWidget.horizontalHeader().height() + tableWidget.verticalHeader().length()

        tableWidget.setMinimumHeight(total_height)
        tableWidget.setMaximumHeight(total_height)

    def set_tooltips(self, tableWidget):
        for row in range(tableWidget.rowCount()):
            for column in range(tableWidget.columnCount()):
                item = tableWidget.item(row, column)
                if item is not None:
                    item.setToolTip(item.text())

    def replaceAction(self):
        if self.current_row is not None:
            standard_article = self.currentItemTableWidget.item(0, 0).text()
            standard_name = self.currentItemTableWidget.item(0, 1).text()
            measure = self.currentItemTableWidget.item(0, 3).text()
            if standard_name in self.data:
                if pd.notna(standard_article) and self.data[standard_name]['standard_article'] != standard_article:
                    conf = ConfirmTableCollision()
                    result = conf.showDialog(standard_name, standard_article, None,
                                             self.data[standard_name]['standard_article'])
                    if result == QMessageBox.Cancel:
                        return

                if pd.notna(measure) and self.data[standard_name]['measure'] != measure:
                    conf = ConfirmTableCollision()
                    result = conf.showDialog(standard_name, None, measure, self.data[standard_name]['measure'])
                    if result == QMessageBox.Cancel:
                        return

            self.sugTableWidget.clearSelection()
            cols = [0, 2, 4, 5]
            for i in range(4):
                val = self.currentItemTableWidget.item(0, i)
                self.tableWidget.setItem(self.current_row, cols[i], (QTableWidgetItem(val)))
                datatype = type(self.df.iloc[self.current_row, i])

                if val.text().isdigit():
                    if datatype != val.text():
                        # Приведение значения к правильному типу данных
                        if datatype is int or datatype is np.int64:
                            self.df.iloc[self.current_row, i] = int(val.text())
                        elif datatype is float or datatype is np.float64:
                            self.df.iloc[self.current_row, i] = float(val.text())
                    else:
                        self.df.iloc[self.current_row, i] = val.text()
                else:
                    if i == 2:
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Warning)
                        msg.setWindowTitle("Предупреждение")
                        msg.setWindowIcon(QtGui.QIcon('logo.ico'))
                        msg.setText(f"В поле количества не числовые данные")
                        msg.exec_()
                    self.df.iloc[self.current_row, i] = val.text()

                item = self.tableWidget.item(self.current_row, i)

                if item is not None:
                    item.setToolTip(item.text())
            self.color_main_row_green(self.tableWidget, self.current_row)

            if self.checkboxSaveReplacement.isChecked():
                article = self.originalItemTableWidget.item(0, 0).text()
                if article == standard_article or pd.isna(article):
                    article = ""
                name = self.originalItemTableWidget.item(0, 1).text()
                if name == standard_name or pd.isna(name):
                    name = ""

                self.replacements_dict[self.current_row] = [standard_name, standard_article, measure, name, article]

                if standard_name in self.data:
                    if pd.notna(standard_article) and self.data[standard_name]['standard_article'] != standard_article:
                        conf = ConfirmDBCollision()
                        result = conf.showDialog(standard_name, standard_article, None,
                                                 self.data[standard_name]['standard_article'])

                        if result == QMessageBox.Cancel:
                            self.replacements_dict[self.current_row][1] = self.data[standard_name]['standard_article']

                    if pd.notna(measure) and self.data[standard_name]['measure'] != measure:
                        conf = ConfirmDBCollision()
                        result = conf.showDialog(standard_name, None, measure, self.data[standard_name]['measure'])

                        if result == QMessageBox.Cancel:
                            self.replacements_dict[self.current_row][2] = self.data[standard_name]['standard_article']

    def resetAction(self):
        try:
            if self.current_row is not None:
                self.sugTableWidget.clearSelection()
                for i in range(4):
                    val = self.df.iloc[self.current_row, i]
                    self.currentItemTableWidget.setItem(0, i, (QTableWidgetItem(str(val))))
        except Exception as error:
            msg = QMessageBox()
            msg.setWindowIcon(QtGui.QIcon('logo.ico'))
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Ошибка")
            msg.setText("Неизвестная ошибка при выполнении сброса: дальнейшая работа программы может быть некорректной")
            msg.setInformativeText(f"{error}")
            msg.exec_()

    def deleteAction(self):
        pass
        # if self.current_row is not None:
        #    msg = QMessageBox.question(
        #        self,
        #        "Внимание подтвердите удаление строки!",
        #        "Вы действительно хотите удалить "
        #        f"строку <b style='color: red;'>{self.current_row + 1}</b> ?",
        #        QMessageBox.Ok | QMessageBox.Cancel
        #    )
        #    if msg == QMessageBox.Cancel:
        #        return
        #    self.tableWidget.removeRow(self.current_row)
        #    try:
        #        self.df.drop(index=self.current_row, inplace=True)
        #    except Exception as error:
        #        print(error)
        #    self.df.reset_index(drop=True, inplace=True)
        #    self.current_row = None
        #    self.tableWidget.clearSelection()
        #    self.currentItemTableWidget.clearContents()
        #    self.sugTableWidget.clearContents()
        #    self.sugTableWidget.setRowCount(0)
        #    self.sug_comment.clear()
        #    self.orig_comment.clear()

    def saveAction(self):

        output, _ = QFileDialog.getSaveFileName(self, "Сохранение файла",
                                                '',
                                                "Excel Files (*.xlsx)")
        if output:

            # Создаем пустой словарь для хранения результирующих датафреймов
            result_dict = {}
            if self.checkboxMergeTabeles.isChecked():
                key = self.df.iloc[0, 4]
                result_dict[key] = self.df.iloc[:, :4]

            else:
                # Проходим по значениям в пятой колонке
                for key in self.df.iloc[:, 4].unique():
                    # Фильтруем строки, соответствующие текущему значению ключа
                    filtered_df = self.df[self.df.iloc[:, 4] == key]
                    # Сохраняем только первые четыре колонки
                    result_df = filtered_df.iloc[:, :4]
                    # Добавляем полученный датафрейм в словарь
                    result_dict[key] = result_df

            dict_to_excel(result_dict, output)

            for key, item in self.replacements_dict.items():
                standard_name = item[0]
                standard_article = item[1]
                measure = item[2]
                name = item[3]
                article = item[4]

                if standard_name not in self.data:
                    save_new_object(self.data, standard_name, standard_article, measure, name, article)

                else:
                    if pd.notna(standard_article) and self.data[standard_name]['standard_article'] != standard_article:
                        save_non_standard_article(self.data, self.data[standard_name]['standard_article'],
                                                  standard_name)
                        self.data[standard_name]['standard_article'] = standard_article

                    if pd.notna(name):
                        save_non_standard_name(self.data, name, standard_name)

                    if pd.notna(measure) and self.data[standard_name]['measure'] != measure:
                        self.data[standard_name]['measure'] = measure

            with open("standards.json", "w", encoding="utf-8") as write_file:
                json.dump(self.data, write_file, ensure_ascii=False, indent=4)

    def backAction(self):
        try:
            conf = ConfirmBack()
            result = conf.showDialog()
            if result == QMessageBox.Cancel:
                return
            dialog = Marker(self.filename)
            dialog.show()
            self.close()
            dialog.exec()
        except Exception as error:
            print(error)

    def on_item_changed(self, item):
        if item.column() == 3:
            if measure_check(item.text()) is None:
                item.setBackground(QBrush(QColor(250, 128, 114)))
            else:
                item.setBackground(QBrush(QColor(Qt.white)))

    def color_row_green(self, tableWidget, row):
        color = QColor(150, 250, 150)  # Цвет фона (светло-зелёный)

        for column in range(tableWidget.columnCount()):
            item = tableWidget.item(row, column)
            if item is not None:
                item.setBackground(QBrush(color))

    def color_main_row_green(self, tableWidget, row):
        color1 = QColor(150, 250, 150)  # Цвет фона (светло-зелёный)
        color2 = QColor(91, 229, 91)  # Цвет фона (тёмно-зелёный)

        for column in range(tableWidget.columnCount()):
            item = tableWidget.item(row, column)
            if item is not None:
                if column == 1 or column == 3:
                    item.setBackground(QBrush(color2))
                else:
                    item.setBackground(QBrush(color1))

    def color_row_yellow(self, tableWidget, row):
        color = QColor(255, 255, 200)  # Цвет фона (светло-жёлтый)

        for column in range(tableWidget.columnCount()):
            item = tableWidget.item(row, column)
            if item is not None:
                item.setBackground(QBrush(color))

    def color_row_blue(self, tableWidget, row):
        color = QColor(127, 199, 255)  # Цвет фона (голубой)
        for column in range(tableWidget.columnCount()):
            item = tableWidget.item(row, column)
            if item is not None:
                item.setBackground(QBrush(color))

    def color_row_grey(self, tableWidget, row):
        color = QColor(235, 235, 235)
        for column in range(tableWidget.columnCount()):
            item = tableWidget.item(row, column)
            if item is not None:
                item.setBackground(QBrush(color))

    def color_columns_grey(self, columns_indexes, tableWidget):
        color = QColor(235, 235, 235)
        for column_index in columns_indexes:
            for row in range(tableWidget.rowCount()):
                item = tableWidget.item(row, column_index)
                if item:
                    item.setBackground(color)


class ConfirmBack(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.resize(250, 150)
        self.setWindowTitle('Подтверждение')
        self.setWindowIcon(QtGui.QIcon('logo.ico'))

    def showDialog(self):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        msg.setText("Вы уверены, что хотите вернуться назад?")
        msg.setInformativeText("Все несохранённые данные будут потеряны")

        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

        msg.button(QMessageBox.Ok).setText("да")
        msg.button(QMessageBox.Cancel).setText("нет")

        retval = msg.exec_()
        return retval


class ConfirmDBCollision(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.resize(250, 150)
        self.setWindowTitle('Подтверждение')
        self.setWindowIcon(QtGui.QIcon('logo.ico'))

    def showDialog(self, standard_name, new_article, new_measure, standard):
        msg = QMessageBox(self)
        if new_article:
            msg.setText(f"Вы уверены, что хотите обновить артикул в базе данных у объекта <b>{standard_name}</b>?")
            msg.setInformativeText(f"Заменить <b>{standard}</b> на новый <b>{new_article}</b>?")
        else:
            msg.setText(
                f"Вы уверены, что хотите обновить единицу измерения в базе данных у объекта <b>{standard_name}</b>?")
            msg.setInformativeText(f"Заменить <b>{standard}</b> на новую <b>{new_measure}</b>?")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

        msg.button(QMessageBox.Ok).setText("да")
        msg.button(QMessageBox.Cancel).setText("нет")

        retval = msg.exec_()
        return retval


class ConfirmTableCollision(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.resize(250, 150)
        self.setWindowTitle('Подтверждение')
        self.setWindowIcon(QtGui.QIcon('logo.ico'))

    def showDialog(self, standard_name, new_article, new_measure, standard):
        msg = QMessageBox(self)
        msg.setText("Вы уверены, что хотите провести замену?")
        if new_article:
            msg.setInformativeText(f"У объекта <b>{standard_name}</b> в базе данных сохранён другой артикул\n"
                                   f"Артикул из базы: <b>{standard}</b>\n"
                                   f"Артикул введённая вами: <b>{new_article}</b>")
        else:
            msg.setInformativeText(f"У объекта <b>{standard_name}</b> в базе данных сохранён другая единица измерения\n"
                                   f"Единица из базы: <b>{standard}</b>\n"
                                   f"Единица введённая вами: <b>{new_measure}</b>")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

        msg.button(QMessageBox.Ok).setText("да")
        msg.button(QMessageBox.Cancel).setText("нет")

        retval = msg.exec_()
        return retval


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = Main()
    main.show()
    sys.exit(app.exec_())
