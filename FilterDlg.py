from PyQt5.QtWidgets import QDialog, QGridLayout,\
    QLineEdit, QLabel, QPushButton
from PyQt5.QtCore import Qt

class FilterDlg(QDialog):
    good = False
    controlled_close = False
    SaveParams = ['', '', '', '', '', '', '']
    EditArr = []
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)

        #массив Edit'ов
        for i in range(7):
            self.EditArr.append(QLineEdit())
        #кнопки
        bOk = QPushButton('Ok')
        bCancel = QPushButton('Cancel')
        bOk.clicked.connect(self.__slot_Ok)
        bCancel.clicked.connect(self.__slot_Cancel)

        #создание и заполнение Layout
        grLay = QGridLayout()
        grLay.addWidget(QLabel('id'), 0, 0, Qt.AlignRight)
        grLay.addWidget(QLabel('Кафедра:'), 1, 0, Qt.AlignRight)
        grLay.addWidget(QLabel('Группа:'), 2, 0, Qt.AlignRight)
        grLay.addWidget(QLabel('Неделя:'), 3, 0, Qt.AlignRight)
        grLay.addWidget(QLabel('День:'), 4, 0, Qt.AlignRight)
        grLay.addWidget(QLabel('№ пары:'), 5, 0, Qt.AlignRight)
        grLay.addWidget(QLabel('Пара:'), 6, 0, Qt.AlignRight)
        for i in range(7):
            grLay.addWidget(self.EditArr[i], i, 1)
        grLay.addWidget(bOk, 7, 0)
        grLay.addWidget(bCancel, 7, 1)
        self.setLayout(grLay)

        self.setWindowTitle('Фильтр')

    def showEvent(self, QShowEvent):
        self.controlled_close = False
        #сохранение параметров на случай отмены
        for i in range(7):
            self.SaveParams[i] = self.EditArr[i].text()

    def __slot_Ok(self):
        self.controlled_close = True
        self.good = True
        self.close()

    def __slot_Cancel(self):
        self.controlled_close = True
        self.good = False
        for i in range(7):
            self.EditArr[i].setText(self.SaveParams[i])
        self.close()

    def closeEvent(self, QCloseEvent):
        if not self.controlled_close:
            self.good = False
            for i in range(7):
                self.EditArr[i].setText(self.SaveParams[i])

    def getFilter(self):
        self.exec()
        if self.good:
            res = [None, None, None, None, None, None, None]
            # arrNames = ['id', 'pulpits.name', 'groups.name', 'weeks.name', 'days.name',
            #             'n_lesson', 'lesson']
            for i in range(7):
                str = self.EditArr[i].text()
                str = str.lstrip()
                if len(str) > 0:
                    res[i] = self.EditArr[i].text()

            return res

        else:
            return None


