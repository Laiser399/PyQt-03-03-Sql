from PyQt5.QtWidgets import QTableView, QAbstractItemView, QMessageBox
from PyQt5.QtSql import QSqlRelationalDelegate
from PyQt5.Qt import pyqtSignal

from PyQt5.QtCore import QAbstractItemModel

class SQLView(QTableView):
    # сигнал изменения модели
    currentModelChanged = pyqtSignal(QAbstractItemView)
    changedData = pyqtSignal()
    def __init__(self, parent=None):
        QTableView.__init__(self, parent)
        self.setItemDelegate(QSqlRelationalDelegate())
        self.setAlternatingRowColors(True)
        self.resizeRowsToContents()
        self.resizeColumnsToContents()

        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.horizontalHeader().setHighlightSections(False)

    def resizeEvent(self, *args, **kwargs):
        self.resizeToContent()

    def setModel(self, model: QAbstractItemModel):
        QTableView.setModel(self, model)
        self.currentModelChanged.emit(model)

        self.resizeRowsToContents()
        self.resizeColumnsToContents()

    def resizeToContent(self):
        self.resizeRowsToContents()
        self.resizeColumnsToContents()

    def saveData(self):
        model = self.model()
        if model == None:
            return

        model.database().transaction()
        if model.submitAll():
            model.database().commit()
            self.changedData.emit()
            self.resizeToContent()
        else:
            model.database().rollback()
            QMessageBox.warning(self, 'Ошибка',
                                'БД сообщила об ошибке: {0}'.format(model.lastError().text()))
            model.select()

    def revertAll(self):
        model = self.model()
        if model == None:
            return
        model.revertAll()

    def insertRow(self):
        model = self.model()
        if model == None:
            return

        insertIndex = self.currentIndex()
        row = None
        if insertIndex.row() == -1:
            row = 0
        else:
            row = insertIndex.row()

        model.insertRow(row)
        insertIndex = model.index(row, 0)
        self.setCurrentIndex(insertIndex)

    def deleteRow(self):
        model = self.model()
        if model == None:
            return

        indexList = self.selectionModel().selectedIndexes()
        for item in indexList:
            print('row: ', item.row())
            model.removeRow(item.row())
        model.submitAll()
        model.select()
        self.changedData.emit()
