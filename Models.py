from PyQt5.QtSql import QSqlRelationalTableModel, QSqlRelation, QSqlTableModel,\
    QSqlDatabase
from PyQt5.QtCore import Qt, QSize

class ScheduleModel(QSqlRelationalTableModel):
    def __init__(self, parent=None, db=QSqlDatabase()):
        QSqlRelationalTableModel.__init__(self, parent, db)
        self.setSort(0, Qt.AscendingOrder)
        self.headerArr = ['id', 'Кафедра', 'Группа',
                          'Неделя', 'День', '№ пары', 'Пара']

        self.setTable('schedule')
        self.setEditStrategy(QSqlTableModel.OnManualSubmit)
        self.setRelation(1, QSqlRelation('pulpits', 'id', 'name'))
        self.setRelation(2, QSqlRelation('groups', 'id', 'name'))
        self.setRelation(3, QSqlRelation('weeks', 'id', 'name'))
        self.setRelation(4, QSqlRelation('days', 'id', 'name'))
        self.select()

    def headerData(self, p_int, Qt_Orientation, role=None):
        if Qt_Orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                if p_int in range(7):
                    return self.headerArr[p_int]
        return None

class PulpitsModel(QSqlRelationalTableModel):
    def __init__(self, parent=None, db=QSqlDatabase()):
        QSqlRelationalTableModel.__init__(self, parent, db)
        self.setSort(0, Qt.AscendingOrder)
        self.setEditStrategy(QSqlTableModel.OnManualSubmit)
        self.setTable('pulpits')
        self.select()

    def headerData(self, p_int, Qt_Orientation, role=None):
        if Qt_Orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                if p_int == 0:
                    return 'id'
                elif p_int == 1:
                    return 'Название кафедры'

        return None

class GroupsModel(QSqlRelationalTableModel):
    def __init__(self, parent=None, db=QSqlDatabase()):
        QSqlRelationalTableModel.__init__(self, parent, db)
        self.setSort(0, Qt.AscendingOrder)
        self.setEditStrategy(QSqlTableModel.OnManualSubmit)
        self.setTable('groups')
        self.select()

    def headerData(self, p_int, Qt_Orientation, role=None):
        if Qt_Orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                if p_int == 0:
                    return 'id'
                elif p_int == 1:
                    return 'Навзание группы'

        return None
