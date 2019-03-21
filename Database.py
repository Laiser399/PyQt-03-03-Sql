from PyQt5.QtSql import QSqlQuery, QSqlError, QSqlDatabase, QSql



ms_name = 'QSQLITE'
name = 'DB1'
user_name = 'root'
user_password = 'JceNjg399_^@'


create_schedule_query = 'CREATE TABLE schedule (' \
                        'id INTEGER PRIMARY KEY AUTOINCREMENT,' \
                        'id_pulpit INT,' \
                        'id_group INT,' \
                        'id_week INT,' \
                        'id_day INT,' \
                        'n_lesson INT,' \
                        'lesson VARCHAR);'
create_pulpits_query = 'CREATE TABLE pulpits (' \
                       'id INTEGER PRIMARY KEY AUTOINCREMENT,' \
                       'name VARCHAR);'
create_groups_query = 'CREATE TABLE groups (' \
                      'id INTEGER PRIMARY KEY AUTOINCREMENT,' \
                      'name VARCHAR);'
create_weeks_query = 'CREATE TABLE weeks (' \
                     'id INTEGER PRIMARY KEY AUTOINCREMENT,' \
                     'name VARCHAR(2));'
create_days_query = 'CREATE TABLE days (' \
                    'id INTEGER PRIMARY KEY AUTOINCREMENT,' \
                    'name VARCHAR);'

insert_weeks_query = []
insert_weeks_query.append('INSERT INTO weeks (id, name) VALUES (1, "I");')
insert_weeks_query.append('INSERT INTO weeks (id, name) VALUES (2, "II");')

insert_days_query = []
insert_days_query.append('INSERT INTO days (id, name) VALUES (1, "Пн");')
insert_days_query.append('INSERT INTO days (id, name) VALUES (2, "Вт");')
insert_days_query.append('INSERT INTO days (id, name) VALUES (3, "Ср");')
insert_days_query.append('INSERT INTO days (id, name) VALUES (4, "Чт");')
insert_days_query.append('INSERT INTO days (id, name) VALUES (5, "Пт");')
insert_days_query.append('INSERT INTO days (id, name) VALUES (6, "Сб");')
insert_days_query.append('INSERT INTO days (id, name) VALUES (7, "Вс");')


def createConnection(dbms_name=ms_name,
                     db_name=name,
                     db_user_name=user_name,
                     db_user_password=user_password):
    db = QSqlDatabase.addDatabase(dbms_name)
    if not db.isValid():
        print('DEBUG: DLL драйверов для {1} не найдены: {0}'.format(
            db.lastError().text(),
            dbms_name
        ))

        print('drivers: ', QSqlDatabase.drivers())
        print('error id: ', db.lastError().type())
        print('isAvailable: ', QSqlDatabase.isDriverAvailable('QMYSQL'))
        return None

    db.setDatabaseName(db_name)
    #db.setDatabaseName('table1')
    db.setUserName(db_user_name)
    db.setPassword(db_user_password)
    #db.setPort(3306)
    if not db.open():
        print('DEBUG: Невозможно подключиться к указанной базе данных!')
        print('Last error: ', db.lastError().text())
        print('Last error id: ', db.lastError().type())
        return None
    return db

def deleteAllTables(db: QSqlDatabase):
    if db == None or not db.isValid():
        return False
    tables_list = db.tables()
    sql = QSqlQuery()
    query = 'DROP TABLE '
    #sql.prepare('DROP TABLE (?)')
    for i in tables_list:
        #sql.bindValue(0, i)
        sql.exec(query + i)
        if sql.lastError().type() != QSqlError.NoError:
            print('ERROR: ', sql.lastError().text())
    return True

def createTables(db: QSqlDatabase):
    if db == None or not db.isValid():
        print('БД не допустима!')
        return False
    sql = QSqlQuery()

    db_tables = db.tables(QSql.Tables)
    if 'schedule' in db_tables \
            or 'pulpits' in db_tables \
            or 'groups' in db_tables \
            or 'weeks' in db_tables \
            or 'days' in db_tables:
        print('Таблицы не созданы, т.к. одна из них уже существует!')
        return False

    sql.exec(create_schedule_query)
    if sql.lastError().type() != QSqlError.NoError:
        print('Ошибка создания таблицы "schedule": ', sql.lastError().text())
        return False

    sql.exec(create_pulpits_query)
    if sql.lastError().type() != QSqlError.NoError:
        print('Ошибка создания таблицы "pulpits": ', sql.lastError().text())
        return False

    sql.exec(create_groups_query)
    if sql.lastError().type() != QSqlError.NoError:
        print('Ошибка создания таблицы "groups": ', sql.lastError().text())
        return False

    sql.exec(create_weeks_query)
    if sql.lastError().type() != QSqlError.NoError:
        print('Ошибка создания таблицы "weeks": ', sql.lastError().text())
        return False
    for query in insert_weeks_query:
        sql.exec(query)

    sql.exec(create_days_query)
    if sql.lastError().type() != QSqlError.NoError:
        print('Ошибка создания таблицы "days": ', sql.lastError().text())
        return False
    for query in insert_days_query:
        sql.exec(query)

    print('Таблицы созданы!')
    return True
