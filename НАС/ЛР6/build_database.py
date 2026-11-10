import os
import sys
import win32com.client

DIR_LR6 = r"C:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР6"
DB_PATH = os.path.join(DIR_LR6, "Продажи.accdb")

def create_database():
    if os.path.exists(DB_PATH):
        try:
            os.remove(DB_PATH)
            print("Existing DB removed.")
        except Exception as e:
            print("Cannot remove old DB:", e)

    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False
    access.NewCurrentDatabase(DB_PATH)
    print("New database created at:", DB_PATH)

    db = access.CurrentDb()

    # 1. Create Tables
    print("Creating tables...")

    # Склады
    sql_skladi = """
    CREATE TABLE [Склады] (
        [Номер склада] LONG CONSTRAINT PK_Склады PRIMARY KEY,
        [Телефон] TEXT(50),
        [Адрес] TEXT(50),
        [Заведующий] TEXT(50)
    );
    """
    db.Execute(sql_skladi)

    # Фирмы
    sql_firmy = """
    CREATE TABLE [Фирмы] (
        [Код фирмы] LONG CONSTRAINT PK_Фирмы PRIMARY KEY,
        [Название фирмы] TEXT(50),
        [Адрес] TEXT(50),
        [Телефон] TEXT(50)
    );
    """
    db.Execute(sql_firmy)

    # Товары
    sql_tovary = """
    CREATE TABLE [Товары] (
        [Код товара] AUTOINCREMENT CONSTRAINT PK_Товары PRIMARY KEY,
        [Наименование] TEXT(50),
        [Марка] TEXT(50),
        [Цена] CURRENCY,
        [Номер склада] LONG,
        [Количество] LONG,
        [Описание] MEMO,
        [Изображение] LONGBINARY
    );
    """
    db.Execute(sql_tovary)

    # Контактные лица
    sql_kontakty = """
    CREATE TABLE [Контактные лица] (
        [Код сотрудника] LONG CONSTRAINT PK_Контактные_лица PRIMARY KEY,
        [Фамилия] TEXT(50),
        [Имя] TEXT(50),
        [Дата рождения] DATETIME,
        [Домашний адрес] TEXT(50),
        [Домашний телефон] TEXT(50),
        [Должность] TEXT(50),
        [Код фирмы] LONG
    );
    """
    db.Execute(sql_kontakty)

    # Продажи
    sql_prodazhi = """
    CREATE TABLE [Продажи] (
        [Дата продажи] DATETIME,
        [Код фирмы] LONG,
        [Код товара] LONG,
        [Количество] LONG,
        [Скидка] SINGLE,
        CONSTRAINT PK_Продажи PRIMARY KEY ([Дата продажи], [Код фирмы], [Код товара])
    );
    """
    db.Execute(sql_prodazhi)
    print("All 5 base tables created.")

    # 2. Establish Relationships (DAO Relations)
    print("Creating relationships with referential integrity and cascade updates/deletions...")
    # dbRelationUpdateCascade = 256, dbRelationDeleteCascade = 4096
    attr_cascade = 256 + 4096

    def add_relation(rel_name, primary_tbl, foreign_tbl, p_field, f_field):
        try:
            rel = db.CreateRelation(rel_name, primary_tbl, foreign_tbl, attr_cascade)
            fld = rel.CreateField(p_field)
            fld.ForeignName = f_field
            rel.Fields.Append(fld)
            db.Relations.Append(rel)
            print(f"Relation '{rel_name}' created ({primary_tbl}.{p_field} -> {foreign_tbl}.{f_field}).")
        except Exception as e:
            print(f"Relation error {rel_name}:", e)

    add_relation("Склады_Товары", "Склады", "Товары", "Номер склада", "Номер склада")
    add_relation("Фирмы_Контакты", "Фирмы", "Контактные лица", "Код фирмы", "Код фирмы")
    add_relation("Фирмы_Продажи", "Фирмы", "Продажи", "Код фирмы", "Код фирмы")
    add_relation("Товары_Продажи", "Товары", "Продажи", "Код товара", "Код товара")

    # 3. Populate Data
    print("Populating data...")
    # Склады
    skladi_data = [
        (10, '8(495)953-0189', 'ул. Свободы,37', 'Иванов П.П.'),
        (20, '8(499)712-2222', 'ул. Мелитопольская, 35', 'Гривко Л. К.'),
        (30, '8(495)901-7044', 'ул. Кустарная, 28', 'Соломоник К.Ф.'),
        (40, '8(495)359-0014', 'ул. Перерва, 2', 'Андреев П.Н.')
    ]
    for r in skladi_data:
        db.Execute(f"INSERT INTO [Склады] VALUES ({r[0]}, '{r[1]}', '{r[2]}', '{r[3]}');")

    # Фирмы
    firmy_data = [
        (100, 'Мир', 'Чонгарский б-р, 16', '8(495) 152-4001'),
        (200, 'М. Видео', 'Маросейка, 6/8', '8(495) 923-2906'),
        (300, 'Диал Электроникс', 'Новослободская, 14/19', '8(495) 978-1693')
    ]
    for r in firmy_data:
        db.Execute(f"INSERT INTO [Фирмы] VALUES ({r[0]}, '{r[1]}', '{r[2]}', '{r[3]}');")

    # Товары (Задание 2 п. 4, 9, 10: добавление NV-DX1E и замена GoldStar на LG Electronics)
    tovary_data = [
        ('Телевизор', 'LG Electronics СМ-2180К', 459, 10, 15, 'Цветной телевизор 21" со стереозвуком'),
        ('Телевизор', 'Philips 25PT9001', 1499, 10, 8, 'Телевизор 25" с плоским экраном BlackLine-S'),
        ('Телевизор', 'Panasonic 25V50R', 765, 10, 12, 'Стереофонический телевизор 25", телетекст'),
        ('Телевизор', 'LG Electronics CF-14E20B', 230, 10, 20, 'Компактный телевизор 14" для кухни/дачи'),
        ('Видеомагнитофон', 'Panasonic HS-800EE', 1400, 20, 10, 'Hi-Fi стерео 6-головочный видеомагнитофон'),
        ('Видеомагнитофон', 'Philips VR-756', 450, 20, 18, '4-головочный видеомагнитофон ShowView'),
        ('Видеокамера', 'Samsung VP-J55', 530, 30, 14, 'Аналоговая видеокамера Hi8, оптический зум 22x'),
        ('Видеокамера', 'Sharp E37', 845, 30, 9, 'Компактная видеокамера с ЖК-дисплеем 2.5"'),
        ('Музыкальный центр', 'Panasonic DH32', 320, 40, 25, 'Минисистема с 3-дисковым CD-чейнджером'),
        ('Музыкальный центр', 'Sony MJ-L1', 1289, 40, 7, 'Микросистема Hi-Fi класса с декой MiniDisc'),
        ('Видеокамера', 'Panasonic NV-DX1E', 2599, 30, 5, 'Профессиональная 3-матричная цифровая видеокамера')
    ]
    for r in tovary_data:
        db.Execute(f"INSERT INTO [Товары] ([Наименование], [Марка], [Цена], [Номер склада], [Количество], [Описание]) VALUES ('{r[0]}', '{r[1]}', {r[2]}, {r[3]}, {r[4]}, '{r[5]}');")

    # Контактные лица (date format in SQL: #YYYY-MM-DD#)
    kontakty_data = [
        (1, 'Смирнов', 'Алексей', '1978-05-14', 'Ленинский пр-т, 45', '8(495)334-1122', 'Директор', 100),
        (2, 'Ковалева', 'Елена', '1985-09-22', 'ул. Профсоюзная, 12', '8(495)421-5566', 'Менеджер', 100),
        (3, 'Васильев', 'Дмитрий', '1975-11-03', 'Кутузовский пр-т, 18', '8(495)240-8899', 'Директор', 200),
        (4, 'Николаева', 'Ольга', '1989-02-17', 'Варшавское ш., 70', '8(495)119-4433', 'Менеджер', 200),
        (5, 'Федоров', 'Михаил', '1981-08-28', 'пр-т Мира, 102', '8(495)687-2211', 'Директор', 300),
        (6, 'Морозова', 'Анна', '1987-12-09', 'ул. Тверская, 9', '8(495)925-7744', 'Менеджер', 300)
    ]
    for r in kontakty_data:
        db.Execute(f"INSERT INTO [Контактные лица] VALUES ({r[0]}, '{r[1]}', '{r[2]}', #{r[3]}#, '{r[4]}', '{r[5]}', '{r[6]}', {r[7]});")

    # Продажи
    prodazhi_data = [
        ('2024-01-15', 100, 1, 3, 0),
        ('2024-01-18', 200, 2, 2, 5),
        ('2024-01-22', 300, 3, 4, 10),
        ('2024-02-05', 100, 4, 5, 0),
        ('2024-02-12', 200, 5, 2, 10),
        ('2024-02-20', 300, 6, 3, 0),
        ('2024-03-10', 100, 7, 2, 5),
        ('2024-03-15', 200, 8, 1, 0),
        ('2024-03-25', 300, 9, 5, 10),
        ('2024-04-02', 100, 10, 1, 0),
        ('2024-04-14', 200, 1, 2, 5),
        ('2024-04-28', 300, 2, 1, 0)
    ]
    for r in prodazhi_data:
        db.Execute(f"INSERT INTO [Продажи] VALUES (#{r[0]}#, {r[1]}, {r[2]}, {r[3]}, {r[4]});")

    print("Data population completed.")

    # 4. Create Queries (Задание 3 и Задание 4)
    print("Creating queries...")

    queries = [
        ("Запрос_Товары_Склады",
         """SELECT Товары.Наименование, Товары.Марка, Товары.Цена, Склады.[Номер склада], Склады.Телефон
            FROM Склады INNER JOIN Товары ON Склады.[Номер склада] = Товары.[Номер склада];"""),

        ("Запрос_Продажи_Телевизоров",
         """SELECT Продажи.[Дата продажи], Товары.Наименование, Товары.Марка, Продажи.Количество, Фирмы.[Название фирмы], Фирмы.Адрес, Фирмы.Телефон
            FROM Фирмы INNER JOIN (Товары INNER JOIN Продажи ON Товары.[Код товара] = Продажи.[Код товара]) ON Фирмы.[Код фирмы] = Продажи.[Код фирмы]
            WHERE Товары.Наименование='Телевизор'
            ORDER BY Продажи.[Дата продажи], Товары.Марка;"""),

        ("Запрос_Продажи_Со_Скидкой",
         """SELECT Товары.Наименование, Товары.Марка, Фирмы.[Название фирмы], Продажи.Количество, Продажи.Скидка
            FROM Фирмы INNER JOIN (Товары INNER JOIN Продажи ON Товары.[Код товара] = Продажи.[Код товара]) ON Фирмы.[Код фирмы] = Продажи.[Код фирмы]
            WHERE Продажи.Скидка > 0
            ORDER BY Товары.Наименование, Фирмы.[Название фирмы];"""),

        ("Запрос_Цены_Со_Скидкой",
         """SELECT Товары.Наименование, Товары.Марка, Товары.Цена, Продажи.Скидка,
                   IIf(Продажи.Скидка>0, Товары.Цена*(1-Продажи.Скидка/100), Товары.Цена) AS [Цена со скидкой]
            FROM Товары INNER JOIN Продажи ON Товары.[Код товара] = Продажи.[Код товара];"""),

        ("Запрос_Директора_Фирм",
         """SELECT [Контактные лица].Фамилия, [Контактные лица].Имя, [Контактные лица].[Домашний телефон], [Контактные лица].[Домашний адрес],
                   Фирмы.[Название фирмы], Фирмы.Телефон AS [Рабочий телефон], [Контактные лица].Должность
            FROM Фирмы INNER JOIN [Контактные лица] ON Фирмы.[Код фирмы] = [Контактные лица].[Код фирмы]
            WHERE [Контактные лица].Должность='Директор'
            ORDER BY [Контактные лица].Фамилия;"""),

        ("Запрос_Товары_По_Цене",
         """PARAMETERS [Введите предельную стоимость:] Currency;
            SELECT Товары.Наименование, Товары.Марка, Товары.Цена
            FROM Товары
            WHERE Товары.Цена <= [Введите предельную стоимость:]
            ORDER BY Товары.Цена;"""),

        ("Запрос_Контакты_По_Фирме",
         """PARAMETERS [Введите код фирмы:] Long;
            SELECT [Контактные лица].*
            FROM [Контактные лица]
            WHERE [Контактные лица].[Код фирмы] = [Введите код фирмы:];"""),

        ("Запрос_Остатки_На_Складах",
         """SELECT Склады.[Номер склада], Склады.Адрес, Склады.Телефон, Товары.Наименование, Товары.Марка,
                   Товары.Количество AS [Начальный запас],
                   Nz(DSum("Количество","Продажи","[Код товара]=" & Товары.[Код товара]), 0) AS [Продано],
                   Товары.Количество - Nz(DSum("Количество","Продажи","[Код товара]=" & Товары.[Код товара]), 0) AS [Остаток]
            FROM Склады INNER JOIN Товары ON Склады.[Номер склада] = Товары.[Номер склада];"""),

        ("Запрос_Средние_Цены",
         """SELECT Товары.Наименование, Count(Товары.[Код товара]) AS [Количество марок],
                   Avg(Товары.Цена) AS [Средняя цена], Min(Товары.Цена) AS [Мин цена], Max(Товары.Цена) AS [Макс цена]
            FROM Товары
            GROUP BY Товары.Наименование;"""),

        ("Запрос_Перекрестный_Продажи",
         """TRANSFORM Sum(Продажи.Количество) AS [Количество]
            SELECT Товары.Наименование
            FROM Фирмы INNER JOIN (Товары INNER JOIN Продажи ON Товары.[Код товара] = Продажи.[Код товара]) ON Фирмы.[Код фирмы] = Продажи.[Код фирмы]
            GROUP BY Товары.Наименование
            PIVOT Фирмы.[Название фирмы];"""),

        ("Запрос_Создание_Товары_1",
         """SELECT Товары.* INTO [Товары 1] FROM Товары;"""),

        ("Запрос_Обновление_Скидка_10",
         """UPDATE [Товары 1] SET [Товары 1].Цена = [Товары 1].Цена * 0.9;"""),

        ("Запрос_Создание_Видеокамеры",
         """SELECT [Товары 1].* INTO Видеокамеры FROM [Товары 1] WHERE Наименование='Видеокамера';"""),

        ("Запрос_Добавление_Видеомагнитофонов",
         """INSERT INTO Видеокамеры SELECT [Товары 1].* FROM [Товары 1] WHERE Наименование='Видеомагнитофон';"""),

        ("Запрос_Удаление_Sony_E150EE",
         """DELETE FROM Видеотовары WHERE Марка LIKE '*Sony*E150EE*' OR Марка LIKE '*Sony*';""")
    ]

    for q_name, q_sql in queries:
        try:
            qdef = db.CreateQueryDef(q_name, q_sql)
            print(f"Query '{q_name}' created.")
        except Exception as e:
            print(f"Error creating query '{q_name}':", e)

    # 5. Execute Action Queries to physically create Товары 1, Видеокамеры -> rename to Видеотовары
    print("Executing action queries for Task 4...")
    try:
        db.Execute("SELECT Товары.* INTO [Товары 1] FROM Товары;")
        print("Table [Товары 1] created.")
        db.Execute("UPDATE [Товары 1] SET [Товары 1].Цена = [Товары 1].Цена * 0.9;")
        print("Table [Товары 1] prices updated with 10% discount.")
        db.Execute("SELECT [Товары 1].* INTO Видеокамеры FROM [Товары 1] WHERE Наименование='Видеокамера';")
        print("Table [Видеокамеры] created.")
        db.Execute("INSERT INTO Видеокамеры SELECT [Товары 1].* FROM [Товары 1] WHERE Наименование='Видеомагнитофон';")
        print("Table [Видеокамеры] appended with video recorders.")
        
        # Rename Видеокамеры to Видеотовары via DoCmd
        access.DoCmd.Rename("Видеотовары", 0, "Видеокамеры") # 0 = acTable
        print("Table renamed to [Видеотовары].")
        
        # Insert Sony E150EE to demonstrate deletion
        db.Execute("INSERT INTO [Видеотовары] ([Наименование], [Марка], [Цена], [Номер склада], [Количество], [Описание]) VALUES ('Видеомагнитофон', 'Sony E150EE', 380, 20, 4, '4-головочный видеомагнитофон');")
        # Run delete
        db.Execute("DELETE FROM [Видеотовары] WHERE Марка='Sony E150EE';")
        print("Sony E150EE deleted from [Видеотовары].")
    except Exception as e:
        print("Action query execution error:", e)

    access.CloseCurrentDatabase()
    access.Quit()
    print("Database build complete and closed successfully!")

if __name__ == '__main__':
    create_database()
