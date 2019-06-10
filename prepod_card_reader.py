import sqlite3
import pip
import datetime
import xlrd

# def install(package):
#     pip.main(['install', package])
#
#
# try:
#     import xlrd
# except ImportError:
#     install("xlrd")


class Reader:
    """"""
    def __init__(self, paths):
        self.paths = paths
        self.path_to_db = "prepod_timetable_base.db"
        self.connect = sqlite3.connect(self.path_to_db)
        self.cursor = self.connect.cursor()

    def create_table(self, table_name):
        """Если Таблица не создана, создать таблицу
            arg1 - Название таблицы
        """
        self.cursor.execute("DROP TABLE IF EXISTS {}".format(table_name))
        self.cursor.execute("""CREATE TABLE {} (date TEXT, time TEXT, room TEXT)""".format(table_name))

    def data_append(self, table_name, date, time, room):
        """Добавление данных в базу данных"""
        self.cursor.execute("""INSERT INTO {} ('date', 'time', 'room') VALUES (?,?,?)""".format(table_name),
                            (date, time, room))

    def run(self):
        def date_convertor(excel_date):
            if isinstance(excel_date, (float, int)):
                date_tuple = xlrd.xldate_as_tuple(excel_date, 0)
                date = datetime.date(year=date_tuple[0], month=date_tuple[1], day=date_tuple[2])
                date_str = date.strftime("%d.%m.%Y")
                return date_str
            else:
                return None

        def time_convertor(excel_date):
            if isinstance(excel_date, (float, int)):
                time_tuple = xlrd.xldate_as_tuple(excel_date, 0)
                date = datetime.time(hour=time_tuple[3], minute=time_tuple[4])
                time_str = date.strftime("%H:%M")
                return time_str
            else:
                return None

        for path_to_file in self.paths:

            book = xlrd.open_workbook(path_to_file)
            sheets = book.sheet_by_index(0)

            f_i_o = str(sheets.cell(6, 4).value.replace(" ", "") + "_" + sheets.cell(7, 4).value.replace(" ", "") + "_" + sheets.cell(8, 4).value.replace(" ", ""))
            print(f_i_o)
            self.create_table(f_i_o)

            position = [1, 14]
            time_position = [1, 13]

            for week in range(1, 18 + 1):
                for day in range(1, 6 + 1):
                    # Записываем дату каждого дня
                    day_date = date_convertor(sheets.cell(position[0] - 1, position[1]).value)

                    for para in range(1, 8 + 1):

                        time = time_convertor(sheets.cell(time_position[0], time_position[1]).value)

                        try:
                            room = int(sheets.cell(position[0], position[1] + 1).value)
                        except ValueError:
                            room = sheets.cell(position[0], position[1] + 1).value

                        if sheets.cell(position[0], position[1] + 1).value:
                            self.data_append(f_i_o, day_date, time, room)
                        if para == 8:
                            position[0] = position[0] + 2
                            time_position[0] = time_position[0] + 2
                        else:
                            position[0] = position[0] + 1
                            time_position[0] = time_position[0] + 1

                position[1] = position[1] + 2
                position[0] = 1
                time_position[0] = 1

        self.connect.commit()
        self.cursor.close()


if __name__ == "__main__":
    reader = Reader(["Polozyuk_Alexey_Grigoryevich.xlsx", "Kolchin_Andrey_Igorevich.xlsx"])
    reader.run()
