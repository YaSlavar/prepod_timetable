import os
import os.path
from prepod_card_reader import Reader
from timetable_writer import Timetable_writer


def run():
    def get_path_to_xlsx_files(source_path):
        paths = []
        for file in os.listdir(source_path):
            if os.path.isfile(os.path.join(source_path, file)):
                paths.append(os.path.join(source_path, file))
        return paths

    # path_to_folder = input("Введите путь к папке с карточками преподавателей: ")
    # path_to_template = input("Введите путь к шаблону: ")
    # name_out_file = input("Введите название готового файла: ")

    path_to_folder = "Y:\МЕДИА ФАЙЛЫ\ПО\prepod_timetable\card"
    path_to_template = "Y:\МЕДИА ФАЙЛЫ\ПО\prepod_timetable/template.xlsx"
    name_out_file = "График присутствия преподавателей во время зимней сессии.xlsx"
    from_date = "01.06.2019"
    to_date = "31.07.2019"

    xls_path_list = get_path_to_xlsx_files(path_to_folder)
    print(xls_path_list)
    dump = Reader(xls_path_list)

    dump.run()

    writer = Timetable_writer(path_to_template, name_out_file, from_date, to_date)
    writer.run()
    print("Готово!\n")


if __name__ == "__main__":
    run()
