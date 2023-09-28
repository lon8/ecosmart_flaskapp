import datetime
import os
import shutil
import glob
import re

import automate
import pdf_compressor as pdfc
import time
import util
import logui


class SpecSheetsError(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
        return f"Ошибка: {self.message}"


"""
'Spec Sheets' и 'Specs Sheets' - разные папки !!!
'Spec Sheets' - папка в проекте
'Specs Sheets' - отдельно лежащая папка со всеми спецификациями
"""


def copy_by_mask(mask, specsheets_folder, target_folder):
    # поиск в папке (и подпапках) specsheets_folder=config['Specs Sheets'] файлов совпадающих с заданной маской
    filenames = glob.glob(f"{specsheets_folder}/**/{mask}", recursive=True)
    # копируем найденные файлы в папку target_folder ("Spec Sheets")
    target_filepaths = []
    for filename in filenames:
        if not os.path.isfile(filename):
            continue
        basename = os.path.basename(filename)
        target_filepath = target_folder + "\\" + basename
        shutil.copy(filename, target_filepath)
        target_filepaths.append(target_filepath)
    return target_filepaths


def process_specsheets(models, zip_foldername, root_folder, specsheets_folder, max_size, close_logwnd=True):
    if not os.path.exists(specsheets_folder):
        raise SpecSheetsError(f"Не найден путь к папке 'Specs Sheets': '{specsheets_folder}'")
    # создаем папку Spec Sheets
    target_folder = root_folder + "\\" + "Spec Sheets"
    # if not os.path.exists(target_folder): # у Антона так не работает, нет доступа на диске C:
    try:
        os.mkdir(target_folder)
    except:
        pass

    target_filepaths = []

    for model in models:
        # меняем в моделях слэши на пробелы
        model = model.replace("/", " ")

        # поиск в папке (и подпапках) specsheets_folder=config['Specs Sheets'] файлов совпадающих с заданной маской
        copied_filepaths = copy_by_mask(f"*{model}*", specsheets_folder, target_folder)
        target_filepaths += copied_filepaths

        # если скопировано менее, чем 2 файла
        if len(copied_filepaths) < 2:
            # находим в названии модели подстроку " ### " заменяем на " #xx "
            res = re.search(r'\D\d\d\d\D', model)

            if res:
                model = model[:res.start() + 2] + "xx" + model[res.end() - 1:]
                # повторно копируем с измененной маской
                target_filepaths += copy_by_mask(f"*{model}*", specsheets_folder, target_folder)

        # удаляем дубликаты в списке скопированных файлов
        target_filepaths = list(set(target_filepaths))

    # сжимаем

    # target_folder_pdf = target_folder + "\\PDF"
    target_folder_pdf = target_folder + f"\\{zip_foldername} - Spec Sheets - {datetime.datetime.now().strftime('%Y.%m.%d')}"

    quality = 0

    stop_processing = False  # для того, чтобы окно лога оставалось висеть после выполнения работы
    logwnd = logui.Logui('Сжатие и архивирование файлов')
    while True:
        if not logwnd.proceed():
            break

        if not stop_processing and logwnd.can_processing():
            logwnd.perform_long_operation(
                lambda: compress(logwnd, target_filepaths, target_folder_pdf, quality))

        if not stop_processing and logwnd.can_after_processing():
            size = util.get_size(target_folder_pdf)

            print(f"Общий размер результирующей папки с файлами: {size / 1024 / 1024:.2f} Мб")
            time.sleep(.01)
            if size <= max_size:    # success
                stop_processing = True
                # break # закоментировал, чтобы не закрывалось окно после выполнения функции

                '''
                Закончили сжатие, архивируем
                '''

                # # удаляем исходные файлы несжатые файлы
                # for file in target_filepaths:
                #     try:
                #         os.remove(file)
                #     except OSError:
                #         pass

                try:
                    # удаляем старый zip, лежащий внутри папки target_folder (Spec Sheets)
                    res_zfilename = target_folder + "\\" + os.path.basename(target_folder) + ".zip"
                    os.remove(res_zfilename)
                    # удаляем старую папку PDF
                    # os.remove(target_folder_pdf)
                except OSError:
                    pass

                # копируем в папку target_folder_pdf остальные (не-PDF) файлы
                for filepath in target_filepaths:
                    if os.path.splitext(filepath)[1] != '.pdf':
                        shutil.copy(filepath, target_folder_pdf)

                # архивируем
                # zfilename = util.zip_folder(target_folder)
                zfilename = util.zip_folder(target_folder_pdf)

                # переносим файл архива внутрь папки target_folder (т.к. после архивации они лежат на одном уровне)
                # shutil.move(zfilename, res_zfilename)

                try:
                    # удаляем старую папку PDF
                    # os.remove(target_folder_pdf)
                    shutil.rmtree(target_folder_pdf, onerror=util.remove_readonly)
                except OSError:
                    pass

            quality += 1

            if quality > 4:
                stop_processing = True
                if close_logwnd:
                    logwnd.close()
                print("Процесс остановлен: достигнут максимальный порог сжатия")
                # break # закоментировал, чтобы не закрывалось окно после выполнения функции

            logwnd.complete_after_processing()


    if close_logwnd:
        logwnd.close()

    # return res_zfilename
    return target_folder


def compress(logwnd, source_filepaths, target_folder, quality):
    # сжимаем файлы
    print(f"\nЗапуск сжатия PDF-файлов с уровнем сжатия: {quality}\n")
    time.sleep(.01)
    # чистим папку target
    if os.path.exists(target_folder):
        shutil.rmtree(target_folder, ignore_errors=True)
    # создаем папку target
    if not os.path.exists(target_folder):
        os.mkdir(target_folder)
    source_count = 0
    target_count = 0
    for source_filepath in source_filepaths:
        try:
            source_count += 1
            target_filepath = target_folder + "\\" + os.path.basename(source_filepath)
            pdfc.compress(source_filepath, target_filepath, quality)
            target_count += 1
            print(f"Создан файл: {target_filepath}")
            time.sleep(.01)
            pass
        except FileNotFoundError as e:
            print(str(e))
        except Exception as e:
            print(f"Ошибка при сжатии файла: {target_filepath}")

    print(f"\nСжатие выполнено для {target_count} из {source_count} файлов")
    time.sleep(.01)
    logwnd.complete_processing()


if __name__ == '__main__':
    specsheets_folder = r"D:\Project\FL\ImageConverter\data\Specs Sheets"
    root_folder = r"D:\Project\FL\ImageConverter\data\projects\Yale Management Services"
    # models = ['LT40W/840-ID', 'L48T8 840 10G-ID DE', 'AGL']
    models = ['LT40W/840-ID']
    result = process_specsheets(models, 'zip-name', root_folder, specsheets_folder, max_size=5 * 1024 * 1024, close_logwnd=False)
