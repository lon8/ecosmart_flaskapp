import os
import shutil
import time

from wand.image import Image

from app_log import debug
from . import util


class ConverterError(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
        return f"Ошибка: {self.message}"


'''
Обязательный параметр:
    source_folder: путь к папке с файлами

Необязательные параметры:
    clear_target_folder: True - удалить результаты предыдущей конвертации, False - не удалять     
'''


def convert_to_jpeg(
        source_folder,
        target_folder,
        quality=100,
        clear_target_folder=True
) -> bool:
    try:
        debug(f"Start convert with quality: {quality} %\n")
        time.sleep(.01)
        # проверяем корректность пути исходной папки
        if not os.path.exists(source_folder):
            debug(f"No found dir {source_folder}")
            raise ConverterError(f"Не найден указанный путь\n\n{source_folder}")
        # чистим папку target
        if clear_target_folder and os.path.exists(target_folder):
            shutil.rmtree(target_folder, ignore_errors=True)
        # создаем папку target
        if not os.path.exists(target_folder):
            os.mkdir(target_folder)
        # читаем файлы исходной папки
        source_count = 0
        target_count = 0
        for file in os.listdir(source_folder):
            # очередной файл
            source_file = source_folder + "\\" + file
            try:
                # проверяем, что не папка
                if os.path.isdir(source_file):
                    continue
                source_count += 1
                # конвертируем
                img = Image(filename=source_file)
                img.format = 'jpg'
                if quality < 100:
                    img.compression_quality = quality
                filename = os.path.basename(source_file)
                file_name, file_extension = os.path.splitext(source_file)
                target_file = target_folder + "\\" + filename.replace(file_extension, ".jpeg")
                # если файл существует, делаем добавку к его имени
                #while True:
                    #if os.path.exists(target_file):
                            #t_name, t_extension = os.path.splitext(target_file)
                            #t_name += "1"
                            #target_file = t_name + t_extension
                    #else:
                        #break
                img.save(filename=target_file)
                img.close()
                target_count += 1
                debug(f"Create file: {target_file}")
                time.sleep(.01)
            except Exception as e:
                debug(f"Error: {source_file}: {e}")

        debug(f"Convert finally {target_count} or {source_count} files")
        time.sleep(.01)
        return True
    except Exception as e:
        debug(f"Error: {e}")


def convert(
        source_folder,
        target_folder,
        max_size=5 * 1024 * 1024,
        start_quality=75,
        quality_threshold=0,
        quality_step=5,
        clear_target_folder=True
) -> str:
    """Процесс сжатия фотографий по заданному пути."""
    try:
        quality = start_quality

        while True:
            convert_to_jpeg(
                source_folder,
                target_folder,
                quality,
                clear_target_folder
            )
            size = util.get_size(target_folder)
            debug(f"All size result dir: {size / 1024 / 1024:.2f} m")
            time.sleep(.01)
            if size <= max_size:
                break
            quality -= quality_step

            if quality < quality_threshold:
                debug("Process stop. Very small quality images.")
                raise ConverterError("Процесс остановлен: получается слишком низкое качество изображений")

        # архивируем
        zip_path = util.zip_folder(target_folder)
        return zip_path
    except Exception as ex:
        debug("Error convert")
        debug(ex)


"""
Для локального тестирования
"""
# if __name__ == '__main__':
#     try:
        # source_folder = r"D:\Project\FL\ImageConverter\data\projects\prj3\Photos"
        # target_folder = r"D:\Project\FL\ImageConverter\data\projects\prj3\Photos\JPEG"
        # source_folder = r"D:\Project\FL\ImageConverter\data\projects\Hospitality At Work\Photos"
        # target_folder = r"D:\Project\FL\ImageConverter\data\projects\Hospitality At Work\Photos\JPEG"
        # result = convert(source_folder, target_folder, max_size=5 * 1024 * 1024, quality_threshold=0, quality_step=15,
        #                  start_quality=40, clear_target_folder=True)

        # """folder with photos"""
        # source_folder = sys.arg[1]
        # target_folder = source_folder + "\\" + "JPEG"
        # result = convert(source_folder, target_folder, max_size=5 * 1024 * 1024, quality_threshold=0, quality_step=15,
        #                  start_quality=75, clear_target_folder=True)

    # except ConverterError as e:
    #     print(str(e))
    # except logui.LoguiError as e:
    #     print(str(e))
    # except Exception as e:
    #     print(str(e))
