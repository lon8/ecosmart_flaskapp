import glob
import os
import stat
import zipfile


def is_fullpath(path: str) -> bool:
    """определяет, является ли path полным путем (к файлу или папке) или относительным"""
    return ":" in path


def get_fullpath(folder_path: str, rel_path: str) -> str:
    """возвращает полный путь, составленный из folder_path и rel_path"""
    if not is_fullpath(rel_path):
        return folder_path + "\\" + rel_path
    else:
        return rel_path


def get_size(folder):
    return sum(os.path.getsize(folder + "\\" + f) for f in os.listdir(folder) if os.path.isfile(folder + "\\" + f))


def zip_folder(folder):
    """
    infolder: True - файлы в архиве остаются в папке, False - файлы архивируется без папки
    """
    try:
        zfilename = os.path.dirname(folder) + "\\" + os.path.basename(folder) + ".zip"
        if os.path.exists(zfilename):
            os.remove(zfilename)
        newzip = zipfile.ZipFile(zfilename, 'w')

        # старый вариант: без субдиректорий, и не работает (вылетает), когда в директории есть субдиректории
        # for file in os.listdir(folder):
        #     newzip.write(folder + "\\" + file, arcname=file)

        # новый вариант: с субдиректориями
        filenames = glob.glob(f"{folder}/**/*.*", recursive=True)
        for file in filenames:
            newzip.write(file, arcname=os.path.relpath(file, folder))

        newzip.close()
        return zfilename
    except Exception as e:
        raise Exception("Ошибка при сжатии папки с файлами")
    return ''


def remove_readonly(func, path, _):
    """
    Удаление дерева каталогов в Windows, где для некоторых файлов установлен бит только для чтения.
    Используется обратный вызов onerror, чтобы очистить бит readonly и повторить попытку удаления.

    Пример использования:
    shutil.rmtree(directory, onerror=remove_readonly)
    """
    os.chmod(path, stat.S_IWRITE)
    func(path)

""" 
Поиск файлов в подпапках
 
root_path/     the dir
**/       every file and dir under my_path
*.txt     every file that ends with '.txt'

files = glob.glob(root_path + '/**/*.txt', recursive=True)

"""

""" 
Поиск файлов в подпапках итератором

for file in glob.iglob(my_path, recursive=False):
    ...

"""