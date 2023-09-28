"""pip install pdfminer.six"""
from pdfminer.high_level import extract_text
import re


class BillError(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
        return f"Ошибка: {self.message}"


def read_items(file_paths):
    # попытки чтения файлов
    # если успешное чтение очередного файла - чтение прекращается
    success = False
    for file_path in file_paths:
        text = extract_text(file_path, page_numbers=[0])
        # print(text)
        text_list = text.split('\n')
        print(text_list)
        items = dict()
        service = False
        found_service = False
        for index, item in enumerate(text_list):
            if 'ACCOUNT NUMBER' in item:
                items['AccNumber'] = text_list[index + 1].lstrip().rstrip()

            if not found_service:
                service = item
                # поиск в строке последовательности "запятая, любое кол-во пробелов, CA"
                res = re.search(r', *CA*', service)
                if res:
                    # нам нужно только первое вхождение ', CA'
                    found_service = True
                    # print("FOUND !!!")
                    # добавляем в строку service строки из списка text_list справа и слева
                    # пока не встретится пустая строка в списке text_list
                    for direction in [-1, 1]:
                        # -1: search backward
                        # 1: search forward
                        for i in range(1, 1000):
                            if text_list[index + direction * i] != '':
                                if direction:
                                    # print('backward')
                                    service = text_list[index + direction * i] + service
                                else:
                                    # print('forward')
                                    service += text_list[index + direction * i]
                            else:
                                break

        # address string examples:
        # service = 'ROBERT J LOWE,  16133 VENTURA BLVD, ENCINO, CA 91436'
        # service = 'ABC, Inc,   16133 VENTURA BLVD, ENCINO, CA 91436'

        if found_service:
            # print(service)
            # поиск в строке последовательности "запятая, любое кол-во пробелов, любая цифра, любое кол-во любых символов"
            res = re.search(r', *\d[\s\S]*', service)
            # поиск подстроки, начинающейся с любой цифры
            if res:
                digit = re.search(r'\d', res.group(0))
                items['Addr'] = res.group(0)[digit.span()[0]:].lstrip().rstrip()
                items['AccName'] = service[:res.span()[0]].lstrip().rstrip()
            success = True
            break
        else:
            continue

    if not success:
        text = ""
        for file_path in file_paths:
            text += f"\n{file_path}"
        raise BillError(f"Не удалось извлечь информацию из файла(ов):\n{text}")

    return items


if __name__ == '__main__':
    try:
        # file_path = r"D:\Project\FL\ImageConverter\data\projects\Hospitality At Work\Application\Bills\Bill.pdf"
        file_path = r"D:\Project\FL\ImageConverter\data\projects\Hospitality At Work\Application\Bills\Automate\Bill.pdf"
        res = read_items(file_path)
        print(res)

        file_path = r"D:\Project\FL\ImageConverter\data\projects\Tri Center Plaza\Application\Bills\Bill Feb.pdf"
        res = read_items(file_path)
        print(res)
    except BillError as e:
        print(e)
