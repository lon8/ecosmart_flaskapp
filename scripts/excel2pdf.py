import os
import pandas
from fillpdf import fillpdfs
import fitz
import json
import datetime
import re
import socket
import util
import string
import shutil
import sys
import openpyxl
import automate
from tkinter import messagebox


class DocumentMakerError(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
        return f"Ошибка: {self.message}"


class DocumentMaker:
    def __init__(self, db_path, template_folder, projects_folder, project_sheet_name="Projects"):
        self.db_path = db_path
        self.template_folder = template_folder
        self.projects_folder = projects_folder
        self.project_sheet_name = project_sheet_name

        self.template_types = ['disclaimer', 'disclaimer_rev9222', 'application', 'excel777']

        self.template_type_folders = {
            "disclaimer": "RV Disclaimer",
            "disclaimer_rev9222": "RV Disclaimer",
            "application": "Application",
            "excel777": "CLIP"}

        self.template_filenames = {
            "disclaimer": "template_ladwp_disclaimer.pdf",
            "disclaimer_rev9222": "RV Disclaimer (Rev 9-2-22).pdf",
            "application": "template_ladwp_application.pdf"
        }

    def get_template_path(self, template_type):
        path = self.template_folder + "\\" + self.template_filenames[template_type]
        if not os.path.exists(path):
            raise DocumentMakerError(f"Не найден файл '{path}'")
        return path

    def get_config_path(self):
        path = self.template_folder + "\\" + "config.json"
        if not os.path.exists(path):
            raise DocumentMakerError(f"Не найден файл '{path}'")
        return path

    def create_project_folders(self, project_path):
        for template_type, folder in self.template_type_folders.items():
            path = project_path + "\\" + folder
            if not os.path.exists(path):
                os.mkdir(path)

    def read_config(self):
        try:
            file = open(self.get_config_path(), 'r')
            content = file.read()
            config_list = json.loads(content)
            file.close()
            return config_list
        except Exception as e:
            raise DocumentMakerError(f"Не удалось загрузить файл настроек шаблонов '{self.get_config_path()}'")

    def get_config_item(self, db_field, config_list=None):
        if config_list is None:
            config_list = self.read_config()
        for item in config_list:
            if item['db'] == db_field:
                return item
        return False

    def print_config_statistic(self):
        config_list = self.read_config()
        not_used_count = 0
        all_count = 0
        default_count = 0
        empty_db_count = 0
        # union_disclaimer_application_count = 0
        special_rule_count = 0
        for item in config_list:
            all_count += 1
            res = re.search("not used", item['comment'])
            if res is not None:
                not_used_count += 1
            else:
                if item['default']:
                    default_count += 1
                if not item['db']:
                    empty_db_count += 1
                # if item['disclaimer'] and \
                #         item['application']:
                #     union_disclaimer_application_count += 1
                for ttype in self.template_types:
                    if "SPECIAL_RULE" in item[ttype]:
                        special_rule_count += 1
        print(f"Конфигурационный файл шаблонов '{self.get_config_path()}':")
        print(f"\t[-] всего записей: {all_count}")
        print(f"\t[-] используется записей: {all_count - not_used_count}")
        print(f"\t\t[-] задано умолчаний: {default_count}")
        print(f"\t\t[-] не задано названий колонок базы данных: {empty_db_count}")
        # print(f"\t\t[-] общих полей шаблонов 'disclaimer' и 'application': {union_disclaimer_application_count}")
        print(f"\t\t[-] специальных правил: {special_rule_count}")

    def get_pdf_fields(self, template_type):
        data = fillpdfs.get_form_fields(self.get_template_path(template_type))
        print(f"В документе '{self.get_template_path(template_type)}' найдено {len(data)} полей")
        print(data)

    def get_db_field_value(self, known_field_key, known_field_value, field_key):
        '''
        Находит первое, подходящее под запрашиваемые условия, вхождение в БД
        '''
        try:
            project_df = pandas.read_excel(self.db_path, sheet_name=self.project_sheet_name)
            return project_df.loc[project_df[known_field_key] == known_field_value][field_key].iat[0]
        except Exception as e:
            raise DocumentMakerError(f"Не удалось найти запрашиваемый код '{field_key}' в базе данных '{self.db_path}'")

    def create_pdf(self, project_rel_path, project_name, template_type):
        # создаем папку проекта, если ее еще нет
        project_path = util.get_fullpath(self.projects_folder, project_rel_path)
        if not os.path.exists(project_path):
            os.mkdir(project_path)

        # создаем структуру папок проекта
        self.create_project_folders(project_path)

        msg = ''
        doc_path = ''
        try:
            # имя создаваемого документа
            # name = self.get_db_field_value("ProjectPath", project_rel_path, "Addr").split(',')[0] + " - " + \
            name = project_name + " - " + self.template_type_folders[template_type] + " - " + datetime.datetime.now().strftime("%Y.%m.%d")
            doc_path = project_path + "\\" + self.template_type_folders[template_type] + "\\" + name + ".pdf"
            # читаем файл настроек шаблонов
            config_list = self.read_config()
            # print(config_list)
            # читаем файл базы данных
            project_df = pandas.read_excel(self.db_path, engine='openpyxl', sheet_name=self.project_sheet_name)
            # project_df = project_df.astype(str)
            series = project_df.loc[(project_df['ProjectName'] == project_name) & (project_df['ProjectPath'] == project_rel_path)]
            if series.empty:
                raise DocumentMakerError(f"В базе данных не найдена ссылка на проект:\n\n'{project_name}' по пути\n\n:'{project_rel_path}'")
            # заполняем документ значениями из базы данных
            # data_dict - словарь полей документа
            data_dict = dict()
            # колонки БД, исключаемые из проверки на их наличие (т.к. они не заполняются данной функцией)
            exclude_db_columns = ['ProjectPath', 'ProjectName', 'Client', 'Rebate']
            # index - имя колонки БД
            for index, value in series.items():
                # val - значение поля колонки index БД
                val = series[index].iat[0]
                # убираем NaN (пустые ячейки)
                if pandas.isna(val):
                    val = ''

                config_item = self.get_config_item(index, config_list)
                if config_item is not False:
                    # print(config_item)
                    # data_keys - список полей шаблона, соответствующих index колонке БД
                    data_keys = config_item[template_type]
                else:
                    if index not in exclude_db_columns:
                        raise DocumentMakerError(f"В файле настроек шаблонов '{self.get_config_path()}' не найдено поле '{index}'")
                    # 'ProjectPath' пропускаем
                    continue

                # если в БД не задано значение, то будем заполнять документ значением по умолчанию из config'a
                if not val:
                    val = config_item['default']

                # key - поле шаблона, соответствующего index колонке БД
                for key in data_keys:
                    # если значение колонки БД при записи в документ обрабатывается специальными правилами
                    if key == 'SPECIAL_RULE':
                        DocumentMaker.process_special_rules(data_dict, index, val, template_type)
                        # print(config_item)
                    # если значение записывается в поле документа как оно есть
                    else:
                        data_dict[key] = val
            # print(data_dict)
            # заполняем документ
            fillpdfs.write_fillable_pdf(self.get_template_path(template_type), doc_path, data_dict)
        except DocumentMakerError as e:
            raise e
        except Exception as e:
            raise DocumentMakerError(f"Не удалось заполнить документ '{doc_path}'")

        print(f"Создан документ '{doc_path}'")
        return doc_path

    def get_pdf_template_comments(self, template_type):
        print(f"Комментарии файла '{self.get_template_path(template_type)}':")
        comments = []
        document = fitz.open(self.get_template_path(template_type))
        for i in range(document.page_count):
            page = document[i]
            for annot in page.annots():
                comments.append(annot.info)
                print(annot.info)
        return comments

    @staticmethod
    def parse_phone(val):
        rule_format = "(310) 853-9722\n310.853.97722 ext 321\n1 310 853 9722\n1 (310) 853 9722\n'+1 (310) 853 9722\n1 310 853 9722 ext 321"
        # code = ''.join(re.findall(r'\((.*?)\)', val))
        # number = val.split(' ')[1]
        # removing non-digits from string
        digs = ''.join(filter(str.isdigit, val))
        if len(digs) == 10 or len(digs) == 13:
            code = digs[0:3]
            number = digs[3:6] + "-" + digs[6:10]
            if len(digs) == 13:
                number += ' ext ' + digs[10:13]
        elif len(digs) == 11 or len(digs) == 14:
            code = digs[1:4]
            number = digs[4:7] + "-" + digs[7:11]
            if len(digs) == 14:
                number += ' ext ' + digs[11:14]
        else:
            raise Exception('')
        # print(f"code: {code}, number: {number}")
        return code, number, rule_format

    @staticmethod
    def process_special_rules(data_dict, index, val, template_type):
        '''
        Обработка специальных правил: например, раскидывает значение поля БД по нескольким полям документа
        data_dict: словарь полей документа
        index: имя колонки БД
        val: значение поля колонки index БД
        template_type: тип шаблона документа
        '''

        if not val:
            return

        try:
            rule_format = ''
            if template_type == 'application':

                phone_indexes = ['ContactPhone', 'InspectionPhone', 'ContractorPhone']
                if index in phone_indexes:
                    code, number, rule_format = DocumentMaker.parse_phone(val)
                    if index == 'ContactPhone':
                        data_dict['Customer Contact Area Code'] = data_dict['_Customer Contact Area Code 5'] = code
                        data_dict['Customer Contact Phone Number'] = data_dict['_Customer Contact Phone Number 5'] = number
                    if index == 'InspectionPhone':
                        data_dict['Inspection Contact Area Code'] = code
                        data_dict['Inspection Contact Phone Number'] = number
                    if index == 'ContractorPhone':
                        data_dict['Contractor Contact Area Code'] = code
                        data_dict['Contractor Contact Phone Number'] = number

                if index == 'OwnerTenant':
                    rule_format = "значение ячейки БД должно быть равно значению поля документа (например, 'Owner')"
                    data_dict[val] = 'Yes'

                if index == 'BldType':
                    rule_format = "значение ячейки БД должно быть равно значению поля документа (например, 'Office Building')"
                    data_dict[val] = 'Yes'

                address_indexes = ['Addr', 'AddrMailing']
                if index in address_indexes:
                    zip_code = val.split(' ')[-1]
                    state = val.split(' ')[-2]
                    city = val.split(',')[-2]
                    pos = val.rfind(',')
                    pos = val.rfind(',', 0, pos - 1)
                    street_address = val[:pos]
                    if index == 'Addr':
                        # service address
                        rule_format = "16133 Ventura Blvd, Encino, CA 91436"
                        data_dict['Street Address'] = street_address
                        data_dict['City'] = city
                        data_dict['State'] = state
                        data_dict['Zip Code'] = zip_code
                        # installation address (the same as service address)
                        data_dict['_Installation Street Address'] = data_dict['_Installation Street Address 5'] = street_address
                        data_dict['_Installation City'] = data_dict['_Installation City 5'] = city
                        data_dict['_Installation State'] = data_dict['_Installation State 5'] = state
                        data_dict['_Installation Zip Code'] = data_dict['_Installation Zip Code 5'] = zip_code
                    if index == 'AddrMailing':
                        rule_format = "16133 Ventura Blvd, Suite 265, Encino, CA 91436"
                        data_dict['Street Address_2'] = street_address
                        data_dict['City 2'] = city
                        data_dict['State 2'] = state
                        data_dict['Zip Code 2'] = zip_code

                square_indexes = ['BldSqr', 'BldSqrCond']
                if index in square_indexes:
                    val = "{:,.0f}".format(val)
                    if index == 'BldSqr':
                        data_dict['Total facility square footage'] = data_dict['_Installation Total facility square footage'] = val
                    if index == 'BldSqrCond':
                        data_dict['Conditioned space square footage'] = val

                hours_indexes = ['Hours', 'Hours2']
                if index in hours_indexes:
                    day_val = float(val) / 365.
                    day_val = "{:.0f}".format(day_val)
                    hrs_wk_val = float(val) / 365. * 7
                    hrs_wk_val = "{:.0f}".format(hrs_wk_val)
                    week_yr_val = '52'
                    hrs_yr_val = "{:,.0f}".format(val)
                    if index == 'Hours':
                        data_dict['_OpShed_HrsWk'] = hrs_wk_val
                        data_dict['_OpShed_WeekYr'] = week_yr_val
                        data_dict['_OpShed_HrsYr'] = hrs_yr_val
                        day_keys = ['_OpShed_M', '_OpShed_T', '_OpShed_W', '_OpShed_Th', '_OpShed_F', '_OpShed_S', '_OpShed_Sun']
                        for day_key in day_keys:
                            data_dict[day_key] = day_val
                    if index == 'Hours2':
                        data_dict['_OpShed_HrsWk 2'] = hrs_wk_val
                        data_dict['_OpShed_WeekYr 2'] = week_yr_val
                        data_dict['_OpShed_HrsYr 2'] = hrs_yr_val
                        day_keys = ['_OpShed_M 2', '_OpShed_T 2', '_OpShed_W 2', '_OpShed_Th 2', '_OpShed_F 2', '_OpShed_S 2', '_OpShed_Sun 2']
                        for day_key in day_keys:
                            data_dict[day_key] = day_val

                if index == 'ContactTitle':
                    data_dict['_Name of Customer Contact Person 5'] = data_dict['Name of Customer Contact Person'] + ", " + val
                    pass

        except DocumentMakerError as e:
            raise e
        except Exception as e:
            raise DocumentMakerError(f"Формат значения {val} ячейки поля БД '{index}' не соответствует заданному правилу:\n'{rule_format}'")


if __name__ == '__main__':
    # сетевое имя компа
    hostname = socket.gethostname()

    if hostname == 'Anton':
        db_file = r"C:\Work\Programming\automate\database.xlsx"
        proj_folder = r"C:\Work\Programming\automate\projects"
        tmpl_folder = r"C:\Work\Programming\automate\templates"
    if hostname == 'bang':
        db_file = r"D:\Project\FL\ImageConverter\database\database.xlsx"
        proj_folder = r"D:\Project\FL\ImageConverter\data\projects"
        tmpl_folder = r"D:\Project\FL\ImageConverter\templates"

    # ::::: ТЕСТ ::::: заполняем документы значениями из БД
    try:
        dm = DocumentMaker(db_path=db_file, template_folder=tmpl_folder, projects_folder=proj_folder)
        # result = dm.create_pdf("prj1", 'PrName Phase 1', 'disclaimer')
        # result = dm.create_pdf("prj2", 'PrName Part 2', 'disclaimer')
        result = dm.create_pdf("prj2", 'PrName Phase 2', 'disclaimer_rev9222')
        print()

        # result = dm.create_pdf("prj1", 'PrName Phase1', 'application')
        # result = dm.create_pdf("prj9", 'PrName', 'application')
        print()

        # ::::: ТЕСТ ::::: получаем значения из БД
        # val = dm.get_db_field_value("ProjectPath", "\\prj1", "Addr")
        # print(val)
        # val = dm.get_db_field_value("Addr", "16133 Ventura Blvd", "Phone Number 1")
        # print(val)
        # print()

        # ::::: ТЕСТ ::::: получаем поля документа PDF
        # dm.get_pdf_fields('disclaimer')
        # print()
        # dm.get_pdf_fields('application')
        # print()

        # ::::: ТЕСТ ::::: читаем комментарии в PDF
        # coms = dm.get_pdf_template_comments('application')
        # print()

        # ::::: ТЕСТ ::::: вывод статистики по config.json
        dm.print_config_statistic()
        print()
    except DocumentMakerError as e:
        print(e)
    except Exception as e:
        print(str(e))
