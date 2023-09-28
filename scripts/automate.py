import xlwings as xw
import ctypes
import pandas as pd
import socket
import util

import excel2pdf
from excel2pdf import DocumentMaker
import convert2jpeg
import bill
import specsheets


debug_filename = r"D:\Project\FL\ImageConverter\log.txt"
debug_config = {"EnvHome": r"C:\Users\vlad.BANG.000\.virtualenvs\ImageConverter",
                "ProjectsFolder": r"D:\Project\FL\ImageConverter\data\projects",
                "TemplatesFolder": r"D:\Project\FL\ImageConverter\templates",
                "AutomateFolder": r"D:\Project\FL\ImageConverter",
                "SpecSheetsFolder": r"D:\Project\FL\ImageConverter\data\Specs Sheets"}

project_sheet_name = 'Projects'


class AutomateError(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
        return f"Ошибка: {self.message}"


def beep():
    import winsound
    frequency = 2500  # Set Frequency To 2500 Hertz
    duration = 1000  # Set Duration To 1000 ms == 1 second
    winsound.Beep(frequency, duration)


def log(line):
    with open(debug_filename, 'w') as file:
        file.write(line + "\n")


MB_OK = 0x0
MB_OKCXL = 0x01
MB_YESNOCXL = 0x03
MB_YESNO = 0x04
MB_HELP = 0x4000
ICON_EXLAIM=0x30
ICON_INFO = 0x40
ICON_STOP = 0x10

WS_EX_TOPMOST = 0x40000


def msg_box(msg, title='Automate', arg=MB_OK | ICON_STOP | WS_EX_TOPMOST) -> None:
    mb = ctypes.windll.user32.MessageBoxW
    mb(None, msg, title, arg)


def read_config(wb):
    hostname = socket.gethostname()
    if hostname == 'bang':
        debug_config['Database'] = wb.fullname
        return debug_config

    config_sheet_name = 'Config'
    try:
        # sheet = wb.sheets(config_sheet_name)
        sheet_value = wb.sheets[config_sheet_name].used_range.value
        df = pd.DataFrame(sheet_value)
        config = dict()
        config['Database'] = wb.fullname
        for index, row in df.iterrows():
            config[row[0]] = row[1]
        return config
    except Exception as e:
        return False


def get_used_range(sheet: xw.Sheet):
    return {"used_range_rows": (sheet.api.UsedRange.Row, sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count - 1),
            "used_range_cols": (sheet.api.UsedRange.Column, sheet.api.UsedRange.Column + sheet.api.UsedRange.Columns.Count - 1)}


def get_column_index(sheet: xw.Sheet, col_name):
    col = 1
    while col < get_used_range(sheet)["used_range_cols"][1]:
        if sheet.cells(1, col).value == col_name:
            return col
        col += 1
    raise AutomateError(f"Не найден индекс колонки '{col_name}'")


# @xw.func
# def python_from_vba(name):
#     return f"Hello {name}"


def vba_python_vba_macros(arg1, arg2):
    """Test: вызов этой функции из VBA, и затем вызов из VBA макроса с параметрами из питона"""
    msg_box(f"arg1: {arg1}, arg2: {arg2}, mult: {arg1*arg2}")
    wb = xw.Book.caller()
    app = wb.app
    macro_vba = app.macro("Module_Aux.auxOpenNewFileDialog")
    sub_arg = "test_filename.xlsx"
    macro_vba(sub_arg)


def call_vba_macro(macro_name, macro_args, excel_filename=''):
    wb = xw.Book(excel_filename) if excel_filename else xw.Book.caller()
    app = wb.app
    macro_vba = app.macro(macro_name)
    macro_vba(macro_args)


# def vba_from_python():
#     automate_addin_file = r"D:\Project\FL\ImageConverter\automate.xlsm"
#     wb = Book.caller(automate_addin_file)
#     macro = wb.macro("Module1.ExampleSub")
#     macro(sub_arg)
#     res = dm.create_pdf(project_rel_path, project_name, template_type)
#     # wb.save()
#     wb.close()


def read_book(excel_filename=''):
    wb = xw.Book(excel_filename) if excel_filename else xw.Book.caller()
    wb.save()

    sheet = wb.selection.sheet
    if sheet.name != project_sheet_name:
        raise AutomateError(f"Лист {project_sheet_name} не является активным")

    config = read_config(wb)
    if not config:
        raise AutomateError(f"Не удалось прочитать лист 'Config'")

    cell_range = wb.app.selection
    project = {}
    project['ProjectPath'] = wb.sheets(project_sheet_name).cells(cell_range.row, get_column_index(wb.sheets(project_sheet_name), "ProjectPath")).value
    project['ProjectName'] = wb.sheets(project_sheet_name).cells(cell_range.row, get_column_index(wb.sheets(project_sheet_name), "ProjectName")).value
    # project['...'] = ...

    # like $G$3
    # sel_address = cell_range.get_address()
    # sel_value = wb.sheets(project_sheet_name).cells(cell_range.row, cell_range.column).value
    # field_key = wb.sheets(project_sheet_name).cells(1, cell_range.column).value
    # log(f"{field_key}={sel_value}")
    # wb.close()
    return config, project


def vba_create_pdf(template_type):
    try:
        config, project = read_book()
        dm = DocumentMaker(db_path=config['Database'], template_folder=config['TemplatesFolder'], projects_folder=config['ProjectsFolder'])
        doc_path = ''
        if project['ProjectPath']:
            doc_path = dm.create_pdf(project['ProjectPath'], project['ProjectName'], template_type)
        if doc_path:
            call_vba_macro("Module_Aux.auxOpenNewFileDialog", doc_path)
    except excel2pdf.DocumentMakerError as e:
        msg_box(str(e))
    except AutomateError as e:
        msg_box(str(e))
    except Exception as e:
        msg_box(str(e))


def vba_convert_photos(start_quality, step_quality):
    try:
        config, project = read_book()
        if not project['ProjectPath']:
            return
        source_folder = util.get_fullpath(config["ProjectsFolder"], project['ProjectPath']) + "\\Photos"
        target_folder = source_folder + "\\" + "JPEG"
        zip_path = convert2jpeg.convert(source_folder, target_folder, max_size=5 * 1024 * 1024, quality_threshold=0,
                                      quality_step=step_quality,
                                      start_quality=start_quality, clear_target_folder=True, close_logwnd=False)
        # if zip_path:
        #     call_vba_macro("Module_Aux.auxOpenNewFileDialog", zip_path)
    except convert2jpeg.ConverterError as e:
        msg_box(str(e))
    except AutomateError as e:
        msg_box(str(e))
    except Exception as e:
        msg_box(str(e))


def vba_process_bill(file_paths_with_delim, db_row):
    # msg_box(file_path)
    # msg_box(str(db_row))
    try:
        file_paths = file_paths_with_delim.split('[_delimiter_]')
        res = bill.read_items(file_paths)
        call_vba_macro("Module_Db.dbSetProjectRowValueArr", (db_row, 'AccName', res['AccName'], False))
        call_vba_macro("Module_Db.dbSetProjectRowValueArr", (db_row, 'AccNumber', res['AccNumber'], False))
        call_vba_macro("Module_Db.dbSetProjectRowValueArr", (db_row, 'Addr', res['Addr'], False))
    except bill.BillError as e:
        msg_box(str(e))
    except Exception as e:
        msg_box(str(e))


def vba_process_specsheets(args_with_delim, zip_foldername, root_folder, specsheets_folder):
    # msg_box(args_with_delim)
    # msg_box(root_folder)
    # msg_box(specsheets_folder)
    """
    :args_with_delim: массив строк в одной строке с разделителем [_delimiter_]
    :root_folder: папка, в которой создается подпапка "Spec Sheets"
    """
    try:
        models = args_with_delim.split('[_delimiter_]')
        res_path = specsheets.process_specsheets(models, zip_foldername, root_folder, specsheets_folder, max_size=5 * 1024 * 1024, close_logwnd=False)
        if res_path:
            call_vba_macro("Module_Aux.auxOpenNewFileDialog", res_path)
    except specsheets.SpecSheetsError as e:
        msg_box(str(e))
    except Exception as e:
        msg_box(str(e))


def vba_set_project_defaults(db_row):
    try:
        config, project = read_book()
        # создаем объект DocumentMaker, чтобы прочитать config.json (т.к. это реализовано в DocumentMaker)
        dm = DocumentMaker(db_path=config['Database'], template_folder=config['TemplatesFolder'],
                           projects_folder=config['ProjectsFolder'])
        # читаем умолчания из config.json и пишем их в ячейки БД
        db_fields = ['ContractorFirm', 'ContractorName', 'ContractorPhone', 'ContractorEmail', 'ContractorTitle']
        for db_field in db_fields:
            config_item = dm.get_config_item(db_field)
            call_vba_macro("Module_Db.dbSetProjectRowValueArr", (db_row, db_field, config_item['default'], True))
    except Exception as e:
        msg_box(str(e))

