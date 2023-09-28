import typing

from PyPDF2 import PdfReader, PdfWriter, PdfFileWriter
from PyPDF2.constants import TrailerKeys, CatalogDictionary
from PyPDF2.errors import PdfReadError
from PyPDF2.generic import IndirectObject, NameObject, BooleanObject, DictionaryObject
from pandas import read_excel
import argparse
import tkinter as tk
import tkinter.messagebox as msgbox
from tkinter import filedialog
from os.path import normpath, exists
from os import makedirs

from app_log import error
from utils import today_as_string


def redact(db_path=None, sheet_name=None, header_row=None, template_path=None, identifier=None):
    root = tk.Tk()
    root.withdraw()
    if template_path is None:
        template_path = normpath(filedialog.askopenfilename(title='Введите путь к Pdf шаблону',
                                                            filetypes=(('Pdf файлы', '*.pdf'), ('Все файлы', '*.*'))))
        if template_path == '.':
            return

    if identifier is None:
        identifier = normpath(filedialog.askdirectory(title='Введите директорию вывода'))
        if identifier == '.':
            return

    try:
        template = PdfReader(template_path)
    except (FileNotFoundError, PdfReadError):
        msgbox.showerror('Файл шаблона не найден', 'Файл шаблона не найден')
        return
    writer = PdfWriter()

    fields = template.get_fields()

    fix_acroform(writer, template)

    # Disable the validation and format on these fields
    phone_field_names = ['Phone Number1', 'Phone Number2', 'Phone Number3', 'Phone Number4', 'CELL NUMBER']

    for phone_field in phone_field_names:
        disable_field_validation(fields[phone_field])

    db = read_excel(
        db_path, sheet_name=sheet_name, header=header_row, keep_default_na=False
    )
    keys = db.keys().values
    identifier_int = db.where(db == identifier).dropna(how='all').dropna(axis=1).index

    for page in template.pages:
        writer.add_page(page)

    fill_fields = {}
    for field in fields.values():
        try:
            if field['/V'] in keys:
                fill_fields[field['/T']] = db[field['/V']][identifier_int].values[0]
        except KeyError:
            continue
        except TypeError:
            error('PDF Disclaimer: incorrect form identifier')
            return

    print(fields)

    fill_fields['Rebate Application Numbers'] = db["AppNumber"][identifier_int].values[0]
    fill_fields['Phone Number1'] = db["ContactPhone"][identifier_int].values[0]
    fill_fields['Phone Number3'] = db["ContractorPhone"][identifier_int].values[0]
    fill_fields['CELL NUMBER'] = db["ContractorPhone"][identifier_int].values[0]

    fill_fields['Verification Contact Name'] = db["ContactName"][identifier_int].values[0]
    fill_fields['Verification Contact Title'] = db["ContactTitle"][identifier_int].values[0]
    fill_fields['Verification Contact EMAIL'] = db["ContactEmail"][identifier_int].values[0]
    fill_fields['Contractor Firm Name'] = db["ContractorFirm"][identifier_int].values[0]
    fill_fields['Contractor EMAIL'] = db["ContractorEmail"][identifier_int].values[0]
    fill_fields['CustTitle'] = db["ContactTitle"][identifier_int].values[0]
    fill_fields['ContTitle'] = db["ContractorTitle"][identifier_int].values[0]

    for page in writer.pages:
        if '/Annots' not in page:
            continue
        writer.update_page_form_field_values(page, fill_fields)

    identifier += '\\Disclaimer\\'
    makedirs(identifier, exist_ok=True)

    project_name = db['ProjectName'][identifier_int].values[0]
    disclaimer_file_name = f"{project_name} - Disclaimer IP - {today_as_string()}.pdf"
    disclaimer_file_path = identifier + disclaimer_file_name

    with open(disclaimer_file_path, 'wb') as output_stream:
        writer.write(output_stream)

    return disclaimer_file_path


# Fix form field values displayed only after a click in Adobe Reader
def fix_acroform(writer: PdfWriter, reader: PdfReader) -> None:
    reader_root = typing.cast(DictionaryObject, reader.trailer[TrailerKeys.ROOT])
    acro_form_key = NameObject(CatalogDictionary.ACRO_FORM)

    if CatalogDictionary.ACRO_FORM in reader_root:
        reader_acro_form = reader_root[CatalogDictionary.ACRO_FORM]
        writer._root_object[acro_form_key] = writer._add_object(reader_acro_form.clone(writer))
    else:
        writer._root_object[acro_form_key] = writer._add_object(DictionaryObject())

    writer.set_need_appearances_writer()


def disable_field_validation(field):
    del field["/AA"]["/F"]
    del field["/AA"]["/K"]


if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='AutoPfd',
                                     description='Автоматически заполняет форму Pdf-файла информацией из базы данных')
    parser.add_argument('-t', '--template', type=str, help='Путь к Pdf шаблону')
    parser.add_argument('-i', '--identifier', type=str, help='Идентификатор строки в бд')
    args = parser.parse_args()

    redact(args.template, args.identifier)
