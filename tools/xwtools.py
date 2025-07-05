# -*- coding: utf-8 -*-
"""
Created on Tue Jul 11 16:53:18 2023

@author: chanboonkwee
"""
import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
import win32com.client
import pythoncom
import json
import psutil
from types import ModuleType



labels = {
    'OFFICIAL OPEN' :                                            '5434c4c7-833e-41e4-b0ab-cdb227a2f6f7',
    'RESTRICTED NON SENSITIVE' :                                 '54803508-8490-4252-b331-d9b72689e942',
    'RESTRICTED SENSITIVE NORMAL':                               '153db910-0838-4c35-bb3a-1ee21aa199ac',
    'RESTRICTED SENSITIVE HIGH':                                 '9c789c2b-f6c5-4645-8d26-38c740fa1736',
    'OFFICIAL CLOSE NON SENSITIVE':                              '4aaa7e78-45b1-4890-b8a3-003d1d728a3e',
    'OFFICIAL CLOSE SENSITIVE NORMAL':                           '770f46e1-5fba-47ae-991f-a0785d9c0dac',
    'OFFICIAL CLOSE SENSITIVE HIGH':                             'a8737adb-0c79-46a2-8440-0996bc024fec',
    'CONFIDENTIAL (CLOUD ELIGIBLE-SENSITIVE/ NON SENSITIVE)':    '0cdb6729-b45c-4a11-ac47-f8584fc7ec0a',
    'CONFIDENTIAL (CLOUD ELIGIBLE-SENSITIVE/ SENSITIVE NORMAL)': 'c477f0d0-1a40-415d-b130-8b40a41c8c21',
    'CONFIDENTIAL (CLOUD ELIGIBLE-SENSITIVE/ SENSITIVE HIGH)':   '8b393779-7a9e-4c5c-a9eb-1cf7cd3fa8f3',
    'CONFIDENTIAL NON SENSITIVE':                                '4f288355-fb4c-44cd-b9ca-40cfc2aee5f8',
    'CONFIDENTIAL SENSITIVE NORMAL':                             '',
    'CONFIDENTIAL SENSITIVE HIGH':                               '',
    }


def get_description(ssid:str='')-> str:
    try:
        cleaned_ssid = ssid.lower().strip()
        if cleaned_ssid not in labels.values():
            print(cleaned_ssid)
            raise ValueError('Classification not found')
        return list(labels.keys())[list(labels.values()).index(cleaned_ssid)]
    except AttributeError:
        print(f"{type(ssid)} - '{ssid}'")
        return ssid

def select_xlsx_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    return file_path

def select_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(filetypes=[
        ("Excel files", "*.xlsx"),
        ("Word files", "*.docx"),
        ("PowerPoint files", "*.pptx")])
    # file_path = filedialog.askopenfilename()
    return file_path


def set_sensitivity_label(filename, label_description='RESTRICTED SENSITIVE NORMAL', with_repair:bool=False):
    wb = None
    p = None
    if filename is [None, ''] or label_description in [None, '']:
        raise ValueError('filename and label description cannot be blank')
    label_id = labels.get(label_description)
    if label_id is None:
        raise ValueError('Invalid label description')
    try:
        # with xw.App(visible=False) as app:
        #     app.display_alerts = False
        # app = xw.App(visible=False, add_book=False)
        app = xw.App(visible=False)
        app.display_alerts = False
        # wb = app.books.api.Open(filename, CorruptLoad=1)
        if with_repair:
            wb = xw.Book(filename, corrupt_load=1, ignore_read_only_recommended=True)
        else:
            wb = xw.Book(filename, ignore_read_only_recommended=True)
        labelinfo = wb.api.SensitivityLabel.CreateLabelInfo()
        # labelinfo = wb.SensitivityLabel.CreateLabelInfo()
        labelinfo.AssignmentMethod = 2
        labelinfo.Justification = 'init'
        labelinfo.LabelId = label_id
        wb.api.SensitivityLabel.SetLabel(labelinfo, labelinfo)
        # wb.SensitivityLabel.SetLabel(labelinfo, labelinfo)
        p = psutil.Process(app.pid)
        wb.save(filename)
        wb.close()
        p.kill()
        del wb, app

    except pythoncom.com_error as e:
        print(f"Encounter {str(e)}. {filename}")
        return
    except Exception:
        raise
    finally:
        if p is not None:
            p.kill()


def k_excel(proc:ModuleType=None):
    if proc is not None:
        ex_p = psutil.Process(proc.pid)
        ex_p.kill()


def set_sensitivity_label_pwd(filename, label_description='RESTRICTED SENSITIVE NORMAL', password=''):
    if filename is [None, ''] or label_description in [None, '']:
        raise ValueError('filename and label description cannot be blank')
    label_id = labels.get(label_description)
    if label_id is None:
        raise ValueError('Invalid label description')

    app = xw.apps.keys()
    wb = xw.Book(filename, password=password, update_links=False)
    # app = xw.apps.active
    labelinfo = wb.api.SensitivityLabel.CreateLabelInfo()
    labelinfo.AssignmentMethod = 2
    labelinfo.Justification = 'init'
    labelinfo.LabelId = label_id
    wb.api.SensitivityLabel.SetLabel(labelinfo, labelinfo)
    wb.save()
    if app:
        wb.close()
    else:
        wb.app.quit()


def get_sensitivity_label(filename):
    app = xw.apps.keys()
    wb = xw.Book(filename)
    obj = wb.api.SensitivityLabel.GetLabel()
    labelinfo = str(obj())
    wb.save()
    if app:
        wb.close()
    else:
        wb.app.quit()
    return labelinfo


def get_file_sensitivity_label(file_path):
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        folder = shell.NameSpace(file_path)
        properties = folder.GetDetailsOf(folder.Items(), 266)  # 266 is the index for sensitivity label

        sensitivity_label = properties if properties else "Sensitivity label not found"
        print(f"The sensitivity label for the file is: {sensitivity_label}")
    except Exception as e:
        print(f"An error occurred: {e}")


def get_sensitive_setting_details(file_path:str='') -> str:
    print(repr(file_path))
    if file_path in ('', None):
        raise ValueError('filename and label description cannot be blank')
    try:
        if '/' in file_path:
            file_path = file_path.replace('/', '\\')
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Set to True if you want to see the Word application

        doc = word.Documents.Open(file_path)
        slbl = doc.SensitivityLabel.GetLabel()
        # settings = doc.BuiltInDocumentProperties

        # print(f"The sensitive setting for the document is: {slbl}")

        doc.Close()
        word.Quit()
        return f"{slbl}"
    except Exception as e:
        print(f"An error occurred: {e}")


def get_pptx_details(file_path:str='') -> str:
    print(repr(file_path))
    if file_path in ('', None):
        raise ValueError('filename and label description cannot be blank')
    try:
        if '/' in file_path:
            file_path = file_path.replace('/', '\\')
        ppt = win32com.client.Dispatch("PowerPoint.Application")

        doc = ppt.Presentations.Open(file_path)
        slbl = doc.SensitivityLabel.GetLabel()
        # settings = doc.BuiltInDocumentProperties

        # print(f"The sensitive setting for the document is: {slbl}")

        doc.Close()
        ppt.Quit()
        return f"{slbl}"
    except Exception as e:
        print(f"An error occurred: {e}")


#Error
def get_accdb_details(file_path:str='') -> str:
    print(repr(file_path))
    if file_path in ('', None):
        raise ValueError('filename and label description cannot be blank')
    try:
        if '/' in file_path:
            file_path = file_path.replace('/', '\\')
        oAccess = win32com.client.Dispatch('Access.Application')
        dbLangGeneral = ';LANGID=0x0409;CP=1252;COUNTRY=0'
        # dbVersion40 64
        dbVersion = 128
        oAccess.DBEngine.OpenDatabase(file_path, dbLangGeneral, dbVersion)
        oAccess.Quit()

        slbl = oAccess.SensitivityLabel.GetLabel()
        # settings = doc.BuiltInDocumentProperties

        # print(f"The sensitive setting for the document is: {slbl}")

        oAccess.Close()
        oAccess.Quit()
        return f"{slbl}"
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == '__main__':
    source_file = os.path.join(os.getcwd(), 'labels2.json')
    print(f"Loading {source_file}")
    with open(source_file, 'r', encoding ='utf8') as json_file:
        labels = json.load(json_file)

    ssid = ''
    lbl = select_file()
    fn, ext = os.path.splitext(lbl)
    if ext.startswith('.doc'):
        ssid = get_sensitive_setting_details(lbl)
    elif ext.startswith('.xls'):
        ssid = get_sensitivity_label(lbl)
    elif ext == '.accdb':
        ssid = get_accdb_details(lbl)
    elif ext.startswith('.ppt'):
        ssid = get_pptx_details(lbl)

    if ssid == '':
        raise ValueError("No filename given")

    print(f"{ssid}")
    print(f"Classification: {get_description(ssid)}")

    with open(source_file, 'w', encoding ='utf8') as json_file:
        json.dump(labels, json_file, ensure_ascii = False)

    # print(get_file_sensitivity_label(lbl))
