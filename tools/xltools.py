# -*- coding: utf-8 -*-
"""
Created on Wed Aug  7 09:21:50 2024

@author: chanboonkwee
"""
import win32com.client
import psutil
import os

def close_hidden_excel():
    for proc in psutil.process_iter(['pid', 'name']):
        if 'excel' in proc.info['name'].lower():
            try:
                proc.kill()
                print(f"Closed Excel process with PID {proc.pid}")
            except psutil.AccessDenied:
                print(f"Access denied to close Excel process with PID {proc.pid}")


def unprotect_xlsx(filename:str='', pw_str:str=''):
    if not os.path.exists(filename):
        raise FileNotFoundError()
    xcl = win32com.client.Dispatch('Excel.Application')
    xcl.DisplayAlerts = False
    xcl.Visible= False
    # param = {
    #     'Filename'                  : filename,
    #     'ReadOnly'                  : False,
    #     'WriteResPassword'          : pw_str,
    #     'IgnoreReadOnlyRecommended' : True,
    #     'Password'                  : pw_str
    #     }
    # wb = xcl.Workbooks.Open(**param)
    wb = xcl.Workbooks.Open(Filename=filename,
                            Password=pw_str,
                            WriteResPassword=pw_str,
                            ReadOnly=False,
                            IgnoreReadOnlyRecommended=True)
    wb.Password = ''
    wb.Save()
    xcl.Quit()
    del xcl

def oc_xlsx(filename:str='', pw_str:str=''):
    if not os.path.exists(filename):
        raise FileNotFoundError()
    xcl = win32com.client.Dispatch('Excel.Application')
    xcl.DisplayAlerts = True
    xcl.Visible= True
    # param = {
    #     'Filename'                  : filename,
    #     'ReadOnly'                  : False,
    #     'WriteResPassword'          : pw_str,
    #     'IgnoreReadOnlyRecommended' : True,
    #     'Password'                  : pw_str
    #     }
    # wb = xcl.Workbooks.Open(**param)
    wb = xcl.Workbooks.Open(filename)
    wb.Save()
    xcl.Quit()
    del xcl


if __name__ == '__main__':
    # filename = 'C:/Users/chanboonkwee/Desktop/CMTF/Book1.xlsx'
    # unprotect_xlsx(filename, pw_str='123')
    # filename = select_xlsx_file()
    # unprotect_xlsx(filename, pw_str='123')
    close_hidden_excel()