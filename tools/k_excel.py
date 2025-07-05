# -*- coding: utf-8 -*-
"""
Created on Fri Jun 28 09:26:17 2024

@author: chanboonkwee
"""
import psutil

def close_hidden_excel():
    for proc in psutil.process_iter(['pid', 'name']):
        if 'excel' in proc.info['name'].lower():
            try:
                proc.kill()
                print(f"Closed Excel process with PID {proc.pid}")
            except psutil.AccessDenied:
                print(f"Access denied to close Excel process with PID {proc.pid}")


close_hidden_excel()