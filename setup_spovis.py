# -*- coding: utf-8 -*-
"""
Created on Thu Sep 14 11:06:44 2023

@author: chanboonkwee
"""
import os
import shutil
from datetime import datetime


if __name__=='__main__':
    cwd = os.getcwd()
    i = 0
    file_list = [
        'SPOVis_gui.py',
        'SPOVis_engine.py',
        'SPOVis_cfg.py',
        'version_SPOVis.py',
        'im_logger.py',
        'setup_spovis.py',
        'tools/inputmapper.py',
        'tools/misc.py',
        'tools/xwtools.py',
        'tools/xltools.py',
        'tools/spov_hash.py',
        'img/clipboard.png',
        'img/exit.png',
        'img/launch.png',
        'img/logpath.png',
        'spovis.json',
        'spovis_automate.ini',
        ]

    # iterate through each file
    print(f"{datetime.now()}")
    for p in file_list:
        # formulate the path to the file
        pth = os.path.join(cwd, p)
        # print(f'{pth}')

        # if the file does not exist
        if not os.path.exists(pth):
            print (f'{pth} do not exist')
            # increment i
            i += 1
    if i==0:
        v = 'spovis/'
        if not os.path.exists(v):
            os.mkdir(v)
            print(f"{v} folder created.")
        for src_file in file_list:
            dest = os.path.join(cwd, v, src_file)

            # split up the path and file,
            pth, file = os.path.split(dest)

            # check if the path exists,
            if not os.path.exists(pth):
                # does not exist, create the folder
                os.mkdir(pth)
                print(f"'{pth}' path created")

            print(f'Copy {src_file:25s} to {dest[-40:]:40s}', end='')
            try:
                shutil.copy(src_file, dest)
                print('.. ok')
            except PermissionError:
                print(f'Attempted to copy \'{src_file}\' but failed.')

    # i is not zero, at least one file is missing
    else:
        print('At least one file is missing')
