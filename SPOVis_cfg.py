# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 14:11:43 2024

@author: chanboonkwee
"""
import os
import configparser
# import json

cfg = configparser.RawConfigParser()

cfg.read('spovis_automate.ini')

try:
    SPOVIS_PATH = eval(cfg['SPOVIS']['PATH'])

    SPOVIS_FILELIST = eval(cfg['SPOVIS']['FILE'])

    SPOVIS_TRIGGER_TIME = eval(cfg['SPOVIS']['TRIGGER_TIME'])

    SPOVIS_DURATION = eval(cfg['SPOVIS']['DURATION'])
except KeyError:
    pass

if __name__=='__main__':
    print(os.path.split(SPOVIS_FILELIST))
    print(SPOVIS_TRIGGER_TIME, type(SPOVIS_TRIGGER_TIME))
    print(SPOVIS_DURATION, type(SPOVIS_DURATION))
