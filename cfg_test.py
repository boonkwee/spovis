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

pth = eval(cfg['SPOVIS']['PATH'])
print(pth)
print(os.path.exists(pth))

fn = eval(cfg['SPOVIS']['FILE'])
print(fn)
print(os.path.exists(fn))
