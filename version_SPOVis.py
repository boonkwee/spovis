# -*- coding: utf-8 -*-
"""
Created on Thu Feb  1 14:15:30 2024

@author: chanboonkwee
"""
from datetime import datetime
class SPOVis:
    __AppName__ = 'SPOVis'
    __version__ = '1.06a'
    # __date__  = datetime(year=2024, month=6, day=17)
    parse_date  = lambda today_date_dBY: datetime.strptime(today_date_dBY, "%d %B %Y")
    __date__    = parse_date("8 October 2024")
    __Author__  = 'Chan Boon Kwee'

if __name__=='__main__':
    print(SPOVis().__date__)
