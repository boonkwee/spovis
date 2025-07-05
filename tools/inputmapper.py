# -*- coding: utf-8 -*-
"""
Created on Mon Sep  4 10:06:49 2023

@author: chanboonkwee
"""
import re
import os
import traceback
import time
# import inspect
from datetime import datetime
from .misc import as_local, is_url, parse_net_use, unc_to_url
import numpy as np


class InputMapper:
    def __init__(self, cell: str='',
                 file_pattern:str='',
                 pattern: str='\d{8}',
                 url: str='',
                 date_fmt:str='%Y%m%d',
                 verbose: bool= True,
                 stdout=print):
        '''
        Create drive mapping to a sharepoint url

        Parameters
        ----------
        cell : str, optional
            DESCRIPTION. A short name to label the purpose of this mapping.
            The default is ''.
        file_pattern : str, optional
            DESCRIPTION. The default is ''.
        pattern : str, optional
            DESCRIPTION. The default is '\d{8}'.
        url : str, optional
            DESCRIPTION. The default is ''.
        date_fmt : str, optional
            DESCRIPTION. The default is '%Y%m%d'.

        Raises
        ------
        ValueError
            DESCRIPTION.

        Returns
        -------
        None.

        '''
        self.cell         = cell
        self.fullpath     = ''
        self.obj          = None
        self.mapped_path  = False
        self.file_pattern = file_pattern
        self.date_fmt     = '' if date_fmt is np.nan else date_fmt
        self.pattern      = '' if pattern is np.nan else pattern
        self.verbose      = verbose
        self.stdout       = stdout

        mapped_resources = parse_net_use()
        for drive, unc in mapped_resources.items():
            mapped_resources[drive] = unc_to_url(unc)

        if cell == '' or not (os.path.exists(url) or is_url(url)):
            # tb = traceback.format_exc()
            # raise ValueError('Required cell name and url or path for data source').with_traceback(tb)
            raise ValueError('Required cell name and url or path for data source\n'+f"Cell: '{cell}'; URL: '{url}'")
        self.cell = cell

        url_path, ext = os.path.splitext(url)
        if ext != '':
            url_path, url_file = os.path.split(url)
            url_path = url_path if url_path.endswith('/') else url_path + '/'
        else:
            # only path is available, url does not point to any file, set url_file empty
            url_file = ''
        self.file = url_file

        if url_path in mapped_resources.values():
            url_index        = list(mapped_resources.values()).index(url_path)
            self.drive       = list(mapped_resources.keys())[url_index]
            self.path        = self.drive
            self.fullpath    = os.path.join(f'{self.drive}:', os.sep, self.file)
            self.mapped_path = True
            if self.verbose:
                self.stdout('Re-using...')

        else:
            if is_url(url_path):
                self.mapped_path = True
                self.obj         = as_local(url_path if url_path.endswith('/') else url_path + '/',
                                            stdout=self.stdout)
                self.path        = self.obj.path

                # self.fullpath    = self.obj.fullpath
                # contains 'F:\\CMO Stocktrend File.xlsx' if self.mapped_path True
                # have to use self.path when self.mapped_path True

                self.drive       = self.obj.drive

            elif os.path.exists(url):
                base_path, ext      = os.path.splitext(url)
                alt_path, file_name = os.path.split(url)
                drive, _pth         = os.path.splitdrive(url)

                self.path        = url_path
                self.fullpath    = url
                self.drive       = drive

        if self.verbose:
            self.stdout(f'{self.path} for {self.cell}')

    @property
    def _file(self):
        file_found = ''
        try:
            for file_name in os.listdir(self.path):
                if file_name.startswith('~'):
                    continue
                if file_name == self.file_pattern:
                    file_found = file_name
                    break
        except Exception as e:
            self.stdout(f'{self.cell} Error:')
            raise e.with_traceback(e.__traceback__)
        return os.path.join(self.path, file_found)


    def file_dated(self, dt:datetime):
        i = 0
        dates = {}

        # calling_function_name = inspect.stack()[1][3]
        # self.stdout(f"The calling function is {calling_function_name}")
        # Loop through files in directory
        try:
            for file_name in os.listdir(self.path):
                if file_name.startswith('~'):
                    continue
                # self.stdout(f'{file_name}, pattern: {self.pattern}')
                # Check if file name matches pattern
                token = re.search(re.escape(self.file_pattern), file_name)
                if token is not None:
                    i += 1
                    # Extract date from file name
                    file_str = token.group()
                    # self.stdout(f'Found [{file_name}]')
                    date_token = re.search(self.pattern, file_str)
                    date_str = date_token.group()

                    # self.stdout(f'date_str is {date_str}')
                    filedate = datetime.strptime(date_str, self.date_fmt).date()
                    # Add date to list
                    dates[file_name] = filedate
        except Exception as e:
            self.stdout(f'{self.cell} Error:')
            raise e.with_traceback(e.__traceback__)
        if i == 0:
            self.stdout(f'{self.cell} Error:')
            tb = traceback.format_exc()
            raise ValueError(f'{self.cell}:  No files matched pattern: \'{self.pattern}\'').with_traceback(tb)

        # self.stdout(dates)
        # Get last file name based on date
        if len(dates) > 0:
            if dt.date() in dates.values():
                dated_file_name = [fn for fn, dt in dates.items() if dt == dt.date()][0]
                self.stdout(f'\nIdentified latest file: {self.drive}:[{dated_file_name}]')
                return dated_file_name
        return ''


    @property
    def latest_file(self):
        i = 0
        dates = []

        # calling_function_name = inspect.stack()[1][3]
        # self.stdout(f"The calling function is {calling_function_name}")
        # Loop through files in directory
        last_file_name = ''
        loop = True
        while loop:
            try:
                for file_name in os.listdir(self.path):
                    if file_name.startswith('~'):
                        continue
                    # self.stdout(f'{file_name}, pattern: {self.pattern}')
                    # Check if file name matches pattern
                    token = re.search(self.file_pattern, file_name)
                    if token is not None:
                        i += 1
                        # Extract date from file name
                        file_str = token.group()
                        # self.stdout(f'Found [{file_name}]')
                        date_token = re.search(self.pattern, file_str)
                        if date_token is not None:
                            fqfn = os.path.join(self.path, file_name)
                            if os.path.isfile(fqfn):
                                if os.path.getsize(fqfn) < 2:
                                    self.stdout(f"Skipping '{file_name}' due to 0 file size")
                                    continue
                            date_str = date_token.group()

                            # self.stdout(f'date_str is {date_str}')
                            if self.date_fmt == '':
                                last_file_name = file_name
                                break
                            else:
                                filedate = datetime.strptime(date_str, self.date_fmt).date()
                                # Add date to list
                                dates.append((file_name, filedate))
                        else:
                            last_file_name = file_name
                loop = False
            except Exception as e:
                self.stdout(f"{type(e).__name__}: Unable to connect to ({self.cell}){self.path}")
                self.__del__()
                raise e.with_traceback(e.__traceback__)

        if i == 0:
            self.stdout(f'{self.cell} : No files matched pattern: \'{self.pattern}\'')
            return last_file_name

        # self.stdout(dates)
        # Get last file name based on date
        if len(dates) > 0:
            last_file_name = max(dates, key=lambda x : x[1])[0]

        return last_file_name


    @property
    def fullpath_filename(self):
        fname = self.latest_file
        return fname if fname == '' else os.path.join(self.path, fname)


    def __del__(self):
        # calling_function_name = inspect.stack()[1][3]
        # self.stdout(f"The calling function is {calling_function_name}")
        if self.obj is not None:
            del self.obj
            self.obj = None

if __name__=='__main__':
    data_dest = 'http://Path.pointing.to/on-prem.sharepoint.accessible.from.local/containing_datasets/'
    obj = InputMapper(cell='Name identifying the process', pattern='Some Regex Pattern identifying specific files', url=data_dest)
    fn = obj.fullpath_filename
    exist_or_not = lambda x: "File exists" if os.path.exists(x) else "Does not exist"
    print(f'{fn} <-- %s' % exist_or_not(fn))
    # cc_data_file    = r'http://Excel.file.residing.in/on-prem.sharepoint.accessible.from.local/Workload%20Summary_20230831.xlsx'
    # cc = InputMapper(cell='test cell', url=cc_data_file)
    # print(f"Testing: {cc.file_dated(datetime(2023, 9, 29))}")
    # del cc
