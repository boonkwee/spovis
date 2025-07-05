# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 15:50:54 2024

@author: chanboonkwee
"""
import os
import shutil
import numpy as np
import pandas as pd
import warnings
from xlrd import XLRDError
from tools.inputmapper import InputMapper
from tools.misc import (
    list_all_cols,
    new_tempfile,
    read_csv_lines,
    update_spreadsheet,
    )

# from tools.xwtools import set_sensitivity_label
from tools.xltools import unprotect_xlsx
from tools.spov_hash import hash_sha256


warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


_debug = False
if _debug:
    print ('Debug ON')
else:
    print ('Debug OFF')

class DataVisSPO:
    AppName = 'SPOVis'

    def __init__(self,
                 url: str='',
                 filename: str='',
                 output_path: str='',
                 output_filename: str='',
                 staging_path: str='',
                 filelist: str='',
                 caa_fmt: str='%d%m%y',
                 caa_regex: str='\d{6}',
                 password: str='',
                 stdout=print):

        self.logging = stdout
        self.tempfile = ''
        self.staging_path = staging_path
        self.mapped_drive = {}
        if url in ['', None]:
            share_point_file = self.YF_data_url
        else:
            share_point_file = url

        if output_path in ['', None]:
            output_path = os.path.join(os.environ['USERPROFILE'], r'Desktop')
            self.logging(f"No output path given, using default: {output_path}")

        if _debug:
            # output_path = self.staging_path
            output_path = os.path.join(os.environ['USERPROFILE'], r'Desktop')

        # Processing output file (Stocktrend)
        caa_flag = filename != output_filename

        try:

            if not caa_flag:
                self.obj = InputMapper(cell=filename,
                                       file_pattern=filename,
                                       url=share_point_file, date_fmt='', pattern='.*', verbose=False,
                                       stdout=self.logging)

                self.fullpath_filename = self.obj._file

            else:
                self.obj = InputMapper(cell=filename,
                                       file_pattern=filename,
                                       url=share_point_file, date_fmt=caa_fmt, pattern=caa_regex, verbose=False,
                                       stdout=self.logging)

                self.fullpath_filename = self.obj.fullpath_filename

            if pd.isna(output_filename):
                self.logging("'file_name_output' is blank, using input filename")
                drive, output_filename = os.path.split(self.fullpath_filename)

            basename = os.path.basename(output_filename)
            # basename contains value like 'some_data.xlsx'
            self.filename = basename
            staging_fqfn = os.path.join(self.staging_path, basename)
            fn, ext = os.path.splitext(staging_fqfn)
            self.tempfile     = os.path.join(self.staging_path, new_tempfile(basename))

            self.logging(f"Copying '{basename}' to cache")
            if _debug:
                self.logging("Have caa" if caa_flag else "No caa")

            # shutil.copy2(self.fullpath_filename, staging_fqfn)
            # Copy the password protected source data to the cache location and use a temp filename
            # src = self.fullpath_filename.replace("/", "\\")
            # tgt = self.tempfile.replace("/", "\\")
            # os.system(f'copy "{src}" "{tgt}"')
            shutil.copy2(self.fullpath_filename, self.tempfile)
            if not (pd.isna(password) or password == ''):
                self.logging("Removing password protection.")
                # Remove password protection
                unprotect_xlsx(self.tempfile, password)
            # Copy the cached source data (temp filename) to the target location,
            # with the original filename
            # src = self.tempfile.replace("/", "\\")
            # tgt = staging_fqfn.replace("/", "\\")
            # os.system(f'copy "{src}" "{tgt}"')
            shutil.copy2(self.tempfile, staging_fqfn)

        except PermissionError:
            self.logging("The file may be opened, please close the file and run the script again.")
            raise
        except FileNotFoundError as e:
            self.logging(f"Unable to process {basename} due to error: {str(e)}")
            raise
        except (Exception, OSError, UnboundLocalError) as e:
            self.logging(f"{str(e)}")
            raise

        # Sheet2 of filelist contains list of worksheet and column information to perform hashing.
        data_to_process_df = pd.read_excel(filelist, sheet_name='Sheet2')
        data_filelist = list(set(data_to_process_df.file_to_process))
        if basename in data_filelist:
            sheet_data = data_to_process_df.loc[data_to_process_df['file_to_process']==basename]
            for index, row in sheet_data.iterrows():
                if np.nan in row.to_list():
                    self.logging(f"Skip processing Sheet2 row {index+2} due to missing values:")
                    self.logging(row)
                    continue
                rows_to_skip =     row['rows_skip']
                sheet_name =       row['worksheet']
                #columns =          row['columns'].split(',')
                columns =          [s.replace(chr(1), ',') for s in row['columns'].replace('\\,', chr(1)).split(',')]
                clear_empty_rows = row['clear_empty_rows']
                try:
                    if ext == '.csv':
                        buf_df = read_csv_lines(staging_fqfn, rows_to_skip, encoding='utf-8-sig')
                        df = pd.read_csv(staging_fqfn, skiprows=rows_to_skip)
                    else:
                        df = pd.read_excel(staging_fqfn, sheet_name=sheet_name, skiprows=rows_to_skip)
                except XLRDError as e:
                    self.logging(f"Encounter {str(e)}")
                    raise
                df_columns = df.columns.to_list()
                cleaned_columns = [col for col in columns if col in df_columns]
                unknown_columns = [col for col in columns if col not in df_columns]
                # df.replace({pd.NaT: None}, inplace = True)

                if len(unknown_columns) > 0:
                    self.logging(f"Sheet2 row {index+2} included columns {','.join(repr(i) for i in unknown_columns)} not found in data, will be skipped.")
                    self.logging(list_all_cols(repr(i) for i in df_columns))
                if clear_empty_rows:
                    df.dropna(subset=cleaned_columns, how='all', inplace=True)
                self.logging(f"Processing '{sheet_name}' columns: ")
                for col in cleaned_columns:
                    self.logging(f"{df_columns.index(col):02d}: {repr(col)}")
                    df[col] = df[col].astype('str').apply(lambda x: hash_sha256(x.encode('utf-8')) if x not in [None, np.nan, ''] else None)
                try:
                    if ext == '.csv':
                        # Append the DataFrame to the CSV file, skipping 5 rows before appending
                        with open(staging_fqfn, 'w', newline='', encoding='utf-8') as file:
                            # Skip 5 rows by writing empty rows
                            for line in buf_df:
                                file.write(f"{','.join(line)}\n")
                            df.to_csv(file, header=True, index=False, mode='a', encoding='utf-8')

                    else:
                        # book = load_workbook(filename=staging_fqfn)
                        # writer = pd.ExcelWriter(path=staging_fqfn, engine='openpyxl', mode='a')
                        # writer.book = book
                        # writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                        # df.to_excel(writer, sheet_name=sheet_name, startrow=rows_to_skip+2, index=False)
                        # writer.save()
                        update_spreadsheet(path=staging_fqfn, _df=df,
                                            startrow=rows_to_skip+2, sheet_name=sheet_name)

                except AttributeError as e:
                    self.logging(f"Encountered error {str(e)} for {basename}")
                    # continue
                    raise

        if ext == '.xlsx':
            self.logging("Setting sensitivity label...")
            # set_sensitivity_label(staging_fqfn,
            #                       label_description='CONFIDENTIAL (CLOUD ELIGIBLE-SENSITIVE/ SENSITIVE HIGH)')

            # if not outcome:
            #     # try again
            #     oc_xlsx(staging_fqfn)
            #     outcome = set_sensitivity_label(staging_fqfn,
            #                           label_description='CONFIDENTIAL (CLOUD ELIGIBLE-SENSITIVE/ SENSITIVE HIGH)')
        output_fqfn = os.path.join(output_path, basename)
        self.logging(f"Copying from cache to '{output_fqfn}'")
        if staging_fqfn != output_fqfn:
            # src = staging_fqfn.replace("/", "\\")
            # tgt = output_fqfn.replace("/", "\\")
            # os.system(f'copy "{src}" "{tgt}"')
            shutil.copy2(staging_fqfn, output_fqfn)

    def __del__(self):
        if os.path.exists(self.tempfile):
            os.remove(self.tempfile)
        if hasattr(self, 'obj'):
            del self.obj

