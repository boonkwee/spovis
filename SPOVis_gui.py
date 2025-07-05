"""
Created on Tue Jun 25 09:55:05 2024

@author: chanboonkwee
"""
import time
import os
import json
import pandas as pd
from enum import Enum
from hashlib import sha256
from PIL import Image, ImageTk
from datetime import datetime
import tkinter as tk
from tkinter import (
    filedialog,
    messagebox,
    #Menu,
    ttk
    )
from im_logger import Logger
from version_SPOVis import SPOVis as SP
from SPOVis_engine import DataVisSPO, _debug
from tools.misc import (
    as_local,
    disconnect_all_net_drive,
    drives_in_use,
    get_missing_items,
    select_xlsx_file
    )
try:
    from SPOVis_cfg import SPOVIS_PATH, SPOVIS_FILELIST
    from SPOVis_cfg import SPOVIS_TRIGGER_TIME, SPOVIS_DURATION
except:
    pass

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass
import ctypes

def display_on():
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

def display_reset():
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

sources = r'http://Path.pointing.to/on-prem.sharepoint.accessible.from.local/containing_datasets/'

class SPOVIS(tk.Tk):
    # AppName = 'SPOVis'
    AppName = f"{SP.__AppName__} {SP.__version__} ({SP.__date__.strftime('%d %b %Y')})"
    configfile = 'spovis.json'
    ini_file = 'spovis_automate.ini'

    class img_names(Enum):
        img_clipboard = 'img/clipboard.png'
        img_launch    = 'img/launch.png'
        img_exit      = 'img/exit.png'
        img_logpath   = 'img/logpath.png'

    def __init__(self, *args, **kwargs):
        try:
            self.spovis_path = SPOVIS_PATH
            self.spovis_filelist = SPOVIS_FILELIST
        except NameError:
            pass

        try:
            self.TRIGGER_TIME = SPOVIS_TRIGGER_TIME
            self.DURATION = SPOVIS_DURATION
        except NameError:
            self.TRIGGER_TIME = ()
            self.DURATION = 30

        self.settings = None
        with open(self.configfile, 'r') as fp:
            self.settings = json.load(fp)
            fp.close()
        self.image_test()
        self.repeated_ends = 2
        self.run_flag = False
        super().__init__(*args, **kwargs)
        self.setup()
        self.menu_setup()
        display_on()
        self.auto_run()
        # self.mainloop()

    def auto_run(self):
        current_time = datetime.now()
        if current_time.minute in self.TRIGGER_TIME:
            # if hits X:15 or X:45
            self.after(self.DURATION * 60 * 1000, self.auto_run)
            self.run()
        else:
            # check every minute
            self.after(60 * 1000, self.auto_run)
            self.logging('waiting...')

    def image_test(self):
        """
        Performs a test on the button image files required by the App.
        First checks if the file with the specific filename exist.
        Then perform a SHA256 hash of the file image, compare it
        against pre-captured hash stored in a json file.
        Either the filename missing or hash mismatch will cause the App to fail.

        Raises
        ------
        FileNotFoundError
            DESCRIPTION.
        IOError
            DESCRIPTION.

        Returns
        -------
        None.

        """

        img_dict = self.settings['button image']

        for img in self.img_names:
            if not os.path.exists(img.value):
                raise FileNotFoundError(f"Button image '{img.value}' not found")
            else:
                with open(img.value, 'rb') as f:
                    dump = f.read()
                    f.close()
                    img_hash = sha256(dump).hexdigest()
                    # print(f"{img.value}: [{sha_dump}]")
                    if img_hash != img_dict[img.value]:
                        # print(f"{img.value} ok")
                        raise IOError(f"Corrupted image : '{img.value}'")
        print('Button image files in order.')


    def setup(self):
        # self.wm_attributes('-topmost', 1)
        self.lift()
        # self.withdraw()
        self.option_add("*tearOff", False)
        self.geometry('1040x600')
        self.protocol("WM_DELETE_WINDOW", self.end)
        self.title(self.AppName)
        # font.nametofont("TkDefaultFont").configure(size=25)
        # self.after(100, self.update)

    def menu_setup(self):
        """
        Creation of Menu and Text widget on screen

        Returns
        -------
        None.

        """
        # mb = Menu(self)
        # self.config(menu=mb)
        # self.option_add("*Font", 'Verdana 14')
        # system = Menu(mb, font = ("Verdana", 12))
        # system.add_command(label='Run', command=self.run, font = ("Verdana", 12))
        # system.add_command(label='Quit', command=self.end, font = ("Verdana", 12))

        # mb.add_cascade(menu=system, label='System', font = ("", 16))
        ribbon = ttk.Frame(self, padding=(10,0,10,5))
        ribbon.pack(side='top', fill='both')

        img_run = Image.open('img/launch.png')
        self.btn_run = ImageTk.PhotoImage(img_run)
        self.button_run = ttk.Button(ribbon, text='Launch', command=self.run, image=self.btn_run)
        self.button_run.grid(ipady=20, ipadx=20)
        self.button_run.pack(side='left')

        img_logpath = Image.open('img/logpath.png')
        self.btn_logpath = ImageTk.PhotoImage(img_logpath)
        self.button_logpath = ttk.Button(ribbon, text='Log Path', command=self.select_new_log_file, image=self.btn_logpath)
        self.button_logpath.pack(side='left')

        img_clipboard = Image.open('img/clipboard.png')
        self.btn_copy = ImageTk.PhotoImage(img_clipboard)
        self.button_copy = ttk.Button(ribbon, text='Copy', command=self.copytext, image=self.btn_copy)
        self.button_copy.pack(side='left')

        img_exit = Image.open('img/exit.png')
        self.btn_exit = ImageTk.PhotoImage(img_exit)
        self.button_exit = ttk.Button(ribbon, text='Exit', command=self.end, image=self.btn_exit)
        self.button_exit.pack(side='right')

        # Display information about ini file and trigger time
        information_panel = ttk.Frame(ribbon, padding=(10,0,10,5))
        information_panel.pack(side='right')
        # information_panel.grid(ipady=20, ipadx=20)

        # Labels, on the left of the information panels
        labels_panel = ttk.Frame(information_panel)
        labels_panel.pack(side='left')

        label_ini = ttk.Label(labels_panel, text='Ini file :')
        label_ini.pack(side='top')

        label_tt = ttk.Label(labels_panel, text='Trigger : ')
        label_tt.pack(side='top')

        # Details, on the right of the labels
        details_panel = ttk.Frame(information_panel)
        details_panel.pack(side='left')

        self.ini_info = tk.Text(details_panel, width=20,
                                    height=1, state='normal', wrap='none', font = ("Courier", 10), bg='lightgray')

        self.ini_info.pack(side='top')
        self.ini_info.insert('1.0', self.ini_file if os.path.exists(self.ini_file) else 'missing')
        self.ini_info['state'] = 'disabled'

        self.tt_info = tk.Text(details_panel, width=20,
                                    height=1, state='normal', wrap='none', font = ("Courier", 10), bg='lightgray')

        self.tt_info.pack(side='top')
        self.tt_info.insert('1.0', ','.join(map(str, self.TRIGGER_TIME)) )
        self.tt_info['state'] = 'disabled'


###############################################################################
        # labels_panel = ttk.Frame(information_panel)
        # labels_panel.pack(side='left')

        # label_src = ttk.Label(labels_panel, text='File list :')
        # label_src.pack(side='top')

        # label_log = ttk.Label(labels_panel, text='Log file : ')
        # label_log.pack(side='top')

        # text_entry_panels = ttk.Frame(information_panel)
        # text_entry_panels.pack(side='left', fill='x', expand=True)

        # self.source_path = tk.Text(text_entry_panels,
        #                            height=1, state='normal', wrap='none', font = ("Courier", 10), bg='lightgray')

        # self.source_path.pack(side='top', fill='both', expand=True)
        # self.source_path.insert('1.0', self.settings['file list path'])
        # self.source_path['state'] = 'disabled'

        # self.log_file_path = tk.Text(text_entry_panels,
        #                              height=1, width=120,
        #                              state='normal', wrap='none', font = ("Courier", 10), bg='lightgray')

        # self.log_file_path.pack(side='top', fill='both', expand=True)
        # self.log_file_path.insert('1.0', self.settings['log file path'])
        # self.log_file_path['state'] = 'disabled'

###############################################################################


        log_file_panel = ttk.Frame(self, padding=(10,0,10,5))
        log_file_panel.pack(side="top", fill='both')
        self.log_file_path = tk.Text(log_file_panel,
                                     height=1, width=120,
                                     state='normal', wrap='none', font = ("Courier", 10))

        self.log_file_path.pack(side='left', fill='both', expand=True)
        self.log_file_path.insert('1.0', self.settings['log file path'])
        self.log_file_path['state'] = 'disabled'

        main_panel = ttk.Frame(self, padding=(10,0,10,20))
        main_panel.pack(side="top", fill='both', expand=True)

        self.text = tk.Text(main_panel, height=8, width=120, state='disabled', wrap='none', font = ("Courier", 10))
        self.text.pack(side='left', fill='both', expand=True)
        # self.text.insert("1.0", "<b>Type something</b>")
        self.text.focus()

        text_scroll = ttk.Scrollbar(main_panel, orient='vertical', command=self.text.yview)
        text_scroll.pack(side="left", fill='y')
        self.text['yscrollcommand'] = text_scroll.set


    def copytext(self):
        self.clipboard_clear()
        self.clipboard_append(self.text.get("1.0", "end"))
        self.update()
        self.logging("Log copied to clipboard")

    def end(self):
        self.settings['log file path'] = self.log_file_path.get('1.0', 'end').replace('\n', '')
        with open(self.configfile, 'w') as fp:
            json.dump(self.settings, fp)
            fp.close()
        display_reset()

        if not self.run_flag:
            self.destroy()
            self.quit()
        else:
            self.logging("Cannot exit when in running state")
            self.repeated_ends -= 1
            if not self.repeated_ends:
                self.run_flag = False
                self.switch_state()
                self.logging('Repeated close failed, running flag is turned off, click on close one more time to Exit.')


    def logging(self, *args):
        self.text['state'] = tk.NORMAL
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        self.text.insert("end", f"{ts}: {''.join(args)}\n")
        self.text['state'] = tk.DISABLED
        self.text.see('end')
        self.update()

    def switch_state(self):
        if self.run_flag:
            self.button_run.state(['disabled'])
            self.button_exit.state(['disabled'])
            self.wait_start()
        else:
            self.button_run.state(['!disabled'])
            self.button_exit.state(['!disabled'])
            self.wait_end()

    def wait_start(self):
        self.config(cursor='watch')
        self.update()

    def wait_end(self):
        self.config(cursor='')

    def select_new_log_file(self):
        log_file_path = self.log_file_path.get('1.0', 'end').replace('\n', '')
        pth, bn = os.path.split(log_file_path)
        new_log_file_path = select_xlsx_file(initialdir=pth, title='Select the log file')
        if len(new_log_file_path) > 0:
            self.log_file_path['state'] = 'normal'
            self.log_file_path.delete("1.0", "end")
            self.log_file_path.insert('1.0', new_log_file_path)
            self.log_file_path['state'] = 'disabled'


    def run(self):
        self.run_flag = True
        self.switch_state()
        self.logging(f"{SP.__AppName__} {SP.__version__} dated {SP.__date__.strftime('%d %b %Y')} by {SP.__Author__}")

        if hasattr(self, 'spovis_path'):
            self.output_path = self.spovis_path
        else:
            self.output_path = filedialog.askdirectory(title="Select the target folder for the de-identified data")

        self.original_drives = drives_in_use()

        if hasattr(self, 'spovis_filelist'):
            fl_path, a_filelist = os.path.split(self.spovis_filelist)
            self.link = as_local(sources, stdout=self.logging)
            self.remote_filelist = os.path.join(self.link.path, a_filelist)
        else:
            self.link = as_local(sources, stdout=self.logging)
            self.remote_filelist = select_xlsx_file(
                initialdir=self.link.path,
                title='Select the filelist in SPOVis Sources, select local copy if unavailable')

        if not os.path.exists(self.remote_filelist):
            msg = "Unable to proceed as filelist is missing."
            self.logging(msg)
            # raise FileNotFoundError()
            self.run_flag = False
            self.switch_state()

        self.logging(f"Processing {self.remote_filelist}")
        staging_path = os.path.join(os.environ['USERPROFILE'], r'Desktop/Cache/')
        if not os.path.exists(staging_path):
            try:
                self.logging(f"Creating {staging_path}..")
                os.mkdir(staging_path)
            except FileNotFoundError:
                self.logging(f"Encounter error creating path: [{staging_path}]")
                # raise
                self.run_flag = False
                self.switch_state()

        filelist = self.remote_filelist
        ex = []
        self.start_timestamp = datetime.now().strftime('%H:%M:%S.%f')
        self.start_date      = datetime.now().strftime('%d %b %Y')
        self.start_time      = time.time()
        if not os.path.exists(filelist):
            self.logging(f"No filelist selected: '{filelist}', unable to proceed.")
            self.run_flag = False
            self.switch_state()
            return
        filelist_df = pd.read_excel(filelist, sheet_name='Sheet1')

        file_in_error = []
        # files = []
        try:
            for index, r in filelist_df.iterrows():
                fn_input  = r['file_name_input_regex']
                fn_output = r['file_name_output']
                url       = r['file_url']
                ## Supporting version 0.3b
                try:
                    date_fmt  = r['caa_format']
                except KeyError:
                    date_fmt = ''
                try:
                    caa_regex = r['caa_regex']
                except KeyError:
                    caa_regex = ''

                try:
                    pwd = r['password_to_unlock_xlsx']
                except KeyError:
                    pwd = ''
                # files.append(os.path.join(output_path, fn_output))

                try:
                    ## Supporting version 0.3b
                    if [date_fmt, caa_regex] == ['','']:
                        self.o = DataVisSPO(output_path=self.output_path,
                                       url=url,
                                       filename=fn_input,
                                       output_filename=fn_output,
                                       staging_path=staging_path,
                                       filelist=filelist,
                                       password=pwd,
                                       stdout=self.logging)
                    else:
                        self.o = DataVisSPO(output_path=self.output_path,
                                       url=url,
                                       filename=fn_input,
                                       output_filename=fn_output,
                                       staging_path=staging_path,
                                       filelist=filelist,
                                       caa_fmt=date_fmt,
                                       caa_regex=caa_regex,
                                       password=pwd,
                                       stdout=self.logging)
                except FileNotFoundError as e:
                    ex.append(f"{str(e)} {e.filename}")
                    self.logging(f"Encounter {str(e)}, skipping {fn_input}")
                    file_in_error.append(fn_input)
                    raise
                    # continue
                except OSError as e:
                    ex.append(f"{str(e)}")
                    self.logging(f"Encounter {str(e)}, skipping {fn_input}")
                    file_in_error.append(fn_input)
                    raise
                    # continue

        except Exception as e:
            ex.append(str(e))
            # raise
            self.logging(f"{str(e)}")
            self.run_flag = False
            self.switch_state()
            raise
        finally:
            del self.link
            self.end_time            = time.time()
            self.end_timestamp = datetime.now().strftime('%H:%M:%S.%f')
            self.end_date      = datetime.now().strftime('%d %b %Y')
            hh = int((self.end_time - self.start_time) // 3600)
            mm = int((self.end_time - self.start_time) % 3600 // 60)
            ss = int((self.end_time - self.start_time) % 60)
            self.logging(f"Duration: {hh:02d}:{mm:02d}:{ss:02d}")

            ## Disable this segment to disable logging
            logfile_url = self.log_file_path.get('1.0', 'end').replace('\n', '')
            self.logging('Appending Log')
            try:
                log = Logger(fileurl=('' if _debug else logfile_url),
                              stdout=self.logging,
                              Script_Name=os.path.split(self.remote_filelist)[1],
                              Script_Version=SP.__version__,
                              Start_Date=self.start_date,
                              Start_Time=self.start_timestamp,
                              UserID=os.getlogin(),
                              End_Date=self.end_date,
                              End_Time=self.end_timestamp,
                              Status='Success' if len(ex)==0 else 'See Exception',
                              Exception='nil' if len(ex)==0 else '\n'.join(ex))
                del log
            except ValueError as e:
                self.logging(f"{str(e)}")

            if hasattr(self, 'o'):
                del self.o
            # self.logging('Processing drives to disconnect')
            drive_to_disconnect = get_missing_items(self.original_drives, drives_in_use())
            disconnect_all_net_drive(drive_to_disconnect, verbose=False, reply_to=self.logging)
            del drive_to_disconnect
            self.logging('All processes complete')
            self.run_flag = False
            self.switch_state()

        if len(file_in_error) > 0:
            error_msg = "Error: the following file(s) failed processing: \n%s" % '\n'.join(file_in_error)
            messagebox.showerror('Failed processing', error_msg)

if __name__=='__main__':
    o = SPOVIS()
    o.mainloop()
    del o
