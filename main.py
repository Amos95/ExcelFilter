#!/usr/bin/python
# -*- coding: GBK -*-
import copy
import tkinter
from tkinter import Button,Label,Entry
from collections import defaultdict
from datetime import datetime
from tkinter.filedialog import askdirectory

import openpyxl
import time
import os

import win32con
import win32ui

_SUFFIX_LIT = ['.xls', '.xlsx']
_OUTPUT_NAME = 'result.xlsx'
_OUTPUT_FOLDER = 'output'


# "C:\\Users\\Logon\\Desktop\\"

def get_merged_cells_value(merged, sheet, row_index, col_index):
    for (_, scol), (_, srow), (_, ecol), (_, erow) in merged:
        if srow <= row_index <= erow:
            if scol <= col_index <= ecol:
                cell_value = sheet.cell(srow, scol).value
                # print("get", row_index, col_index, srow, scol, cell_value)
                return cell_value
                break
    return None


def transfer_rows(pyxl_row):
    row_list = []
    for cell in pyxl_row:
        row_list.append(cell.value)
    print(row_list)
    return row_list


def write_result(key_row):
    out_wb = openpyxl.load_workbook("C:\\Users\\Logon\\Desktop\\中文、路径2\\result.xlsx")
    for key, value_rows in key_row.items():
        print(key)
        try:
            out_sheet = out_wb[key]
        except KeyError:
            out_sheet = out_wb.create_sheet(key)
        for value_row in value_rows:
            out_sheet.append(value_row)
    out_wb.save("C:\\Users\\Logon\\Desktop\\中文、路径2\\result.xlsx")


def obtain_row_contain_key(sheet, keys):
    """ obtain row that contain key
    :param sheet: sheet
    :param keys: the list of key
    :return: key_row
    """
    key_row = defaultdict(list)
    merged = sheet.merged_cells
    row_num = sheet.max_row
    col_num = sheet.max_column

    for row_idx in range(1, row_num):
        for col_idx in range(1, col_num):
            cell_value = sheet.cell(row_idx, col_idx).value
            if not(cell_value is None or cell_value == '') and cell_value in keys:
                row_list = []
                for cur_col in range(1, col_num):
                    cur_value = sheet.cell(row_idx, cur_col).value
                    if cur_value is None or cur_value == '':
                        # print('merged', row, cur_col)
                        cur_value = get_merged_cells_value(merged, sheet, row_idx, cur_col)
                    row_list.append(cur_value)
                print(row_list)
                key_row[cell_value].append(tuple(row_list))
    return key_row


def process_file(file, keys):
    """
    :param file:
    :param keys:
    :return:
    """
    file_wb = openpyxl.load_workbook(file)
    print("load file done")
    sheets = file_wb.sheetnames
    for per_sheet in sheets:
        file_sheet = file_wb[per_sheet]
        print("process sheet:"+file_sheet.title)
        write_result(obtain_row_contain_key(file_sheet, keys))


def obtain_excel_files(folder_path):
    """ get all excel files from in the specified folder with the suffix xls or xlsx
    :param folder_path: Absolute path of folder containing excel file
    :return: excel file list
    """
    excel_list = []
    walk = os.walk(folder_path)
    for path, dir_list, file_list in walk:
        for file_name in file_list:
            if os.path.splitext(file_name)[-1] in _SUFFIX_LIT:
                excel_list.append(os.path.join(path, file_name))
    return excel_list


def get_filter_path():
    api_flag = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST | win32con.OFN_EXPLORER
    file_type = 'All File(*.*)|*.*|' \
        'Excel File(*.xls .xlsx)|*.xls;*.xlsx|'\
        '|'
    fg = win32ui.CreateFileDialog(1, None, None, api_flag, file_type)
    fg.SetOFNInitialDir("C:")
    fg.DoModal()
    filter_path = fg.GetPathName()
    print(filter_path)


# Need to get rid of the repetition
def filter_main(filter_path, filter_keys):
    print(filter_path)
    print(filter_keys)
    excel_files = []
    if os.path.isdir(filter_path):
        excel_files = obtain_excel_files(filter_path)
    elif os.path.isfile(filter_path):
        excel_name = os.path.basename(filter_path)
        if os.path.splitext(excel_name)[-1] in _SUFFIX_LIT:
            excel_files.append(filter_path)
    else:
        print("it's not a folder or file.")
    if len(excel_files) > 0:
        print(excel_files)
    for excel_file in excel_files:
        print("process file:" + excel_file)
        process_file(excel_file, filter_keys)


class FileWinDow:
    def __init__(self, width, height):
        self.window = tkinter.Tk()
        self.width = width
        self.height = height
        self.folder_path = tkinter.StringVar()
        self.keys_list = tkinter.StringVar()

    def select_folder(self):
        self.folder_path.set(askdirectory())

    def config_window(self):
        screenwidth = self.window.winfo_screenwidth()
        screenheight = self.window.winfo_screenheight()
        size = '%dx%d+%d+%d' % (self.width, self.height, (screenwidth - self.width) / 2, (screenheight - self.height) / 2)
        self.window.geometry(size)

    def set_view(self):
        Label(self.window, text='文件夹路径:').grid(row=0, column=0)
        Entry(self.window, textvariable=self.folder_path).grid(row=0, column=1)
        Button(self.window, text='选择文件夹', command=self.select_folder).grid(row=0, column=2)
        Label(self.window, text='关键字:').grid(row=1, column=0)
        Entry(self.window, textvariable=self.keys_list).grid(row=1, column=1)
        Button(self.window, text='开始过滤', command=self.start_filter).grid(row=2, column=2)
        self.window.mainloop()

    def start_filter(self):
        get_path = self.folder_path.get()
        get_keys = self.keys_list.get()
        if not(get_path == "" or get_keys == ""):
            filter_main(get_path, get_keys.split())


if __name__ == "__main__":
    if 1:
        # target = "C:\\Users\\Logon\\Desktop\\中文路径\\"
        FileWinDow = FileWinDow(400, 500)
        FileWinDow.config_window()
        FileWinDow.set_view()
    else:
        FileWinDow = FileWinDow(400, 500)
        FileWinDow.config_window()
        FileWinDow.set_view()

