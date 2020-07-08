import keyboard
import mouse
import time
import xlrd
import xlwt
import pandas as pd
from tkinter import *
from tkinter import messagebox #消息框模块
import tkinter.filedialog #文件导入模块
import re #正则表达式
import os
import openpyxl

workbook_pd = pd.read_excel(r'./anjian模板.xlsx')
print(workbook_pd)

from openpyxl import Workbook
from openpyxl import load_workbook
workbook_openpyxl = load_workbook(r'./anjian模板.xlsx')
print(workbook_openpyxl)
workbook_openpyxl.save(r'./模板导出.xlsx')