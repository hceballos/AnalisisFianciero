# -*- coding: utf-8 -*-
# https://packaging.python.org/tutorials/packaging-projects/
from lib.spreadsheet import Spreadsheet
from os import path
from os import remove

excel_path_Devengo  = 'input/Cartera_Financiera_Presupuestaria_Devengo/*/*'
excel_path_Contable = 'input/Cartera_Financiera_Contable/*/*'

if path.exists("save_pandas.db"):
    remove('save_pandas.db')

if __name__ == '__main__':

    Libro = Spreadsheet(excel_path_Devengo, excel_path_Contable)
