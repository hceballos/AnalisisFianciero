# -*- coding: utf-8 -*-
from xls2db import xls2db
from lib.spreadsheet import Spreadsheet

excel_path_Devengo  = 'input/Cartera_Financiera_Presupuestaria_Devengo/*/*'
excel_path_Contable = 'input/Cartera_Financiera_Contable/*/*'

if __name__ == '__main__':

    Libro = Spreadsheet(excel_path_Devengo, excel_path_Contable)
