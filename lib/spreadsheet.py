import pandas as pd
import numpy as np
import glob
from xls2db import xls2db
from os import path
from os import remove
import sqlite3
import re
import datetime
import time

class Spreadsheet():

    def __init__(self, excel_path_Devengo, excel_path_Contable):
        self.excel_path_Devengo = excel_path_Devengo
        self.excel_path_Contable = excel_path_Contable
        self.read_Excel()

    def read_Excel(self):
        if path.exists("devengo.db"):
            remove('devengo.db')
            remove('output.xlsx')

        all_data = pd.DataFrame()
        for f in glob.glob(self.excel_path_Devengo):
            df = pd.read_excel(f)
            print("Analizando  : ", f)
            all_data = all_data.append(df,ignore_index=True)

        all_data_Contable = pd.DataFrame()
        for f in glob.glob(self.excel_path_Contable):
            df = pd.read_excel(f)
            print("Analizando  : ", f)
            all_data_Contable = all_data_Contable.append(df,ignore_index=True)

        all_data['Tipo Vista']          = all_data.drop( all_data[ all_data['Tipo Vista'] == 'Saldo Inicial' ].index , inplace=True )
        all_data['Monto Documento.1']   = [w.replace('(', '-') for w in all_data['Monto Documento.1']]
        all_data['Monto Documento.1']   = [w.replace(')', '' ) for w in all_data['Monto Documento.1']]
        all_data['Monto Documento.1']   = [w.replace('.', '' ) for w in all_data['Monto Documento.1']]
        all_data['Monto Documento.1']   = pd.to_numeric(all_data['Monto Documento.1'])
        all_data['Monto Documento']     = [w.replace('.', '' ) for w in all_data['Monto Documento']]
        all_data['Monto Documento']     = pd.to_numeric(all_data['Monto Documento'])
        all_data["N Concepto"]          = all_data["Concepto"].str.split(" ", n = 1, expand = True)[0]
        all_data["Concepto Nombre"]     = all_data["Concepto"].str.split(" ", n = 1, expand = True)[1]
        all_data["Rut"]                 = all_data["Principal"].str.split(" ", n = 1, expand = True)[0]
        all_data["Rut Nombre"]          = all_data["Principal"].str.split(" ", n = 1, expand = True)[1]

        print( all_data["Fecha Generación"] )

        # ----------------------------------------------

        all_data_Contable['Tipo Vista']             = all_data_Contable.drop( all_data_Contable[ all_data_Contable['Tipo Vista'] == 'Saldo Inicial' ].index , inplace=True )
        all_data_Contable['Saldo']                  = [w.replace('(', '-') for w in all_data_Contable['Saldo']]
        all_data_Contable['Saldo']                  = [w.replace(')', '' ) for w in all_data_Contable['Saldo']]
        all_data_Contable['Saldo']                  = [w.replace('.', '' ) for w in all_data_Contable['Saldo']]
        all_data_Contable['Saldo']                  = pd.to_numeric(all_data_Contable['Saldo'])
        all_data_Contable['Debe']                   = [w.replace('(', '-') for w in all_data_Contable['Debe']]
        all_data_Contable['Debe']                   = [w.replace(')', '' ) for w in all_data_Contable['Debe']]
        all_data_Contable['Debe']                   = [w.replace('.', '' ) for w in all_data_Contable['Debe']]
        all_data_Contable['Debe']                   = pd.to_numeric(all_data_Contable['Debe'])
        all_data_Contable['Haber']                  = [w.replace('(', '-') for w in all_data_Contable['Haber']]
        all_data_Contable['Haber']                  = [w.replace(')', '' ) for w in all_data_Contable['Haber']]
        all_data_Contable['Haber']                  = [w.replace('.', '' ) for w in all_data_Contable['Haber']]
        all_data_Contable['Haber']                  = pd.to_numeric(all_data_Contable['Haber'])
        all_data_Contable['Saldo Acumulado']        = [w.replace('(', '-') for w in all_data_Contable['Saldo Acumulado']]
        all_data_Contable['Saldo Acumulado']        = [w.replace(')', '' ) for w in all_data_Contable['Saldo Acumulado']]
        all_data_Contable['Saldo Acumulado']        = [w.replace('.', '' ) for w in all_data_Contable['Saldo Acumulado']]
        all_data_Contable['Saldo Acumulado']        = pd.to_numeric(all_data_Contable['Saldo Acumulado'])
        all_data_Contable["N Cuenta Contable"]      = all_data_Contable["Cuenta Contable"].str.split(" ", n = 1, expand = True)[0]
        all_data_Contable["Cuenta Contable Nombre"] = all_data_Contable["Cuenta Contable"].str.split(" ", n = 1, expand = True)[1]
        all_data_Contable["Rut"]                    = all_data_Contable["Principal"].str.split(" ", n = 1, expand = True)[0]
        all_data_Contable["Rut Nombre"]             = all_data_Contable["Principal"].str.split(" ", n = 1, expand = True)[1]
        # ----------------------------------------------
        print("Creando Excel    ")
        writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
        print("Creando pestaña Devengo    ")
        all_data.to_excel(writer, sheet_name='Devengo')
        print("Creando pestaña Contable    ")
        all_data_Contable.to_excel(writer, sheet_name='Contable')
        writer.save()

        print("Creando Base de Datos    ")
        xls2db("output.xlsx", "devengo.db")


        """
        cnx = sqlite3.connect('devengo.db')
        consulta ="\
            SELECT \
                Contable.'Cuenta Contable', \
                Devengo.Concepto, \
                Devengo.Principal, \
                Contable.'Número Documento' , \
                Contable.'Debe', \
                Contable.'Haber'  \
            FROM \
                Contable INNER JOIN Devengo \
                ON Contable.Rut = Devengo.Rut \
                AND Contable.'Número Documento' = Devengo.'Número Documento'  \
            WHERE \
                Contable.Rut like '%' \
                and Contable.'Número Documento' like '%' \
        "

        datos = pd.read_sql_query(consulta, cnx)

        writer = pd.ExcelWriter('output_1.xlsx', engine='xlsxwriter')
        datos.to_excel(writer, sheet_name='Sql')
        writer.save()
        """

        # 10673579-4
        # 3167
        #https://stackabuse.com/converting-strings-to-datetime-in-python/
        #https://es.pornhubpremium.com/view_video.php?viewkey=ph5e14678dede01
