# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import glob
import sqlite3
from sqlalchemy import create_engine


class Spreadsheet():

    def __init__(self, excel_path_Devengo, excel_path_Contable):
        self.excel_path_Devengo = excel_path_Devengo
        self.excel_path_Contable = excel_path_Contable
        self.read_Excel()

    def read_Excel(self):
        devengo = pd.DataFrame()
        for f in glob.glob(self.excel_path_Devengo):
            df = pd.read_excel(f)
            print("Analizando  : ", f)
            devengo = devengo.append(df,ignore_index=True)

        contable = pd.DataFrame()
        for f in glob.glob(self.excel_path_Contable):
            df = pd.read_excel(f)
            print("Analizando  : ", f)
            contable = contable.append(df,ignore_index=True)


        devengo['Tipo Vista']          = devengo.drop( devengo[ devengo['Tipo Vista'] == 'Saldo Inicial' ].index , inplace=True )
        devengo['Monto Documento.1']   = [w.replace('(', '-') for w in devengo['Monto Documento.1']]
        devengo['Monto Documento.1']   = [w.replace(')', '' ) for w in devengo['Monto Documento.1']]
        devengo['Monto Documento.1']   = [w.replace('.', '' ) for w in devengo['Monto Documento.1']]
        devengo['Monto Documento.1']   = pd.to_numeric(devengo['Monto Documento.1'])
        devengo['Monto Documento']     = [w.replace('.', '' ) for w in devengo['Monto Documento']]
        devengo['Monto Documento']     = pd.to_numeric(devengo['Monto Documento'])
        devengo["N Concepto"]          = devengo["Concepto"].str.split(" ", n = 1, expand = True)[0]
        devengo["Concepto Nombre"]     = devengo["Concepto"].str.split(" ", n = 1, expand = True)[1]
        devengo["Rut"]                 = devengo["Principal"].str.split(" ", n = 1, expand = True)[0]
        devengo["Rut Nombre"]          = devengo["Principal"].str.split(" ", n = 1, expand = True)[1]
        devengo["Fecha Generación"]    = pd.to_datetime(devengo["Fecha Generación"]).dt.date
        devengo["Folio"]               = devengo["Folio"].astype(str)
        devengo["Número Documento"]    = devengo["Número Documento"].astype(str)


        contable['Tipo Vista']             = contable.drop( contable[ contable['Tipo Vista'] == 'Saldo Inicial' ].index , inplace=True )
        contable['Saldo']                  = [w.replace('(', '-') for w in contable['Saldo']]
        contable['Saldo']                  = [w.replace(')', '' ) for w in contable['Saldo']]
        contable['Saldo']                  = [w.replace('.', '' ) for w in contable['Saldo']]
        contable['Saldo']                  = pd.to_numeric(contable['Saldo'])
        contable['Debe']                   = [w.replace('(', '-') for w in contable['Debe']]
        contable['Debe']                   = [w.replace(')', '' ) for w in contable['Debe']]
        contable['Debe']                   = [w.replace('.', '' ) for w in contable['Debe']]
        contable['Debe']                   = pd.to_numeric(contable['Debe'])
        contable['Haber']                  = [w.replace('(', '-') for w in contable['Haber']]
        contable['Haber']                  = [w.replace(')', '' ) for w in contable['Haber']]
        contable['Haber']                  = [w.replace('.', '' ) for w in contable['Haber']]
        contable['Haber']                  = pd.to_numeric(contable['Haber'])
        contable['Saldo Acumulado']        = [w.replace('(', '-') for w in contable['Saldo Acumulado']]
        contable['Saldo Acumulado']        = [w.replace(')', '' ) for w in contable['Saldo Acumulado']]
        contable['Saldo Acumulado']        = [w.replace('.', '' ) for w in contable['Saldo Acumulado']]
        contable['Saldo Acumulado']        = pd.to_numeric(contable['Saldo Acumulado'])
        contable["N Cuenta Contable"]      = contable["Cuenta Contable"].str.split(" ", n = 1, expand = True)[0]
        contable["Cuenta Contable Nombre"] = contable["Cuenta Contable"].str.split(" ", n = 1, expand = True)[1]
        contable["Rut"]                    = contable["Principal"].str.split(" ", n = 1, expand = True)[0]
        contable["Rut Nombre"]             = contable["Principal"].str.split(" ", n = 1, expand = True)[1]
        contable["Fecha"]                  = pd.to_datetime(contable["Fecha"]).dt.date

        engine = create_engine('sqlite:///save_pandas.db', echo = True)
        devengo.to_sql("devengo", con=engine)
        contable.to_sql("contable", con=engine)
