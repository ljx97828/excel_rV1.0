import pandas as pd

def excel_rnum_col_row(col,row):
    df=pd.read_excel('D:/2018dachuang/excel_rw/excel_rw/exec.xlsx')
    sinf=df.loc[col,row]
    return sinf

def add(a,b):
    df=pd.read_excel('D:/2018dachuang/excel_rw/excel_rw/exec.xlsx')

    return a+b;
    
