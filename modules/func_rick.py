#import pandas as pd
from pandas import Period
import yfinance as yf
import numpy as np
import matplotlib.pyplot as plt
import xlwings as xw

def data(ticket:str,s:str,f:str,interval:str):
    df = yf.download(ticket,start=s,end=f,interval = interval)
    df.drop(['Open','High','Low','Close','Volume'], axis = 1, inplace=True)
    df.dropna(inplace=True)
    return df

def ln_rend(df):
    df['R_LN'] =  np.log(df['Adj Close']/df['Adj Close'].shift(periods=1))

def con_excel(df, ticket:str):
    wb = xw.Book()
    sheet_data = wb.sheets[0]
    sheet_data
    xw.sheets.active
    xw.books.active
    sheet_data.name = ticket
    sheet_data.range('A1').value = df



