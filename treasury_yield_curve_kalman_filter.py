#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Sat May 26 16:59:43 2018

@author: wyx
"""

import quandl
from pykalman import KalmanFilter
from numpy import ma
import pandas as pd
import xlwings as xw
import datetime as dt

EXCEL_FILE_NAME = 'kalman_filter.xlsm'
FACTORS = ['level', 'slope', 'convexity']
DATA_SHEET_NAME = 'Data'
PARAMETER_SHEET_NAME = 'Kalman Filter Parameters'
SMOOTH_SHEET_NAME = 'Smoothed Factors'
FILTER_SHEET_NAME = 'Filtered Factors'
CONFIG_SHEET_NAME = 'Config'

def getConfig():
    '''get configuration'''
    configValue = getSheetFromBook(xw.Book(EXCEL_FILE_NAME), CONFIG_SHEET_NAME).range('A2').expand().value
    
    return {l[0]:l[1] for l in configValue}
    
def updateTreasuryYieldData():
    '''load data from quandl and save as csv'''
    
    dateList = getDateList()
    
    if dateList:
        df = quandl.get("USTREASURY/YIELD", start_date = dateList[-1].date()+dt.timedelta(days=1))
        storeDataToExcel(df.reset_index().as_matrix(), DATA_SHEET_NAME, bookName=EXCEL_FILE_NAME, clear=False,
                         startPosition='A{}'.format(2+len(dateList)))
    else:
        df = quandl.get("USTREASURY/YIELD")
        storeDataToExcel(df, DATA_SHEET_NAME, bookName=EXCEL_FILE_NAME, clear=True)
    
def getSheetFromBook(book, sheetName):
    '''get the sheet with specified name'''
    
    return [s for s in book.sheets if s.name==sheetName][0]
    
def dfToMaskedArray(df):
    '''convert a data frame to a masked array'''

    return ma.array(df.as_matrix(),mask=df.isnull().as_matrix())

    
def storeDataToExcel(data, sheetName=None, index=None, 
                          startPosition='A1', bookName=None, clear=False, 
                          isCovariance=False):
    '''store a data frame to a sheet
    
    if isCovariance:
        
       store only the lower left triangle of a covariance matrix
       easier to handle the symmetry of covariance matrix
    
    '''
    
    if sheetName and index:
        raise Exception('Ambiguous input: both index and sheetName are specified')
    elif not sheetName and not index:
        raise Exception('One of following must be specified: sheet name, index ')
        
    book = xw.Book(bookName)
    
    sheet = book.sheets[index] if index else getSheetFromBook(book, sheetName)
        
    if clear:
        sheet.clear()
        
    if not isCovariance:
        sheet.range(startPosition).value = data
    else:
        putCovarianceMatrix(sheet, startPosition, data)
        
    
def putCovarianceMatrix(sheet, startPosition, data):
    '''store covariance matrix to a sheet'''
    
    startCol = ''.join([i for i in startPosition if not i.isdigit()])
    startRow = int(''.join([i for i in startPosition if i.isdigit()]))
    
    if isinstance(data, pd.DataFrame):
        data = data.as_matrix()
            
    rowNum, colNum = data.shape
    
    if rowNum != colNum:
        raise Exception('Covariance matrix must be symmetric')
        
    for i in xrange(rowNum):
        sheet.range(startCol+str(startRow+i)).value = data[i:i+1,:i+1]
    
def getStaticMatrics(factors, maturityList):
    '''get static matrics, i.e. 
    observationMatrices
    transitionOffsets
    observationOffsets
    '''
    
    observationMatrices = [[x**i for i in xrange(len(factors))] for x in maturityList]
    
    #transition offsets
    transitionOffsets = [0] * len(factors)
    
    #observation offsets
    observationOffsets = [0] * len(maturityList)
    
    return observationMatrices, observationOffsets, transitionOffsets
    
def getMaturityList(df):
    '''parse maturity string'''
    
    return [  float(col[:-2])/(1 if col[-2:]=='YR' else 12)  for col in df.columns.tolist()]


def getKalmanFilter(df, **kargs):
    '''get Kalman filter object for yield curve dynamic'''
    
    observationMatrices, observationOffsets, transitionOffsets = \
    getStaticMatrics(FACTORS, getMaturityList(df))
    
    
    return KalmanFilter(n_dim_obs=len(df.columns), observation_matrices=observationMatrices, 
                                transition_offsets=transitionOffsets, 
                                observation_offsets=observationOffsets, **kargs)    


def trainKalmanFilterFromHistoricalData():
    '''train kalman filter from historaical data'''
    
    df = getMeasurementDataFromExcelSheet()
    
    #initialize a Kalman Filter object
    kf = getKalmanFilter(df)
    
    #fitting parameters with EM algorithms
    kf = kf.em(dfToMaskedArray(df), 
               n_iter=int(getConfig().get('Number of iteration', 5)), 
               em_vars=['transition_matrices', 
                        'transition_covariance', 
                        'observation_covariance'
                       ])
    
    
    for data, pos, isCov in [(kf.transition_matrices, 'C18', False),
                       (kf.transition_covariance, 'H18', True),
                       (kf.observation_covariance, 'B3', True)]:
        
        storeDataToExcel(data, 
                     PARAMETER_SHEET_NAME, 
                     startPosition=pos,
                     bookName=EXCEL_FILE_NAME, clear=False,
                     isCovariance=isCov)
    
def getDateList():
    '''get date list available in data tab'''
    book = xw.Book(EXCEL_FILE_NAME)
    sheet = getSheetFromBook(book, DATA_SHEET_NAME)
    
    res = sheet.range('A2').expand().value
    return [x[0] if isinstance(x,list) else x for x in res] if res else []
    
    
def getMeasurementDataFromExcelSheet():
    '''get measurement from Excel'''
    
    book = xw.Book(EXCEL_FILE_NAME)
    sheet = getSheetFromBook(book, DATA_SHEET_NAME)
    dateList = sheet.range('A2').expand().value
    value = sheet.range('B1:L{}'.format(1+len(dateList))).value
    return pd.DataFrame(value[1:],
                      columns =value[0],
                      index=dateList)
    
def getKalmanFilterParametersFromExcel():
    '''get the Kalman Filter parameters from Excel, i.e.
       transition_matrices, transition_covariance, observation_covariance
    '''
    
    book = xw.Book(EXCEL_FILE_NAME)
    sheet = getSheetFromBook(book, PARAMETER_SHEET_NAME)
    
    transition_matrices=sheet.range('C18:E20').value
    transition_covariance=sheet.range('H18:J20').value
    observation_covariance=sheet.range('B3:L13').value
    
    return transition_matrices, transition_covariance, observation_covariance
    
def applyKalmanFilter():
    '''get filtered data and smoothed data'''
    df = getMeasurementDataFromExcelSheet()
    
    transition_matrices, transition_covariance, observation_covariance = \
    getKalmanFilterParametersFromExcel()
    
    kf = getKalmanFilter(df,
                             transition_matrices=transition_matrices,
                             transition_covariance=transition_covariance,
                             observation_covariance=observation_covariance)
    
    measurement = dfToMaskedArray(df)
    filtered_state_means, filtered_state_covariances = kf.filter(measurement)
    smoothed_state_means, smoothed_state_covariances = kf.smooth(measurement)
    
    startDate = getConfig().get('Start Date', dt.datetime(2016, 3, 1, 0, 0))
    
    for data, sheetName in [(filtered_state_means, FILTER_SHEET_NAME),
                            (smoothed_state_means, SMOOTH_SHEET_NAME)]:
    
        storeDataToExcel(pd.DataFrame(filtered_state_means, 
                                      columns=FACTORS,
                                      index=df.index).loc[startDate:], 
                         sheetName, 
                         startPosition='A1',
                         bookName=EXCEL_FILE_NAME, clear=True)
    
    
if __name__ == "__main__":
    # Used for frozen executable
    pass#SSapplyKalmanFilter()