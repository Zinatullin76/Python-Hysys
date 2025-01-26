import os 
import win32com.client as win32  


def HysysBridge(TableName):
    
    """Connect between Hysys SpreadSheet an Python script

    Args:
        TableName (str): SpreadSheet's name 

    Returns:
        win32 object: Contains data from SpreadSheet
    """
    App = win32.Dispatch('HYSYS.Application')
    Case = App.ActiveDocument
    TableSheet = Case.Flowsheet.Operations.Item(TableName)
    Solver = Case.Solver

    return TableSheet, Solver


def example(Flowname):
    
    """Connect between Hysys SpreadSheet an Python script

    Args:
        TableName (str): SpreadSheet's name 

    Returns:
        win32 object: Contains data from SpreadSheet
    """
    App = win32.Dispatch('HYSYS.Application')
    Case = App.ActiveDocument
    Item = Case.Flowsheet.MaterialStreams.Item(Flowname)

    return Item