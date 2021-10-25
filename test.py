# Communication with windows API devices
import comtypes
from comtypes.client import CreateObject
from comtypes.client import GetActiveObject

import os # For compiling file paths
import time

if __name__ == "__main__":
    '''Commands to ecexute when run as standalone file'''
    # Open TNMR application / get reference to it
    try:
        A = GetActiveObject('NTNMR.Application')
    except OSError:
        A = CreateObject('NTNMR.Application')
        print('Opening new TNMR window')
        # Closes the empty created document
        if A.CloseFile(''): pass
        else: print('Failed to close file')
        

    file_path = 'C:\\TNMR\\data\\Nejc test\\Cu NQR\\'
    file_name = 'test'

    print(A.GetDocumentList)
    print(A.GetActiveDocPath)
    # Create new TNMR document
    B = CreateObject('NTNMR.Document')
    B.SetComment('Testing file creation and manipulation')
    B.GetComment
    B.LoadSequence(file_path + 'T1.tps')
    # Set correct parameters
    B.SetNMRParameter('Observe Freq.', '25.98MHz')
    B.SetNMRParameter('Dwell Time', '200n')
    B.SetNMRParameter('Receiver Gain', '30')
    B.SetNMRParameter('Receiver Phase', '122') # Guess it changes
    B.SetNMRParameter('Scans 1D', '256')
    B.SetNMRParameter('Points 2D', '20')
    #B.ZG()
    #time.sleep(10)
    #if B.Abort: print('Aborting')
    #else: print('Failed to abort!')

    x = B.GetTableList
    print(x)
    print(type(x))
    table_list = x.split(',')
    print(table_list)

    for table in table_list:
        tmp = B.GetTable(table)
        print(tmp, type(tmp))

    
    B.GetDataSize
    x = B.GetData

    B.SaveAs(file_path + file_name + '.tnt')
    B.Export(file_path + file_name + '.txt', 0) # 0 for txt file
    # Unfortunately doesnt save the support information. Only raw columns

    print(A.GetDocumentList)
    print(A.GetActiveDocPath)

    #print(B.CheckAcquisition)
    #print(B.ZG)
    #print(B)
