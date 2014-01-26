##This small program will allow the user to check the postal status after the first call...

import pandas as pd
import numpy as np
import re
import datetime

def clear_cell():
    return

CL = pd.read_csv(r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ContactList\ContactList_Wave3.csv", index_col = 'UID')
row = 6
if Cell(4, 2).is_empty():
    for i in CL.index:
        if i != 'TEST' and CL['send_docs'][i] == 1 and pd.isnull(CL['docs_sent'][i]):
                Cell(row, 2).value = i
                Cell(row, 3).value = CL['Fact_name'][i]
                Cell(row, 4).value = CL['address'][i]
                Cell(row, 5).value = CL['module'][i]
                if not pd.isnull(CL['Cnam_upd'][i]):
                    Cell(row, 6).value = CL['Cnam_upd'][i]
                elif not pd.isnull(CL['Cnam_orig'][i]):
                    Cell(row, 6).value = CL['Cnam_orig'][i]
                if not pd.isnull(CL['status'][i]):
                    Cell(row, 7).value = CL['status'][i]
                row +=1


row = 6
Error = 0
DT = str(datetime.datetime.today())
Date = DT.replace(" ", "_")
if not Cell(6, 2).is_empty() and Cell(3, 4).value == 'Y':
    while not Cell(row, 2).is_empty():
        if not Cell(row, 8).value == 'Y' and not Cell(row, 8).is_empty():
            Cell(row, 10).value = "INPUT ERROR: This cell may either be empty or eqaul to Y"
            Error +=1
        row +=1

    if Error == 0:
        Cell(4, 8).value = clear_cell()
        row = 6
        while not Cell(row, 2).is_empty():
            if Cell(row, 8).value == 'Y':
                CL['docs_sent'][Cell(row, 2).value] = Date
            for x in xrange(2, 11):
                Cell(row, x).value = clear_cell()
            row +=1
        Cell(3, 4).value = clear_cell()
        CL.to_csv(r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ContactList\ContactList_Wave3.csv")
    else:
        Cell(4, 8).value = "YOU HAVE ERRORS"

