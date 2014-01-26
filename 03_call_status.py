import pandas as pd
import numpy as np
import re
import datetime

def clear_cell():
    return

CL = pd.read_csv(r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ContactList\ContactList_Wave1.csv", index_col = 'UID')

## Insert status of First Call
row = 6
for UID in CL.index:
    if CL['attempts_fc'][UID] == 0 and CL['first_call'][UID] != 1 and UID != 'TEST':
        Cell("FirstCall", row, 2).value = UID
        Cell("FirstCall", row, 3).value = CL['attempts_fc'][UID]
        Cell("FirstCall", row, 4).value = CL['Fact_name'][UID]
        Cell("FirstCall", row, 5).value = CL['address'][UID]
        if not pd.isnull(CL['Cnum_orig_fact'][UID]):
            Cell("FirstCall", row, 6).value = CL['Cnum_orig_fact'][UID]
        if not pd.isnull(CL['Cnum_orig_off'][UID]):
            Cell("FirstCall", row, 7).value = CL['Cnum_orig_off'][UID]
        if not pd.isnull(CL['Cnum_upd'][UID]):
            Cell("FirstCall", row, 8).value = CL['Cnum_upd'][UID]
        Cell("FirstCall", row, 9).value = CL['module'][UID]
        Cell("FirstCall", row, 10).value = CL['price'][UID]
        if not pd.isnull(CL['status'][UID]):
            Cell("FirstCall", row, 11).value = CL['status'][UID]
        row +=1


row +=1

Cell("FirstCall", row, 1).copy_from(Cell("FirstCall", 3, 1))
Cell("FirstCall", row, 1).value = "INCOMPLETE FIRST CALLS"
row +=1
CellRange("FirstCall", (row, 2), (row, 12)).copy_from(CellRange("FirstCall", (5, 2), (5, 12)))
row+=1

for UID in CL.index:
    if CL['attempts_fc'][UID] != 0 and CL['first_call'][UID] != 1 and UID != 'TEST':
        Cell("FirstCall", row, 2).value = UID
        Cell("FirstCall", row, 3).value = CL['attempts_fc'][UID]
        Cell("FirstCall", row, 4).value = CL['Fact_name'][UID]
        Cell("FirstCall", row, 5).value = CL['address'][UID]
        if not pd.isnull(CL['Cnum_orig_fact'][UID]):
            Cell("FirstCall", row, 6).value = CL['Cnum_orig_fact'][UID]
        if not pd.isnull(CL['Cnum_orig_off'][UID]):
            Cell("FirstCall", row, 7).value = CL['Cnum_orig_off'][UID]
        if not pd.isnull(CL['Cnum_upd'][UID]):
            Cell("FirstCall", row, 8).value = CL['Cnum_upd'][UID]
        Cell("FirstCall", row, 9).value = CL['module'][UID]
        Cell("FirstCall", row, 10).value = CL['price'][UID]
        if not pd.isnull(CL['status'][UID]):
            Cell("FirstCall", row, 11).value = CL['status'][UID]
        row +=1


## Insert status of Follow Up One
row = 6
for UID in CL.index:
    if CL['attempts_fu1'][UID] == 0 and CL['fu1_call'][UID] != 1 and CL['send_docs'][UID] == 1 and UID != 'TEST':
        Cell("FollowUp1", row, 2).value = UID
        Cell("FollowUp1", row, 3).value = CL['attempts_fu1'][UID]
        if not pd.isnull(CL['docs_sent'][UID]):
            Cell("FollowUp1", row, 4).value = str(CL['docs_sent'][UID])
        Cell("FollowUp1", row, 5).value = CL['Fact_name'][UID]
        if not pd.isnull(CL['Cnum_orig_fact'][UID]):
            Cell("FollowUp1", row, 6).value = CL['Cnum_orig_fact'][UID]
        if not pd.isnull(CL['Cnum_orig_off'][UID]):
            Cell("FollowUp1", row, 7).value = CL['Cnum_orig_off'][UID]
        if not pd.isnull(CL['Cnum_upd'][UID]):
            Cell("FollowUp1", row, 8).value = CL['Cnum_upd'][UID]
        Cell("FollowUp1", row, 9).value = CL['module'][UID]
        Cell("FollowUp1", row, 10).value = CL['price'][UID]
        if not pd.isnull(CL['status'][UID]):
            Cell("FollowUp1", row, 11).value = CL['status'][UID]
        row +=1


row +=1
Cell("FollowUp1", row, 1).copy_from(Cell("FollowUp1", 3, 1))
Cell("FollowUp1", row, 1).value = "INCOMPLETE FOLLOW UP 1 CALLS"
row +=1
CellRange("FollowUp1", (row, 2), (row, 12)).copy_from(CellRange("FollowUp1", (5, 2), (5, 12)))
row+=1

for UID in CL.index:
    if CL['attempts_fu1'][UID] != 0 and CL['fu1_call'][UID] != 1 and UID != 'TEST':
        Cell("FollowUp1", row, 2).value = UID
        Cell("FollowUp1", row, 3).value = CL['attempts_fu1'][UID]
        if not pd.isnull(CL['docs_sent'][UID]):
            Cell("FollowUp1", row, 4).value = str(CL['docs_sent'][UID])
        Cell("FollowUp1", row, 5).value = CL['Fact_name'][UID]
        if not pd.isnull(CL['Cnum_orig_fact'][UID]):
            Cell("FollowUp1", row, 6).value = CL['Cnum_orig_fact'][UID]
        if not pd.isnull(CL['Cnum_orig_off'][UID]):
            Cell("FollowUp1", row, 7).value = CL['Cnum_orig_off'][UID]
        if not pd.isnull(CL['Cnum_upd'][UID]):
            Cell("FollowUp1", row, 8).value = CL['Cnum_upd'][UID]
        Cell("FollowUp1", row, 9).value = CL['module'][UID]
        Cell("FollowUp1", row, 10).value = CL['price'][UID]
        if not pd.isnull(CL['status'][UID]):
            Cell("FollowUp1", row, 11).value = CL['status'][UID]
        row +=1

## Insert status of Second Follow Up
row = 6
for UID in CL.index:
    if CL['attempts_fu2'][UID] == 0 and CL['fu2_call'][UID] != 1 and CL['free_mod'][UID] == 1 and UID != 'TEST':
        print CL['free_mod'][UID] 
        Cell("SecondFollowUp", row, 2).value = UID
        Cell("SecondFollowUp", row, 3).value = CL['attempts_fu2'][UID]
        if not pd.isnull(CL['docs_sent'][UID]):
            Cell("SecondFollowUp", row, 4).value = str(CL['docs_sent'][UID])
        Cell("SecondFollowUp", row, 5).value = CL['Fact_name'][UID]
        if not pd.isnull(CL['Cnum_orig_fact'][UID]):
            Cell("SecondFollowUp", row, 6).value = CL['Cnum_orig_fact'][UID]
        if not pd.isnull(CL['Cnum_orig_off'][UID]):
            Cell("SecondFollowUp", row, 7).value = CL['Cnum_orig_off'][UID]
        if not pd.isnull(CL['Cnum_upd'][UID]):
            Cell("SecondFollowUp", row, 8).value = CL['Cnum_upd'][UID]
        Cell("SecondFollowUp", row, 9).value = CL['module'][UID]
        Cell("SecondFollowUp", row, 10).value = CL['price'][UID]
        if not pd.isnull(CL['status'][UID]):
            Cell("SecondFollowUp", row, 11).value = CL['status'][UID]
        row +=1


row +=1
Cell("SecondFollowUp", row, 1).copy_from(Cell("SecondFollowUp", 3, 1))
Cell("SecondFollowUp", row, 1).value = "INCOMPLETE SECOND FOLLOW UP PROCESS"
row +=1
CellRange("SecondFollowUp", (row, 2), (row, 12)).copy_from(CellRange("SecondFollowUp", (5, 2), (5, 12)))
row+=1

for UID in CL.index:
    if CL['attempts_fu2'][UID] != 0 and CL['fu2_call'][UID] != 1 and UID != 'TEST':
        Cell("SecondFollowUp", row, 2).value = UID
        Cell("SecondFollowUp", row, 3).value = CL['attempts_fu2'][UID]
        if not pd.isnull(CL['docs_sent'][UID]):
            Cell("SecondFollowUp", row, 4).value = str(CL['docs_sent'][UID])
        Cell("SecondFollowUp", row, 5).value = CL['Fact_name'][UID]
        if not pd.isnull(CL['Cnum_orig_fact'][UID]):
            Cell("SecondFollowUp", row, 6).value = CL['Cnum_orig_fact'][UID]
        if not pd.isnull(CL['Cnum_orig_off'][UID]):
            Cell("SecondFollowUp", row, 7).value = CL['Cnum_orig_off'][UID]
        if not pd.isnull(CL['Cnum_upd'][UID]):
            Cell("SecondFollowUp", row, 8).value = CL['Cnum_upd'][UID]
        Cell("SecondFollowUp", row, 9).value = CL['module'][UID]
        Cell("SecondFollowUp", row, 10).value = CL['price'][UID]
        if not pd.isnull(CL['status'][UID]):
            Cell("SecondFollowUp", row, 11).value = CL['status'][UID]
        row +=1
