import pandas as pd
import numpy as np
import re
import datetime

#######################################################################################################################################################################
#####                                                                                                                                                             #####
#####****************************************                         DEFINING THE INTERACTION CLASS                       ***************************************#####                                                                                                                   #####
#####                                                                                                                                                             #####
#######################################################################################################################################################################

class InterAction:

    def __init__(self) :
        """This reads two csv files into pandas:
                PD_RO: a csv with UIDs, Names, Prices and Modules. It is called RO as it is Read Only. It will never be altered. NOTE: The UIDs are
                       imported to pandas as the index values.
                PD_CL: a csv with UIDs, Names, Contact Details, Status Variables, and Comments. It is called CL as it is a contact list.
                       This csv will be directly modifiable by the program according to user input. Note: the UIDs are imported to pandas as index values."""

        self.PD_RO = pd.read_csv(r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\FactoryList\Wave1\FactoryList.csv", index_col = 'UID')
        self.PD_CL = pd.read_csv(r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ContactList\ContactList_Wave1.csv", index_col = 'UID')
        self.DF = pd.load(r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\DataFrames\Wave1\FollowUp1DataFrame")


    def CheckRO_UID(self, CellVal):
        """Tests to see if user input UID is in the PD_RO index values, and returns a statement to that effect"""
        self.UID = str(CellVal)
        if self.UID in self.PD_RO.index.values:
            return 'TRUE'
        else:
            return "ERROR, the Factory Code you have entered is not recignised"

    def CheckCL_UID(self):
        """Tests to see if user input UID is in the PD_CL index values, and returns a statement to that effect"""
        if self.UID in self.PD_CL.index.values:
            return 'TRUE'
        else:
            return "ERROR, the Factory Code you have entered is not in the Contact List"

    def GetUIDDetails(self):
        """If the user entered UID is in the PD_RO index then a list is returned that contains information on the factory identified by that UID. If the user entered
           UID is not in the RO index then a list of empty strings is returned"""
        UIDDetails = []
        if self.UID in self.PD_RO.index.values:
            UIDDetails.append(self.PD_RO['name'][self.UID])
            UIDDetails.append(self.PD_RO['price'][self.UID])
            UIDDetails.append(self.PD_RO['module'][self.UID])
        else:
            UIDDetails.append("")
            UIDDetails.append("")
            UIDDetails.append("")

        return UIDDetails

    def GetCLDetails(self):
        """If the user entered UID is in the PD_CL index then a list is returned that contains information on the factory identified by that UID. If the UID is
           not in the CL index then a list of empty strings is returned."""

        CLCols = ['Cnum_orig_fact', 'Cnum_orig_off', 'Cnum_gm', 'Cnum_upd', 'Cnam_orig', 'Cnam_gm', 'Cnam_upd', 'Cdesig_orig', 'Cdesig_gm', 'Cdesig_upd', \
                  'Cdesig_code', 'Cemail', 'address', 'status']
        CLDetails = []

        if self.UID in self.PD_CL.index.values:
            for elem in CLCols:
                if pd.isnull(self.PD_CL[elem][self.UID]):
                    CLDetails.append("")
                else:
                    CLDetails.append(self.PD_CL[elem][self.UID])
        else:
            for elem in CLCols:
                CLDetails.append("")

        return CLDetails


    def UpdateStatus(self, Cell_status):
        """if the user has input an status update then this is entered into the PD_CL dataframe"""
        if not Cell_status.is_empty():
            self.PD_CL['status'][self.UID] = str(Cell_status.value)

    def UpdateAttempts(self, Cell_attempts, attempts_col):
        """For each time the data program is run (exporting data) the number of attempts will increase by one in the PD_CL data frame. This is so that the
           user can keep track of how many attempted calls he has made"""
        if not Cell_attempts.is_empty():
            self.PD_CL[attempts_col][self.UID] = self.PD_CL[attempts_col][self.UID] + 1

    def UpdateTerminate(self, Cell_terminate, terminate_col):
        """If the user has entered Y into the call process terminate cell, then the PD_CL data frame is updated. This is so the user can know when
           call process is over for a given factory."""
        if Cell_terminate.value == 'Y':
            self.PD_CL[terminate_col][self.UID] = 1

    def UpdateContactDetails(self, Cell_GMNum, Cell_CNum, Cell_GMNam, Cell_CNam, Cell_GMDesig, Cell_CDesig, Cell_CDesigCode, Cell_Email, Cell_Address):
        """If user enters information in the refernced cells in order to update the contact list, then the informaiton is imported to pandas and then
           exported to the contact list CSV. The code would be very easy to combine into a for loop with list of Cells, to simplify the code. Consider doing
           this at some stage"""

        Empty = 0

        if not Cell_GMNum.is_empty():
             self.PD_CL['Cnum_gm'][self.UID] = str(Cell_GMNum.value)
             Empty +=1

        if not Cell_CNum.is_empty():
             self.PD_CL['Cnum_upd'][self.UID] = str(Cell_CNum.value)
             Empty +=1

        if not Cell_GMNam.is_empty():
             self.PD_CL['Cnam_gm'][self.UID] = str(Cell_GMNam.value)
             Empty +=1

        if not Cell_CNam.is_empty():
             self.PD_CL['Cnam_upd'][self.UID] = str(Cell_CNam.value)
             Empty +=1

        if not Cell_GMDesig.is_empty():
            self.PD_CL['Cdesig_gm'][self.UID] = str(Cell_GMDesig.value)
            Empty +=1

        if not Cell_CDesig.is_empty():
             self.PD_CL['Cdesig_upd'][self.UID] = str(Cell_CDesig.value)
             Empty +=1

        if not Cell_CDesigCode.is_empty():
             self.PD_CL['Cdesig_code'][self.UID] = str(Cell_CDesigCode.value)
             Empty +=1

        if not Cell_Email.is_empty():
             self.PD_CL['Cemail'][self.UID] = str(Cell_Email.value)
             Empty +=1

        if not Cell_Address.is_empty():
             self.PD_CL['address'][self.UID] = str(Cell_Address.value)
             Empty +=1

        if Empty > 0:
            Cell_GMNum.clear()
            Cell_CNum.clear()
            Cell_GMNam.clear()
            Cell_CNam.clear()
            Cell_GMDesig.clear()
            Cell_CDesig.clear()
            Cell_CDesigCode.clear()
            Cell_Email.clear()
            Cell_Address.clear()

    def InsertSinglePoint(self, CellRef, ColRef, Test, Value):
        """Ascertains is a cell value is eqaul to a test, and if it is then is updates the value in the FollowUp1 DataFrame in the ColRef indicated"""
        if CellRef.value == Test:
            self.DF[ColRef][self.UID] = Value

    def InsertSinglePoint_MA(self, CellRef, ColRef, Test1, Test2, Value1, Value2):
        """As above but with a two tailed test, and two possible values"""
        if CellRef.value == Test1:
            self.DF[ColRef][self.UID] = Value1
        elif CellRef.value == Test2:
            self.DF[ColRef][self.UID] = Value2

    def InsertSinglePoint_NE(self, CellRef, ColRef, Value, Type):
        """As above but the test is now whether the cell is empty"""
        if not CellRef.is_empty():
            self.DF[ColRef][self.UID] = Type(Value)

    def UpdateInsertSinglePoint_NE(self, CellRef, ColRef, Value, Type):
        """As above but the test is now whether the cell is empty"""
        if not CellRef.is_empty():
            if self.DF[ColRef][self.UID] == 'nan':
                self.DF[ColRef][self.UID] = Type(Value)
            else:
                self.DF[ColRef][self.UID] = self.DF[ColRef][self.UID] + '+' + Type(Value)

    def InsertList(self, CellRef, ColRef, Type):
        """If the CellRef is not empty then the contents are converted to a list of integers and appended to the DataFrame in the column indiated"""
        if not CellRef.is_empty():
            Vals = CellRef.value
            ValsList = [Type(i) for i in Vals.split(',')]
            self.DF[ColRef][self.UID].append(ValsList)

    def IncreaseByOne(self, CellRef, ColRef):
        """If a CellRef value is equal to the Test then the FollowUp1 DataFrame value at the column indicated is increased by 1"""
        if not CellRef.is_empty():
            self.DF[ColRef][self.UID] = self.DF[ColRef][self.UID] + 1

    def ExportToCSV(self):
        """Exports the updates PD_CL back to the Excel CSV."""
        self.PD_CL.to_csv(r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ContactList\ContactList_Wave1.csv")

    def ExportDF(self):
        """Saves the FirsCall dataframe"""
        pd.DataFrame.save(self.DF, r"C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\DataFrames\Wave1\FollowUp1DataFrame")


#######################################################################################################################################################################
#####                                                                                                                                                             #####
#####****************************************                           DEFINING INPUT CHECKING FUCNTIONS                  ***************************************#####                                                                                                                   #####
#####                                                                                                                                                             #####
#######################################################################################################################################################################


def clear_messages():
    """When Cell.value = clear_messages() the cell value is cleared without the formatting changing. Therefore this is preferred to .clear() method"""
    return


def check_inputsYN(CellRef):
    """Checks whether a cell is equal to Y or N and returns an error message if it does not."""
    if CellRef.value == "Y" or CellRef.value == "N":
        return
    else:
        return "INPUT ERROR: Input cell can only contain the following characters: Y N"

def check_inputsYNU(CellRef):
    """Checks whether a cell is equal to Y or N and returns an error message if it does not."""
    if CellRef.value == "Y" or CellRef.value == "N" or CellRef.value == 'U':
        return
    else:
        return "INPUT ERROR: Input cell can only contain the following characters: Y N U"


def CodeInputCheck(CellRef, PermittedList):
    """Checks whether the contents of a user input cell which should contain a list of codes is valid (by checking against the PermitedList. If a code is used outside
       of permitted range a message is returned to that effect. If a character is used which prevents the cell being analysed as a list of integers, then the exception
       raised by the function is captured and returned to the user"""

    CodeVal = CellRef.value
    BadList = []
    try:
        IntCodeValList = [int(i) for i in CodeVal.split(',')]
        for elem in IntCodeValList:
            if elem not in PermittedList:
                BadList.append(elem)
        if len(BadList) > 0:
            return "INPUT ERROR: The following value(s) are not permitted: " + str(BadList)
        else:
            return
    except ValueError, Argument:
        return "INPUT ERROR: Invalid character used : " + str(Argument)



def CheckLogics(KeyCellRef, CellRef, KeyVal1, KeyVal2, Message1, Message2, Message3):
    if KeyCellRef.value == KeyVal1 and CellRef.is_empty():
        return Message1
    elif KeyCellRef.value == KeyVal1 and not CellRef.is_empty():
        return
    elif KeyCellRef.value == KeyVal2 and not CellRef.is_empty():
        return Message2
    elif KeyCellRef.value == KeyVal2 and CellRef.is_empty():
        return
    elif KeyCellRef.is_empty() and not CellRef.is_empty():
        return Message3
    elif KeyCellRef.is_empty() and CellRef.is_empty():
        return



def CheckCodeLogic(KeyCellRef, CellRef, KeyVal, Message1, Message2):
    if KeyCellRef.is_empty():
        if not CellRef.is_empty():
            return Message1
        else:
            return
    else:
        try:
            KeyString = KeyCellRef.value
            KeyInts = [int(i) for i in KeyString.split(',')]
            if KeyVal in KeyInts:
                if CellRef.is_empty():
                    return Message2
                else:
                    return
            else:
                if not CellRef.is_empty():
                    return Message1

        except ValueError:
            return


def Confirm_error_free(RowList, Col):
    Errors = 0
    ErrorsList = []
    for elem in RowList:
        if not Cell(elem, Col).is_empty():
            Errors +=1
            ErrorsList.append(elem)
    if Errors == 0:
        return "TRUE"
    else:
        return "FALSE: errors in row(s) " + str(ErrorsList)

#######################################################################################################################################################################
#######################################################################################################################################################################
#######################################################################################################################################################################



#######################################################################################################################################################################
#********* Interaction Program *********#

Inter = InterAction()
Cell("FollowUp1", 5, 3).value = Inter.CheckRO_UID(Cell(4, 3).value)
Cell("FollowUp1", 6, 3).value = Inter.CheckCL_UID()

UIDDetails = Inter.GetUIDDetails()
row = 7
for elem in UIDDetails:
    Cell("FollowUp1", row, 3).value = elem
    row+=1



CLDetails = Inter.GetCLDetails()
row = 3

Cell("FollowUp1", 11, 3).value = CLDetails.pop()
Cell("FollowUp1", 10, 3).value = CLDetails.pop()
for elem in CLDetails:
    Cell("FollowUp1", row, 7).value = elem
    row+=1
#######################################################################################################################################################################
#********* Check Logic/Inputs Program *********#

if Cell(37, 6).value == 'Y':

## Clear the error messages
    AllRows = [22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 33, 34, 35]
    for elem in AllRows:
        Cell(elem, 11).value = clear_messages()

## Check first cell as this must be filled in no matter what happens.
    Cell(22, 11).value = check_inputsYN(Cell(22, 6))

## Check the logics
    Cell(34, 11).value = CheckLogics(Cell(23, 6), Cell(34, 6), 'Y', 'N', 'LOGIC ERROR: Q2.1 = Y so this cell should not be empty', \
    'LOGIC ERROR: Q2.1 = N so this cell must be empty', '')
    Cell(30, 11).value = CheckLogics(Cell(29, 6), Cell(30, 6), 'N', 'Y', 'LOGIC ERROR: Q4.1 = N so this cell should not be empty', \
    'LOGIC ERROR: Q4.1 = Y so this cell must be empty', 'LOGIC ERROR: Q4.1 is empty so this cell must also be empty')
    Cell(31, 11).value = CheckLogics(Cell(29, 6), Cell(31, 6), 'N', 'Y', 'LOGIC ERROR: Q4.1 = N so this cell should not be empty', \
    'LOGIC ERROR: Q4.1 = Y so this cell must be empty', 'LOGIC ERROR: Q4.1 is empty so this cell must also be empty')

    RowList1 = [27, 28, 29]
    for elem in RowList1:
        Cell(elem, 11).value = CheckLogics(Cell(23, 6), Cell(elem, 6), 'Y', 'N', 'LOGIC ERROR: Q2.1 = Y so this cell should not be empty', \
        'LOGIC ERROR: Q2.1 = N so this cell must be empty', 'LOGIC ERROR: Q2.1 is empty so this cell must also be empty')

    RowList2 = [25, 26]
    for elem in RowList2:
        Cell(elem, 11).value = CheckCodeLogic(Cell(24, 6), Cell(elem, 6), -99, "LOGIC ERROR: Q2.2 does not equal Not Interested Code, so this cell should be empty", "LOGIC ERROR: Q2.2 = Not Interested Code, so this cell may not be empty")

    Cell(24, 11).value = CheckLogics(Cell(23, 6), Cell(24, 6), 'N', 'Y', 'LOGIC ERROR: Q2.1 = N so this cell should not be empty', \
        'LOGIC ERROR: Q2.1 = Y so this cell must be empty', 'LOGIC ERROR: Q2.1 is empty so this cell must also be empty')

    Cell(23, 11).value = CheckLogics(Cell(22, 6), Cell(23, 6), 'Y', 'N', 'LOGIC ERROR: Q1 = Y so this cell should not be empty', \
    'LOGIC ERROR: Q1 = N so this cell must be empty', '')

## Check the inputs
    RowList3 = [23, 33]
    for elem in RowList3:
        if not Cell(elem, 6).is_empty():
            if Cell(elem, 11).is_empty():
                Cell(elem, 11).value = check_inputsYN(Cell(elem, 6))
            elif not re.match("LOGIC", Cell(elem, 11).value):
                Cell(elem, 11).value = check_inputsYN(Cell(elem, 6))

    if not Cell(29, 6).is_empty():
        if Cell(29, 11).is_empty():
            Cell(29, 11).value = check_inputsYNU(Cell(29, 6))
        elif not re.match("LOGIC", Cell(29, 11).value):
            Cell(29, 11).value - check_inputsYNU(Cell(29, 6))

    if not Cell(24, 6).is_empty():
        if Cell(24, 11).is_empty():
            Cell(24, 11).value = CodeInputCheck(Cell(24, 6), [-99, -55, 1, 2, 3, 4])
        elif not re.match("LOGIC", Cell(24, 11).value):
            Cell(24, 11).value = CodeInputCheck(Cell(24, 6), [-99, -55, 1, 2, 3, 4])

    if not Cell(25, 6).is_empty():
        if Cell(25, 11).is_empty():
            Cell(25, 11).value = CodeInputCheck(Cell(25, 6), [-55, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26])
        elif not re.match("LOGIC", Cell(25, 11).value):
            Cell(25, 11).value = CodeInputCheck(Cell(25, 6), [-55, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26])

    if not Cell(27, 6).is_empty():
        if Cell(27, 11).is_empty():
            Cell(27, 11).value = CodeInputCheck(Cell(27, 6), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
        elif not re.match("LOGIC", Cell(27, 11).value):
            Cell(27, 11).value = CodeInputCheck(Cell(27, 6), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])

    if not Cell(28, 6).is_empty():
        if Cell(28, 11).is_empty():
            Cell(28, 11).value = CodeInputCheck(Cell(28, 6), [-88, -55, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18])
        elif not re.match("LOGIC", Cell(28, 11).value):
            Cell(28, 11).value = CodeInputCheck(Cell(28, 6), [-88, -55, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18])

    if not Cell(30, 6).is_empty():
        if Cell(30, 11).is_empty():
            Cell(30, 11).value = CodeInputCheck(Cell(30, 6), [-55, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26])
        elif not re.match("LOGIC", Cell(30, 11).value):
            Cell(30, 11).value = CodeInputCheck(Cell(30, 6), [-55, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26])

    if Cell(33, 6).is_empty():
        Cell(33, 11).value = "INPUT ERROR: This cell must not be left blank"
    else:
        Cell(33, 11).value = check_inputsYN(Cell(33, 6))

    if Cell(35, 6).is_empty():
        Cell(35, 11).value = "Are you sure you don't want to update the status?"

##Check Data is Error Free
    Cell(37, 11).value = Confirm_error_free([22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 33, 34], 11)


#######################################################################################################################################################################
#********* Export/Save Data Program *********#


if Cell(38, 6).value == 'Y' and Cell(37, 11).value == "TRUE" and Cell(5, 3).value == "TRUE":
    print 'I Got here'
## save a raw version of excel file for security reasons
    D = str(datetime.datetime.now())
    D1 = D.replace(' ','_')
    D2 = D1.replace(":", "_")
    D3 = D2[0:D2.index('.')]
    save_copy(r'C:\Users\IPAB\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\RawData\FollowUp1\Wave1\FC_' + str(Inter.UID) + '_' + D3 + '.xls')

## export the data to the FollowUp1 DataFrame
    Inter.IncreaseByOne(Cell(22, 6), 'no_attempts')
    Inter.InsertSinglePoint(Cell(23, 6), 'Discuss_Dummy', 'Y', 1)
    Inter.InsertList(Cell(24, 6), 'nDiscuss_code', int)
    Inter.InsertList(Cell(25, 6), 'nInterest_code', int)
    Inter.InsertSinglePoint_NE(Cell(26, 6), 'nInterest_str', Cell(26, 6).value, str)
    Inter.InsertList(Cell(27, 6), 'Discuss_p', int)
    Inter.InsertList(Cell(28, 6), 'Qs_concerns', int)
    Inter.InsertSinglePoint_MA(Cell(29, 6), 'free_dummy', 'Y', 'N', 1, 0)
    Inter.InsertList(Cell(30, 6), 'nFree_code', int)
    Inter.InsertSinglePoint_NE(Cell(31, 6), 'nFree_str', Cell(31, 6).value, str)
    Inter.UpdateInsertSinglePoint_NE(Cell(34, 6), 'Comments', Cell(34, 6).value, str)
    Inter.ExportDF()

    ## Update contact list information
    Inter.UpdateStatus(Cell(35, 6))
    Inter.UpdateAttempts(Cell(22, 6), 'attempts_fu1')
    Inter.UpdateTerminate(Cell(33, 6), 'fu1_call')
    Inter.UpdateContactDetails(Cell("FollowUp1", 4, 12), Cell("FollowUp1",5, 12), Cell("FollowUp1",6, 12), Cell("FollowUp1",7, 12), \
        Cell("FollowUp1", 8, 12), Cell("FollowUp1",9, 12), Cell("FollowUp1", 10, 12), Cell("FollowUp1", 11, 12), Cell("FollowUp1", 12, 12))
    Inter.ExportToCSV()

    ## clear the cells ready for next use
    for x in xrange(4, 15):
        Cell(x, 3).value = clear_messages()
        Cell(x, 12).value = clear_messages()
    for x in xrange(3, 15):
        Cell(x, 7).value = clear_messages()
    for x in xrange(22, 39):
        Cell(x, 6).value = clear_messages()
        Cell(x, 11).value = clear_messages()
