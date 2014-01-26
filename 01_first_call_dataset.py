import pandas as pd
import numpy as np

PD_RO = pd.read_csv(r"C:\Users\rcreedon\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\FactoryList\Wave1\FactoryList.csv", index_col = 'UID')
UIDList = [i for i in PD_RO.index]
FirstCallData = pd.DataFrame(index = UIDList)
FirstCallData['Module'] = PD_RO['module']
FirstCallData['Price'] = PD_RO['price']
FirstCallData['C_desig'] = np.nan
FirstCallData['no_attempts'] = 0
FirstCallData['nPitch_code'] = [[] for i in FirstCallData.index]
FirstCallData['nInterest_code'] = [[] for i in FirstCallData.index]
FirstCallData['nInterest_str'] = str(np.nan)
FirstCallData['Pitch_Dummy'] = np.nan
FirstCallData['Discuss_p'] = [[] for i in FirstCallData.index]
FirstCallData['Qs_concerns'] = [[] for i in FirstCallData.index]
FirstCallData['Info_dummy'] = np.nan
FirstCallData['nInfo_code'] = [[] for i in FirstCallData.index]
FirstCallData['nInfo_str'] = str(np.nan)
FirstCallData['Comments'] = str(np.nan)
pd.DataFrame.save(FirstCallData, r"C:\Users\rcreedon\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\DataFrames\Wave1\FirstCallDataFrame")

