import pandas as pd
import numpy as np

PD_RO = pd.read_csv(r"C:\Users\rcreedon\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\FactoryList\Wave1\FactoryList.csv", index_col = 'UID')
UIDList = [i for i in PD_RO.index]
FollowUp1 = pd.DataFrame(index = UIDList)
FollowUp1['no_attempts'] = 0
FollowUp1['nDiscuss_code'] = [[] for i in FollowUp1.index]
FollowUp1['nInterest_code'] = [[] for i in FollowUp1.index]
FollowUp1['nInterest_str'] = str(np.nan)
FollowUp1['Discuss_Dummy'] = 0
FollowUp1['Discuss_p'] = [[] for i in FollowUp1.index]
FollowUp1['Qs_concerns'] = [[] for i in FollowUp1.index]
FollowUp1['free_dummy'] = np.nan
FollowUp1['nFree_code'] = [[] for i in FollowUp1.index]
FollowUp1['nFree_str'] = str(np.nan)
FollowUp1['Comments'] = str(np.nan)
pd.DataFrame.save(FollowUp1, r"C:\Users\rcreedon\Dropbox\MarketingTeamFolder\DataEntry\ReadOnlyFiles\DataFrames\Wave1\FollowUp1DataFrame")
