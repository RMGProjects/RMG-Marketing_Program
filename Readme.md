|     |     |
| --- | --- |
| Application: | **RMG-Marketing_Program** |
| Author:      | **Rory Creedon** (rcreedon@poverty-action.org) |
| Use:         | **Primarily intended for deployment through DataNitro Excel Add in** |
| Date:		   | **August 2013** |
| Status:	   | **The project is now complete. The code and documents are shared for information purposes only** |

The files displayed here were created in order to facilitate the collection of data that related to a marketing exercise undertaken by IPA Bangladesh in the Ready Made Garments ('RMG') industry industry . A brief overview of the structure of the project is given below.

**NB** Please note that the code here is no longer active. It was one of the first system of programs that I wrote and is therefore likely to be full of errors and poor coding practice. I apologise for that. Were I to do this again I would do it differently, but nevertheless the data collected were high quality and as such I believe there is value in sharing the code along with the associated files. I cannot explain the meaning of every component and variable beyond giving an overview of the structure as a whole. If you are interested (particularly IPA/J-PAL staff) to know more please get in touch at the email address provided above.

**Contents of Repository:**
	
|     |     | --- |
| --- | --- | --- |
| **Folder**   		| **File Name** 					| **Description**|
| <>		   		| 01_first_call_dataset.py			| Creates a data set that will be populated by the first_call_program.py |
| <>           		| 02_follow_up_1_dataset.py			| Creates a data set that will be populated by the follow_up1_call_program.py |
| <>		   		| 03_call_status.py					| Program to be run in DataNitro inside Call_Status.xlsm; checks on the call status (i.e. progress made) |
| <>		   		| 04_postal_status.py				| Program to be run in DataNitro inside Postal_Status.xlsm; checks on the postal status of documents to be sent |
| <>           		| 05_first_call_program.py			| Program to be run in DataNitro inside First_call.xlsm; updates contact info, receives and checks data, stores data in data frame and saves raw data (in excel format) as backup|
| <>		   		| 06_follow_up1_call_program.py		| Program to be run in DataNitro inside Follow_up1_call.xlsm; updates contact info, receives and checks data, stores data in data frame and saves raw data (in excel format) as backup |
| DataCollection	| first_call.xlsm					| Data collection tool for first call |
| DataCollection	| follow_up1_call.xlsm				| Data collection tool for first follow up call |
| OtherTools		| call_status.xlsm					| Displays status of calls made |
| OtherTools		| postal_status.xlsm				| Displays status of documents posted |
| BackgroundData	| example_factory_list.csv			| An example of the factory list that feeds into the .py files |
| BackgroundData	| example_contact_list.csv			| An example of a factory contact list that feeds into the .py files |
| SupportingDocs	| Code_Book.docx					| The codebook used when entering data |
| SupportingDocs	| Data_structures_outcomes.docx		| Explanation of general data structures and outcomes at each stage of the call process |
| SupportingDocs	| FAQs.docx							| Frequently asked questions tool for use by marketing team |	
| SupportingDocs	| First_call_Data_structure.docx	| Explanation of data structure	of first call |
| SupportingDocs	| Follow_Up1_data_structure.docx	| Explanation of data structure of first follow up call |
| SupportingDocs	| Processes_procedures.docx			| Not entirely up to date explanation of the project processes and procedures |

**NB** files in the BackgroundData folder have had key information removed to preserve anonymity of participating factories. Files presented relate only to the first group of factories that participated (96 in total).
	
**Method:** - The data collection method was to enter data in Excel using templates created by project management, and then to use DataNitro to execute python scripts in excel in order to test the quality of the data (and return errors to the user) and then store that data. Several additional tools were created apart from the main data collection tools in order that the user might keep track of progress made.

The key back-end data was contained in the contact list. The contact list was not available to the user, but was modifiable by the user operating through the data collection tools. The contact list contained the addresses, contact numbers and variables that determined call status, postal status, and other outcomes. 

**DataNitro:** - DataNitro is an excel add in that executes python scripts in excel. It has its own functions and methods. See the documentation for more information (https://datanitro.com/docs/) . 

**Project Background:** - The project that these tools relate to aimed to find the price elasticity of demand for a training program provided by a Dhaka based training centre to supervisors from within RMG factories. The training was comprised of three modules:
	
	1. P: Production and Work Study
	2. Q: Quality Control Management 
	3. H: Leadership and Social Compliance
	
A list of factories and contact details were obtained (total 600 factories), and these factories were randomly sorted into four groups of 96 (the remainder were either part of the pilot or were not targeted).

Each factory was randomly assigned to one of the above modules. This module would be offered for free to the factory as a means of learning about the training. Additionally the factory was randomly assigned to a price. If the factory was assigned to price 'HIGH' then they would be able to purchase additional training for their supervisor staff at a rate of 14,000 taka. If they were assigned to price 'LOW' then they could purchase extra training at the cost of 9,250. The value of the pirces HIGH and LOW was varied between the four groups in order to establish the demand at different price points. 

The call process was as follows:

	1. First Call - Calls were made until a decision maker was reached. Once reached the marketing team would pitch the training and mention both the price and the free module. If the decision maker agreed to receive a pack of marketing materials by post, the first call process was deemed to be at an end and that factory would progress to the first follow up call
	2. Follow Up 1 - Calls were made until the decision maker decided whether or not the factory would send someone to the free module. If they agreed they would progress to the Second Follow Up Call
	3. Follow Up 2 - This call was made post free module to determine if the factory would purchase more training for their staff. 

**NB** The tools and documents presented here relate only to 1 & 2, 3 has been omitted as the author only has so so much time and energy. The ransomization was done in stata and the .do-files are not presented here. 

The work-flow process during one full day was as follows:
	
	1. Check Call Status - to determine which factories need to be called, called again, followed up with etc. 
	2. Make Calls - select correct data collection tool depending on the nature of the call. 
	3. Check Postal Status - to determine to which factories the marketing materials needed to be sent. 
	4. End
