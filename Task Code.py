import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings("ignore")

### --- Importing data ---

info_xlsx = pd.ExcelFile('Humanyze_Internal_Dataset_email.xlsx')
info = pd.read_excel(info_xlsx, sheetname = 'ParticipantInfo')

dm_xlsx = pd.ExcelFile('Humanyze_Internal_Dataset_email.xlsx')
dm = pd.read_excel(dm_xlsx, sheetname = 'Daily E-mail')


### --- Merging data to create data mart ---

info.rename(columns = {'UserID':'Sender'}, inplace = True)
merged1 = pd.merge(dm,info, on='Sender', how='left')
merged1.rename(columns = {'Team':'Send_team','Location':'Send_loc','Manager?':'Send_mngr'}, inplace = True)

info.rename(columns = {'Sender':'Receiver'}, inplace = True)
merged2 = pd.merge(merged1,info, on='Receiver', how='left')
merged2.rename(columns = {'Team':'Rec_team','Location':'Rec_loc','Manager?':'Rec_mngr'}, inplace = True)


### --- Split dataframe and define new columns for task 1 ---

merged2['SendRec_loc'] = np.where(merged2['Send_loc']==merged2['Rec_loc'], 'Same Location', 'Different Location')

sl = merged2.query('SendRec_loc == "Same Location"')
dl = merged2.query('SendRec_loc == "Different Location"')

sl_bo = sl.query('Send_loc == "Boston"')
sl_pa = sl.query('Send_loc == "Palo Alto"')

sl_bo['SendRec_team'] = np.where(sl_bo['Send_team']==sl_bo['Rec_team'], 'Same Team', 'Different Team')
sl_pa['SendRec_team'] = np.where(sl_pa['Send_team']==sl_pa['Rec_team'], 'Same Team', 'Different Team')

sl_bo_st = sl_bo.query('SendRec_team == "Same Team"')
sl_pa_st = sl_pa.query('SendRec_team == "Same Team"')

print("--- Cohesion ---")
print("Teammates in Boston email each other on average : %.2f"
       % np.round(sl_bo_st['Count'].aggregate(np.mean),2)
      + " times")
print("Teammates in Palo Alto email each other on average : %.2f"
      % np.round(sl_pa_st['Count'].aggregate(np.mean),2)
      + " times")

sl_bo_st_mngr = sl_bo_st.query('Rec_mngr == "Y"')
sl_pa_st_mngr = sl_pa_st.query('Rec_mngr == "Y"')

print("\n--- Manager Visibility ---")
print("Teammates in Boston email their manager on average : %.2f"
      % np.round(sl_bo_st_mngr['Count'].aggregate(np.mean),2)
      + " times")
print("Teammates in Palo Alto email their manager on average : %.2f"
      % np.round(sl_pa_st_mngr['Count'].aggregate(np.mean),2)
      + " times")

byloc = merged2.groupby(['Send_loc'])
print("\n--- Overall Engagement ---")
print("Teammates in both the locations send emails on average as follows:")
print(round(byloc['Count'].aggregate(np.mean),2).to_string(header=None))


### --- task 4 ---

## 4a

byloc_flag = merged2.groupby(['Send_loc','SendRec_loc'])
print("\n--- Communication across sites ---")
print(round(byloc_flag['Count'].aggregate(np.mean),2).to_string(header=None))

temp = round(byloc_flag['Count'].aggregate(np.mean),2).to_frame()
temp.plot(kind = 'barh', figsize = (10,4), color = 'b', label = "Avg. # of emails", align = "center")
plt.xlabel('')
plt.ylabel('')
plt.title('Communication across sites')
plt.legend()
plt.subplots_adjust(left=0.22, bottom=0.20, right=0.94, top=0.90, wspace=0.2, hspace=0)

## 4b and 4c 

byteam = merged2.groupby(['Send_team','Rec_team'])
print("\n--- Communication between teams ---")
print(round(byteam['Count'].aggregate(np.mean),2).to_string(header=None))

temp = round(byteam['Count'].aggregate(np.mean),2).to_frame()
temp.plot(kind = 'barh', figsize = (10,4), color = 'b', label = "Avg. # of emails", align = "center")
plt.xlabel('')
plt.ylabel('')
plt.title('Communication between teams')
plt.legend()
plt.subplots_adjust(left=0.22, bottom=0.09, right=0.94, top=0.90, wspace=0.2, hspace=0)
plt.show()
