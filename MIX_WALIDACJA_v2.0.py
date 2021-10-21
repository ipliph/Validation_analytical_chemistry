import numpy as np 
import pandas as pd

###Data extraction part

df = pd.ExcelFile('06-21_MIX_walidacja_Short.xls')
sheet_names = df.sheet_names

#sheet dictionary
sheet_dict = dict(zip(range(1,9),sheet_names))

###{1: 'IS', 2: 'C3399met_221.1', 3: 'C3390met_170', 4: 'C3377_170.2', 5: 'C3390_199.2', 6: 'C3399_221.2', 7: 'C3389_91.2', 8: 'C3380_161.3'}

sheet_to_process = sheet_dict[4]
df = pd.read_excel('06-21_MIX_walidacja_Short.xls', sheet_name=sheet_to_process, index_col=0)


df.columns = df.iloc[3]

#Unnamed: 0 - Filename ()
#Unnamed: 4 - Area (3)
#Unnamed: 5 - ISTD Area (4)
#Unnamed: 6 - Area Ratio (5)
#Unnamed: 14 - RT

###Dataframe.loc[["row1", "row2"], ["column1", "column2", "column3"]]
df_short = df.iloc[0:,[3,4,5,13]]

assert df.iloc[3,3] == 'Area'
assert df.iloc[3,4] == 'ISTD Area'
assert df.iloc[3,5] == 'Area Ratio'
assert df.iloc[3,13] == 'RT'

###DataFrame: calibration curve
item_list = [num for num in range(1,20)]
df_cc = pd.DataFrame()


num = '1'
for item in item_list:
	if 'MIX_{}CC1'.format(num) in df_short.index:
		df_cc = df_cc.append(df_short.loc[['{}blank'.format(num)]], ignore_index=False)
		df_cc = df_cc.append(df_short.loc[['{}zero'.format(num)]], ignore_index=False)
		item = df_short.loc[['MIX_{}CC{}'.format(item,x) for x in range(1,14)]]
		df_cc = df_cc.append(item, ignore_index=False)
		num = int(num)
		num = num +1
		num = str(num)

###DataFrame: Area ratios
concentrations = {
	'Sample lables':['blank','zero','CC1','CC2','CC3','CC4','CC5','CC6','CC7','CC8','CC9','CC10','CC11','CC12','CC13'],
	'Specified Amount':['ns','ns',0.01,0.025,0.05,0.1,0.25,0.5,1,2.5,5,10,20,35,50]}

df_concentrations = pd.DataFrame(concentrations)
#Array only with Area Ratios values (1D)
ratio_array = np.array(df_cc['Area Ratio'])
cc_index = int(len(ratio_array)/15)
#Array only with Area Ratios values (2D)
ratio_array = np.array(df_cc['Area Ratio']).reshape(cc_index,15)
ratio_array = np.rot90(ratio_array,k=1,axes=(1, 0))
ratio_array = np.fliplr(ratio_array)
df_ratios = pd.DataFrame(ratio_array,columns=['{}CC'.format(i) for i in range(1,cc_index+1)])

df_cc_ratios = pd.concat([df_concentrations,df_ratios],axis=1)

###DataFrame: quality concentration
df_qc = pd.DataFrame()

num = '1'
for item in item_list:
	if 'MIX_{}QC1'.format(num) in df_short.index:
		item = df_short.loc[['MIX_{}QC{}'.format(item,x) for x in range(1,5)]]
		df_qc = df_qc.append(item, ignore_index=False)
		num = int(num)
		num = num +1
		num = str(num)

###DataFrame: post extractum samples
df_pe = pd.DataFrame()

num = '1'
for item in item_list:
	if 'MIX_{}PE1'.format(num) in df_short.index:
		item = df_short.loc[['MIX_{}PE{}'.format(item,x) for x in range(1,5)]]
		df_pe = df_pe.append(item, ignore_index=False)
		num = int(num)
		num = num +1
		num = str(num)

###DataFrame: solutions (roztworowe)
df_r = pd.DataFrame()

num = '1'
for item in item_list:
	if 'MIX_{}R1'.format(num) in df_short.index:
		item = df_short.loc[['MIX_{}R{}'.format(item,x) for x in range(1,5)]]
		df_r = df_r.append(item, ignore_index=False)
		num = int(num)
		num = num +1
		num = str(num)

###DataFrame: Stability

df_stab = pd.DataFrame()

for item in range(1,4):
	item = df_short.loc[['MIX_{}SK{}'.format(item,x) for x in range(1,5)]]
	df_stab = df_stab.append(item, ignore_index=False)

for item in range(4,7):
	item = df_short.loc[['MIX_{}QC{}_A24'.format(item,x) for x in range(1,5)]]
	df_stab = df_stab.append(item, ignore_index=False)

for item in range(1,4):
	item = df_short.loc[['MIX_{}ZR{}'.format(item,x) for x in range(1,5)]]
	df_stab = df_stab.append(item, ignore_index=False)

### Computational Part

### Recovery
qc_av_dict = {}
	
for qc_numb in range(1,5):

	day1_QC_av = np.mean([(df_qc['Area Ratio'].loc['MIX_{}QC{}'.format(num,qc_numb)]) for num in range(1,4)])
	qc_av_dict['day{}_QC{}_av'.format(1,qc_numb)] = day1_QC_av
	
	day2_QC_av = np.mean([(df_qc['Area Ratio'].loc['MIX_{}QC{}'.format(num,qc_numb)]) for num in range(4,7)])
	qc_av_dict['day{}_QC{}_av'.format(2,qc_numb)] = day2_QC_av

	day3_QC_av = np.mean([(df_qc['Area Ratio'].loc['MIX_{}QC{}'.format(num,qc_numb)]) for num in range(7,10)])
	qc_av_dict['day{}_QC{}_av'.format(3,qc_numb)] = day3_QC_av

	day4_QC_av = np.mean([(df_qc['Area Ratio'].loc['MIX_{}QC{}'.format(num,qc_numb)]) for num in range(10,13)])
	qc_av_dict['day{}_QC{}_av'.format(4,qc_numb)] = day4_QC_av

pe_av_dict = {}
	
for pe_numb in range(1,5):

	day1_PE_av = np.mean([(df_pe['Area Ratio'].loc['MIX_{}PE{}'.format(num,pe_numb)]) for num in range(1,4)])
	pe_av_dict['day{}_PE{}_av'.format(1,pe_numb)] = day1_PE_av
	
	day2_PE_av = np.mean([(df_pe['Area Ratio'].loc['MIX_{}PE{}'.format(num,pe_numb)]) for num in range(4,7)])
	pe_av_dict['day{}_PE{}_av'.format(2,pe_numb)] = day2_PE_av

	day3_PE_av = np.mean([(df_pe['Area Ratio'].loc['MIX_{}PE{}'.format(num,pe_numb)]) for num in range(7,10)])
	pe_av_dict['day{}_PE{}_av'.format(3,pe_numb)] = day3_PE_av

	day4_PE_av = np.mean([(df_pe['Area Ratio'].loc['MIX_{}PE{}'.format(num,pe_numb)]) for num in range(10,13)])
	pe_av_dict['day{}_PE{}_av'.format(4,pe_numb)] = day4_PE_av

index_list = ['day1_lvl1','day2_lvl1','day3_lvl1','day4_lvl1',
				'day1_lvl2','day2_lvl2','day3_lvl2','day4_lvl2',
				'day1_lvl3','day2_lvl3','day3_lvl3','day4_lvl3',
				'day1_lvl4','day2_lvl4','day3_lvl4','day4_lvl4']

df_recovery = pd.DataFrame(index=index_list)

df_recovery['QC_av']=qc_av_dict.values()
df_recovery['PE_av']=pe_av_dict.values()
df_recovery['Recovery[%]']= (df_recovery['QC_av']/df_recovery['PE_av'])*100

###Normalised matrix effect


pe_av_area_dict = {}
for pe_numb in range(1,5):
	day1_PE_area_av = np.mean([(df_pe['Area'].loc['MIX_{}PE{}'.format(num,pe_numb)]) for num in range(1,4)])
	pe_av_area_dict['day{}_PE{}_av'.format(1,pe_numb)] = day1_PE_area_av

pe_av_ISarea_dict = {}
for pe_numb in range(1,5):
	day1_PE_ISarea_av = np.mean([(df_pe['ISTD Area'].loc['MIX_{}PE{}'.format(num,pe_numb)]) for num in range(1,4)])
	pe_av_ISarea_dict['day{}_PE{}_av'.format(1,pe_numb)] = day1_PE_ISarea_av

r_av_area_dict = {}
for r_numb in range(1,5):
	day1_R_area_av = np.mean([(df_r['Area'].loc['MIX_{}R{}'.format(num,r_numb)]) for num in range(1,4)])
	r_av_area_dict['day{}_R{}_av'.format(1,r_numb)] = day1_R_area_av

r_av_ISarea_dict = {}
for r_numb in range(1,5):
	day1_R_ISarea_av = np.mean([(df_r['ISTD Area'].loc['MIX_{}R{}'.format(num,r_numb)]) for num in range(1,4)])
	r_av_ISarea_dict['day{}_R{}_av'.format(1,r_numb)] = day1_R_ISarea_av

df_nme = pd.DataFrame(index=['Lvl1','Lvl2','Lvl3','Lvl4'])
df_nme['PE area'] = pe_av_area_dict.values()
df_nme['PE IS area'] = pe_av_ISarea_dict.values()
df_nme['R area'] = r_av_area_dict.values()
df_nme['R IS area'] = r_av_ISarea_dict.values()

df_nme['NME[%]'] = (df_nme['PE area']/df_nme['R area'])/(df_nme['PE IS area']/df_nme['R IS area'])*100

###Saving Results
with pd.ExcelWriter(str(sheet_to_process)+' output.xlsx') as writer:  
	df_cc.to_excel(writer,sheet_name=str(sheet_to_process)+' df_cc')
	df_cc_ratios.to_excel(writer,sheet_name=str(sheet_to_process)+' df_cc_ratios')
	df_qc.to_excel(writer,sheet_name=str(sheet_to_process)+' df_qc')
	df_pe.to_excel(writer,sheet_name=str(sheet_to_process)+' df_pe')
	df_r.to_excel(writer,sheet_name=str(sheet_to_process)+' df_r')
	df_stab.to_excel(writer,sheet_name=str(sheet_to_process)+' df_stab')
	df_recovery.to_excel(writer,sheet_name=str(sheet_to_process)+' df_recovery')
	df_nme.to_excel(writer,sheet_name=str(sheet_to_process)+' df_nme')
	

print('Finished')