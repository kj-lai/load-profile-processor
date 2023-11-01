from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import PySimpleGUI as sg
import seaborn as sns
sg.theme('DarkAmber')
sns.set()

def data_processing(LOGGER_FILE,PVSYST_FILE,OUTPUT_DIR,COMPANY,CAPACITY,DESC):
	
	# OUTPUT FILE DIRECTORY
	if len(OUTPUT_DIR) == 0:
		OUTPUT_DIR  = str(Path.cwd())
	OUTPUT_FILE = OUTPUT_DIR+f'/energy_graph_load_profile'
	OUTPUT_FILE = OUTPUT_FILE+'_'+COMPANY if len(COMPANY)!=0 else OUTPUT_FILE
	CAPACITY    = f'{float(CAPACITY):.3f}kWp' if len(CAPACITY)!=0 else CAPACITY
	OUTPUT_FILE = OUTPUT_FILE+'_'+CAPACITY if len(CAPACITY)!=0 else OUTPUT_FILE
	OUTPUT_FILE = OUTPUT_FILE+'_'+DESC if len(DESC)!=0 else OUTPUT_FILE
	FILE_END    = '.xlsx'
	OUTPUT_FILE = OUTPUT_FILE+FILE_END

	# PROCESS PVSYST DATA
	if len(PVSYST_FILE) != 0:
		COLUMN_NAME = ['date','EOutInv']

		data = pd.read_csv(PVSYST_FILE, skiprows=13, names=COLUMN_NAME, delimiter=';')
		data['datetime'] = pd.to_datetime(data.date, format='%d/%m/%y %H:%M')
		data1 = data[['datetime','EOutInv']].copy()
		data1['hour'] = data1.datetime.apply(lambda x: x.hour)
		data1 = data1.groupby(['hour']).agg({'EOutInv':'mean'}).reset_index().copy()

	# PROCESS ENERGY LOGGER DATA
	DATE_COL   = ['Start(Malay Peninsula Standard Time)', 'Stop(Malay Peninsula Standard Time)']
	VOL_COL    = ['Vrms_AN_max', 'Vrms_BN_max', 'Vrms_CN_max']
	AMP_COL    = ['Irms_A_max', 'Irms_B_max', 'Irms_C_max']
	PWR_COL    = ['PowerP_A_max', 'PowerP_B_max', 'PowerP_C_max', 'PowerP_Total_max']
	LOGGER_COL = DATE_COL + VOL_COL + AMP_COL + PWR_COL

	logger = pd.read_csv(LOGGER_FILE, delimiter=';')
	logger1 = logger[LOGGER_COL].copy()
	logger1.iloc[:, [0]] =  pd.to_datetime(logger1.iloc[:, 0], format='%Y-%m-%d %H:%M:%S.%f').apply(lambda x: x.replace(microsecond=0))
	logger1.iloc[:, [1]] =  pd.to_datetime(logger1.iloc[:, 1], format='%Y-%m-%d %H:%M:%S.%f').apply(lambda x: x.replace(microsecond=0))
	for col in PWR_COL:
		logger1[col] =  logger1[col]/1000
	logger1['date'] = logger1.iloc[:, 0].apply(lambda x: x.date())
	logger1['weekday'] = logger1.iloc[:, 0].apply(lambda x: x.day_name())
	logger1['hour'] = logger1.iloc[:, 0].apply(lambda x: x.hour)

	agg_dict = {}
	for col in VOL_COL+AMP_COL+PWR_COL:
		agg_dict.update({col:'mean'})
	logger2 = logger1.groupby(['date', 'weekday', 'hour'])\
					 .agg(agg_dict)\
					 .reset_index()\
					 .copy()

	# MERGE BOTH DATA
	if len(PVSYST_FILE) != 0:

		final = pd.merge(logger2, data1, on=['hour'])\
				  .sort_values(['date','hour'])\
				  .reset_index(drop=True)
	else:
		final = logger2.copy()
	final.to_excel(OUTPUT_FILE, index=False)

	# DAYTIME PEAK & LOWEST DEMAND
	peak = final[final.hour.isin(range(11,16))]
	max_row = peak.PowerP_Total_max.idxmax()
	peak = np.round(peak.loc[[max_row],'PowerP_Total_max'].values[0], 2)
	low = final[final.hour.isin(range(7,22))]
	min_row = low.PowerP_Total_max.idxmin()
	low = np.round(low.loc[[min_row],'PowerP_Total_max'].values[0], 2)

	print(f'Final Table:',final.shape)
	print(f'Daytime Peak Demand: {peak} kW')
	print(f'Daytime Lowest Demand: {low} kW')

	# PLOT GRAPH
	nDays = final.date.nunique()
	Date = str(final.date.values[0])
	nInterval = 4 # 4 hours intervals
	final['datetime'] = final.apply(lambda x: str(x.date.day).zfill(2)+'/'+str(x.date.month).zfill(2)
									+' '+str(x.hour).zfill(2)+':00', axis=1)
	
	TITLE    = ''
	TITLE    = ' for \n'+COMPANY if len(COMPANY)!=0 else TITLE
	TITLE    = TITLE+'_'+CAPACITY if len(CAPACITY)!=0 else TITLE

	# FIGURE 1
	fig = plt.figure(figsize=(14,12)) # set (width, height)
	fig.patch.set_facecolor('white')
	plt.plot(final.index, final[PWR_COL[-1]], label='Load Profile')
	plt.title('Load Profile for'+TITLE+# \n'+COMPANY+'_'+CAPACITY+
			  '\nfrom '+str(final.date.values[0])+' to '+str(final.date.values[-1]), fontsize=16, linespacing=2)
	plt.xlabel('Datetime', fontsize=16)
	last_label = pd.Series(Date[-2:]+'/'+Date[-5:-3]+' 00:00')
	plt.xticks(range(0,len(final),nInterval), final.datetime[::nInterval], fontsize=12, rotation=75)
	plt.ylabel('Power (kW)', fontsize=16)
	plt.yticks(fontsize=14)
	plt.legend(fontsize=12)
	plt.savefig(OUTPUT_DIR+'/load_profile_graph.png')

	# FIGURE 2
	if len(PVSYST_FILE) != 0:
		fig = plt.figure(figsize=(14,12)) # set (width, height)
		fig.patch.set_facecolor('white')
		plt.plot(final.index, final[PWR_COL[-1]], label='Load Profile')
		plt.plot(final.index, final.EOutInv, label='Inverter Generation')
		plt.title('Comparison of Load Profile against Inverter Generation'+TITLE+# for \n'+COMPANY+'_'+CAPACITY+
				  '\nfrom '+str(final.date.values[0])+' to '+str(final.date.values[-1]), fontsize=16, linespacing=2)
		plt.xlabel('Datetime', fontsize=16)
		last_label = pd.Series(Date[-2:]+'/'+Date[-5:-3]+' 00:00')
		plt.xticks(range(0,len(final),nInterval), final.datetime[::nInterval], fontsize=12, rotation=75)
		plt.ylabel('Power (kW)', fontsize=16)
		plt.yticks(fontsize=14)
		plt.legend(fontsize=12)
		plt.savefig(OUTPUT_DIR+'/load_profile_vs_inverter_generation_graph.png')
	plt.show()

# ------------------------------------------------------------------------------------------------------------------------

# FOR TESTING PURPOSES
# while True:
# 	# PVsyst filename
# 	PVSYST_FILE = 'sample_pvsyst_file.CSV'
# 	# Energy logger filename
# 	LOGGER_FILE = 'sample_energy_logger_file.txt'
# 	COMPANY     = 'SAMPLE'
# 	CAPACITY    = 120.95
# 	DESC        = ''
# 	OUTPUT_DIR  = 'D:/Workspace/01 - Eco Solar/Project/sample'

# 	data_processing(PVSYST_FILE,LOGGER_FILE,OUTPUT_DIR,COMPANY,CAPACITY,DESC)
# 	break;

# ------------------------------------------------------------------------------------------------------------------------

# CHECK FOR VALID FILEPATH
def is_valid_filepath(FILEPATH):
	if FILEPATH and Path(FILEPATH).exists():
		return True
	sg.popup_error('Invalid File Path!!')
	return False

# GUI DEFINITION
layout = [
	[sg.Text('Logger File:'), sg.Input(key='-In1-'), sg.FileBrowse(file_types=(('Text Files', '*.txt*'),))],
	[sg.Text('PVsyst File:'), sg.Input(key='-In2-'), sg.FileBrowse(file_types=(('Text Files', '*.CSV*'),))],
	[sg.Text('Output Directory:'), sg.InputText(key='-Out-'), sg.FolderBrowse()],
	[sg.Text('Company Name:'), sg.InputText(key='-Name-')],
	[sg.Text('PV Capacity kWp:'), sg.InputText(key='-Cap-')],
	[sg.Text('Description:'), sg.InputText(key='-Desc-')],
	[sg.Button('Submit'), sg.Exit()]
]

window = sg.Window('Load Profile Processor', layout)

while True:
	event, values = window.read()
	print(event, values)
	if event in (sg.WINDOW_CLOSED, 'Exit'):
		break

	if(event == 'Submit'):
		if is_valid_filepath(values['-In1-']):
			try:
				data_processing(LOGGER_FILE=values['-In1-'],
								PVSYST_FILE=values['-In2-'],
								OUTPUT_DIR=values['-Out-'],
								COMPANY=values['-Name-'],
								CAPACITY=values['-Cap-'],
								DESC=values['-Desc-'],
								)
			except:
				sg.popup_error('Processing Error!!')
window.close()