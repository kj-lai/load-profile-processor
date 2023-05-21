from glob import glob
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import PySimpleGUI as sg
import seaborn as sns
sg.theme('DarkAmber')
sns.set()

def data_processing(PVSYST_FILE,LOGGER_FILE,OUTPUT_DIR,START_DATE,END_DATE,COMPANY,CAPACITY,DESC):
	# PVsyst filename
	# PVSYST_FILE = 'sample_pvsyst_file.CSV'
	# Energy logger filename
	# LOGGER_FILE = 'sample_energy_logger_file.txt'

	# START_DATE  = '2022-08-03'
	# END_DATE	= '2022-08-07' # Add 1 day extra

	# COMPANY 	= 'SAMPLE'
	# CAPACITY	= '%.3f' % CAPACITY
	# DESC		= ''
	OUTPUT_FILE = OUTPUT_DIR+'/energy_graph_load_profile_'+COMPANY+'_'+CAPACITY+'kwp'+DESC+'.xlsx'

	DAYS 		= ['Friday', 'Saturday', 'Sunday', 'Monday']
	sorterIndex = dict(zip(DAYS, range(len(DAYS))))

	# Process PVsyst data
	COLUMN_NAME = ['date','EOutInv']

	data = pd.read_csv(PVSYST_FILE, skiprows=13, names=COLUMN_NAME, delimiter=';')
	data['datetime'] = pd.to_datetime(data.date, format='%d/%m/%y %H:%M')
	data1 = data[['datetime','EOutInv']].copy()
	data1['hour'] = data1.datetime.apply(lambda x: x.hour)
	data1['weekday'] = data1.datetime.apply(lambda x: x.day_name())
	data1 = data1[data1.weekday.isin(DAYS)]

	data2 = data1.groupby(['weekday','hour']).agg({'EOutInv':'mean'}).reset_index().copy()
	data2['rank'] = data2['weekday'].map(sorterIndex)
	data2.sort_values(['rank','hour'], inplace=True)
	data2.reset_index(drop=True, inplace=True)
	data2.drop('rank', axis=1, inplace=True)

	# Process Energy Logger data
	DATE_COL   = ['Start(Malay Peninsula Standard Time)', 'Stop(Malay Peninsula Standard Time)']
	VOL_COL    = ['Vrms_AN_max', 'Vrms_BN_max', 'Vrms_CN_max']
	AMP_COL    = ['Irms_A_max', 'Irms_B_max', 'Irms_C_max']
	PWR_COL    = ['PowerP_A_max', 'PowerP_B_max', 'PowerP_C_max', 'PowerP_Total_max']
	LOGGER_COL = DATE_COL + VOL_COL + AMP_COL + PWR_COL

	logger = pd.read_csv(LOGGER_FILE, delimiter=';')
	logger1 = logger[LOGGER_COL].copy()
	logger1['Start(Malay Peninsula Standard Time)'] = pd.to_datetime(logger1.iloc[:, 0], format='%Y-%m-%d %H:%M:%S.%f').apply(lambda x: x.replace(microsecond=0))
	for col in PWR_COL:
	    logger1[col] =  logger1[col]/1000
	logger1['date'] = logger1.iloc[:, 0].apply(lambda x: x.date())
	logger1 = logger1[(logger1.iloc[:, 0] >= START_DATE) & (logger1.iloc[:, 0] < END_DATE)]
	logger1['weekday'] = logger1.iloc[:, 0].apply(lambda x: x.day_name())
	logger1['hour'] = logger1.iloc[:, 0].apply(lambda x: x.hour)

	agg_dict = {}
	for col in VOL_COL+AMP_COL+PWR_COL:
		agg_dict.update({col:'mean'})
	logger2 = logger1.groupby(['date', 'weekday', 'hour'])\
					 .agg(agg_dict)\
					 .reset_index()\
					 .copy()

	assert len(data2) == len(logger2), 'Length not match with PVsyst '+str(len(data2))+' and Logger '+str(len(logger2))
	final = pd.merge(logger2, data2, on=['weekday', 'hour'])
	final.to_excel(OUTPUT_FILE, index=False)

	# Daytime Peak & Lowest Demand
	peak = final[final.hour.isin(range(11,16))]
	max_row = peak.PowerP_Total_max.idxmax()
	peak = np.round(peak.loc[[max_row],'PowerP_Total_max'].values[0], 2)
	low = final[final.hour.isin(range(7,22))]
	min_row = low.PowerP_Total_max.idxmin()
	low = np.round(low.loc[[min_row],'PowerP_Total_max'].values[0], 2)


	print(f'Final Table:',final.shape)
	print(f'Daytime Peak Demand: {peak} kW')
	print(f'Daytime Lowest Demand: {low} kW')

#check for valid filepath
def is_valid_filepath(FILEPATH):
	if FILEPATH and Path(FILEPATH).exists():
		return True
	sg.popup_error('Invalid File Path!!')
	return False

# GUI Definition
layout = [
	[sg.Text('PVsyst File:'), sg.Input(key='-In1-'), sg.FileBrowse(file_types=(('Text Files', '*.CSV*'),))],
	[sg.Text('Logger File:'), sg.Input(key='-In2-'), sg.FileBrowse(file_types=(('Text Files', '*.txt*'),))],
	[sg.Text('Output Folder:'), sg.Input(key='-Out-'), sg.FolderBrowse()],
	[sg.Text('Start Date:'), sg.InputText(key='-Start-')],
	[sg.Text('End Date:'), sg.InputText(key='-End-')],
	[sg.Text('Company Name:'), sg.InputText(key='-Name-')],
	[sg.Text('PV Capacity:'), sg.InputText(key='-Cap-')],
	[sg.Text('Description:'), sg.InputText(key='-Desc-')],
	[sg.Button('Submit'), sg.Exit()]
]

window = sg.Window('Load Profile Data Processor', layout)

while True:
	event, values = window.read()
	print(event, values)
	if event in (sg.WINDOW_CLOSED, 'Exit'):
		break

	if(event == 'Submit'):
		if is_valid_filepath(values['-In1-']) and is_valid_filepath(values['-In2-']) and is_valid_filepath(values['-Out-']):
			try:
				data_processing(PVSYST_FILE=values['-In1-'],
								LOGGER_FILE=values['-In2-'],
								OUTPUT_DIR=values['-Out-'],
								START_DATE=values['-Start-'],
								END_DATE=values['-End-'],
								COMPANY=values['-Name-'],
								CAPACITY=values['-Cap-'],
								DESC=values['-Desc-'],
								)
			except:
				sg.popup_error('Processing Error!!')
window.close()