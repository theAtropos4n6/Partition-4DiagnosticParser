# DKats 2020 - Based on the research of Alexandros Vasilaras, Evangelos Dragonas, Dimitrios Katsoulis
# This script will take a device S/N and a Microsoft-Windows-Partition%4Diagnostic.evtx file to parse
# and will return all the volumes' S/Ns that ever existed on the media historically throughout the whole log.
# More info on the documentation at https://github.com/theAtropos4n6/Partition%4DiagnosticParser
#---command cheatsheet---
# admin cmd command to show your disk S/N: wmic diskdrive get model,name,serialnumber
# cmd command to show S/N of a specific volume: vol C:
# admin powershell command to show your disk S/N: Get-PhysicalDisk | Select-Object FriendlyName,SerialNumber
#---Thanks to---
# evtx module from  Omer Ben-Amram at https://pypi.org/project/evtx/
# PySimpleGUI module from  MikeTheWatchGuy at https://pypi.org/project/PySimpleGUI/
# executable's icon downloaded from www.freeiconspng.com
import PySimpleGUI as sg
from evtx import PyEvtxParser
import json
import webbrowser


#---functions definition
def volumeSNParser(rec, vbrNo, sig): #stelnw to signature tou ka8e volume, olo to record(json object) kai ton ari8mo tou vbr pou parsarw gia na parv to S/N tou volume	
	if sig == '4D53444F5335': # FAT32 volume
		s = rec['Event']['EventData'][vbrNo][134:142]
		sn = "".join(reversed([s[i:i+2] for i in range(0, len(s), 2)]))
		return sn
	elif sig == '455846415420': #ExFAT volume
		s = rec['Event']['EventData'][vbrNo][200:208]
		sn = "".join(reversed([s[i:i+2] for i in range(0, len(s), 2)]))
		return sn
	elif sig == '4E5446532020': # NTFS volume
		s = rec['Event']['EventData'][vbrNo][144:152]
		sn = "".join(reversed([s[i:i+2] for i in range(0, len(s), 2)])) #gia na gyrisw to SN anapoda me little endian byte order
		return sn
	else:
		s = 'Unknown Volume Type'
		return s

def fullparse():
	filename = values['-IN-']
	FullParseRecordsDict = {}
	AllPluggedInSerials = []
	EachPluggedDeviceDict = {}
	IsPartitionDiagnosticEVTXFullParse = True
	FullParseHTMLWritten = True

	try: #checkarw an einai legit evtx log
		parser = PyEvtxParser(filename)		
		for record in parser.records_json():
			data = json.loads(record['data'])	
			FullParseRecordsDict[(data['Event']['System']['EventRecordID'])] = data #add gia na sortaro meta me event id kai na exo xronologika sosti seira			
			if data['Event']['System']['EventID'] != 1006:
				IsPartitionDiagnosticEVTXFullParse = False
		
		FullParseLogStartTime = FullParseRecordsDict[1]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
		FullParseLogEndTime = FullParseRecordsDict[len(FullParseRecordsDict)]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
		if IsPartitionDiagnosticEVTXFullParse:			
			for i in sorted(FullParseRecordsDict.keys()):
				if FullParseRecordsDict[i]['Event']['EventData']['SerialNumber'] not in AllPluggedInSerials: # prwth fora pou vrskv to usb sto log opote to grafw sigoura
					AllPluggedInSerials.append(FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']) # pros8etw sth lista to S/N tou usb
					if FullParseRecordsDict[i]['Event']['EventData']['PartitionStyle'] == 1: #an einai GPT schemed media
						EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']] = [FullParseRecordsDict[i]['Event']['EventData']['Manufacturer'], FullParseRecordsDict[i]['Event']['EventData']['Model'], FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'), FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'), ['GPT - NO VSN INFO']] # EachPluggedDeviceDict = {'serial':[manufacturer, model, first pluggedIn, lastpluggedIn, [VSN1, VSN2 klp]]
					else:
						SN0 = volumeSNParser(FullParseRecordsDict[i], 'Vbr0', FullParseRecordsDict[i]['Event']['EventData']['Vbr0'][6:18])
						if FullParseRecordsDict[i]['Event']['EventData']['Vbr1'] != '':
							SN1 = volumeSNParser(FullParseRecordsDict[i], 'Vbr1', FullParseRecordsDict[i]['Event']['EventData']['Vbr1'][6:18])
						else:
							SN1 = '-'
						if FullParseRecordsDict[i]['Event']['EventData']['Vbr2'] != '':
							SN2 = volumeSNParser(FullParseRecordsDict[i], 'Vbr2', FullParseRecordsDict[i]['Event']['EventData']['Vbr2'][6:18])
						else:
							SN2 = '-'
						if FullParseRecordsDict[i]['Event']['EventData']['Vbr3'] != '':
							SN3 = volumeSNParser(FullParseRecordsDict[i], 'Vbr3', FullParseRecordsDict[i]['Event']['EventData']['Vbr3'][6:18])
						else:
							SN3 = '-'
						EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']] = [FullParseRecordsDict[i]['Event']['EventData']['Manufacturer'], FullParseRecordsDict[i]['Event']['EventData']['Model'], FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'), FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'), [SN0, SN1, SN2, SN3]]
				else: #to exw ksanavrei to S/N tou usb opote einai hdh sto dict kai apla 8elw na kanvw update to timestamp kai ta VSN
					if FullParseRecordsDict[i]['Event']['EventData']['PartitionStyle'] == 1: #an einai GPT schemed media opote sigoua den exei allaksei VSN afou den ta deixnei
						EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][3] = FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC') # kanw update mono to last pluggedin time
					else: #einai eite GPt pou egine unplugged h' MBR
						if FullParseRecordsDict[i]['Event']['EventData']['Vbr0'] != '': #VBR0 oxi keno opote exw vrei MBR plugged in record
							SN0 = volumeSNParser(FullParseRecordsDict[i], 'Vbr0', FullParseRecordsDict[i]['Event']['EventData']['Vbr0'][6:18])
							if FullParseRecordsDict[i]['Event']['EventData']['Vbr1'] != '':
								SN1 = volumeSNParser(FullParseRecordsDict[i], 'Vbr1', FullParseRecordsDict[i]['Event']['EventData']['Vbr1'][6:18])
							else:
								SN1 = '-'
							if FullParseRecordsDict[i]['Event']['EventData']['Vbr2'] != '':
								SN2 = volumeSNParser(FullParseRecordsDict[i], 'Vbr2', FullParseRecordsDict[i]['Event']['EventData']['Vbr2'][6:18])
							else:
								SN2 = '-'
							if FullParseRecordsDict[i]['Event']['EventData']['Vbr3'] != '':
								SN3 = volumeSNParser(FullParseRecordsDict[i], 'Vbr3', FullParseRecordsDict[i]['Event']['EventData']['Vbr3'][6:18])
							else:
								SN3 = '-'
							EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][3] = FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
							if SN3 not in EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]: #checkarw an einai hdh sth lista me ta VSN tou media alliws to pros8etw
								EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN3)
							if SN2 not in EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]: #checkarw an einai hdh sth lista me ta VSN tou media alliws to pros8etw
								EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN2)
							if SN1 not in EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]: #checkarw an einai hdh sth lista me ta VSN tou media alliws to pros8etw
								EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN1)
							if SN0 not in EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]: #checkarw an einai hdh sth lista me ta VSN tou media alliws to pros8etw
								EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN0)
						else:
							continue #next iteration giati exw vrei record eite gia MBR eite GPT alla pou einai gia unplugged
			FullParsehtml_code ='''<!DOCTYPE html>
			<html>
			<head>
			<meta charset="utf-8" />
			<title> Report </title>
			<meta name="viewport" content="width=device-width, initial-scale=1">
			<link rel="stylesheet" type="text/css" media="screen" href="style.css" />
			</head>
			<body>
			<div class="wrapper">
			<div class="header">                        
			<H1> Report </H1>
			</div>'''

			FullParsehtml_code += f'\n<table align="center"> \n <caption><h2><b>Full Report for all connected devices</b></h2></caption>'
			FullParsehtml_code += '\n<tr style="background-color:DarkGrey"> \n <th>Media S/N</th> \n <th>Manufacturer</th> \n <th>Model</th> \n <th>First connected Timestamp (UTC)</th> \n <th>Last connected Timestamp (UTC)</th> \n <th>Every VSN ever existed on the media</th>'
			
			for serial in AllPluggedInSerials:
				FullParsehtml_code += f'\n<tr> \n <td>{serial}</td> \n<td>{EachPluggedDeviceDict[serial][0]}</td> \n <td>{EachPluggedDeviceDict[serial][1]}</td> \n<td>{EachPluggedDeviceDict[serial][2]}</td> \n<td>{EachPluggedDeviceDict[serial][3]} </td>\n <td>'				
				for i in range(len(EachPluggedDeviceDict[serial][4])):
					if EachPluggedDeviceDict[serial][4][i] == '-' or EachPluggedDeviceDict[serial][4][i] == 'Unknown Volume Type':
						continue		
					FullParsehtml_code += f'{EachPluggedDeviceDict[serial][4][i]} '
				FullParsehtml_code += '</td>'

			FullParsehtml_code += f'\n</table> \n<br>\n<p align="left" style="color:white">\n <u style="font-size:20px"> Complete Log Timeline</u> <br>\n From: {FullParseLogStartTime} <br>\nTo: {FullParseLogEndTime} <br>\n</p> \n<br> <br> <br>\n<div class="push"></div> \n</div> \n <div class="footer">--Partition%4DiagnosticParser Ver. 1.3.0</div>\n</body>\n</html>'
			FullParsecss_code = '''table{border-collapse:collapse;}
			th{text-align:center;background-color:#4a4343;color=white;}
			table,th,td{border:1px solid #000;}
			tr{text-align:center;background-color:#595555; color:white;}

			html, body {
			height: 100%;
			margin: 0;
			}

			.wrapper {
			min-height: 100%;
			background-color: #4a4349;
			/* Equal to height of footer */
			/* But also accounting for potential margin-bottom of last child */
			margin-bottom: -50px;
			font-family: "Courier New", sans-serif;
			color=white;

			}

			.header{
			background-color: dark grey;
			color=white;
			}
			.header h1 {
			text-align: center;
			font-family: "Courier New", sans-serif;
			color=red;
			}
			.push {
			height: 50px;
			background-color: #4a4349;
			}
			.footer {
			height: 50px;
			background-color: #4a4349;
			color=white;
			text-align: right;
			}		'''
			if values['-INSAVE-'] == defaultOutputPathText: #dhl tickare gia html report alla den epelekse fakelo gia save alla afise to default
				try:
					with open('Full_Report.html', 'w', encoding='utf8') as fout:
						fout.write(FullParsehtml_code)
					with open('style.css', 'w', encoding= 'utf8') as cssout:
						cssout.write(FullParsecss_code)
				except:
					FullParseHTMLWritten = False
			else: #dhl exei dwsei output save folder
				try:
					with open(f"{values['-INSAVE-']}/Full_Report.html", 'w', encoding='utf8') as fout:
						fout.write(FullParsehtml_code)
					with open(f"{values['-INSAVE-']}/style.css", 'w', encoding= 'utf8') as cssout:
						cssout.write(FullParsecss_code)
				except:
					FullParseHTMLWritten = False
			
			if FullParseHTMLWritten:				
				if values['-DISKSN-'] == '':					
					print('Initializing parsing')
					print('.................')										
					print('Parsing complete')
					print('.................')
					print('Initializing analysis')
					print('.................')
					print('Analysis complete')
					print('.................')
					print('---------------------------')
					print('Analysis Results')
					print('---------------------------')
					print()
					print()
					print(f'Log timeline:')
					print(f'Start: {FullParseLogStartTime}')
					print(f'End: {FullParseLogEndTime}')
					print()
				print('Full analysis report for all connected devices completed')
				sg.PopupOK('Full report created succesfully!', title=':)', background_color='#2a363b')
			else:
				print('Full report has been created succesfully but was unable to be written to disk\nCheck write permissions for the selected output folder!')
		else:
			sg.PopupOK('Not a PartitionDiagnostic EVTX log file!', title='!!!', background_color='#2a363b')
			window['-IN-'].update('')
			window['-DISKSN-'].update('')
	except Exception as e:
		print(e)
		sg.PopupOK('Error parsing the chosen EVTX log file!', title='Error', background_color='#2a363b')					
		window['-IN-'].update('')
		window['-DISKSN-'].update('')

#---menu definition
menu_def = [['File', ['Exit']],
			['Help', ['Documentation', 'About']],] 
#---layout definition
DiskSNFrameLayout = [[sg.Text('Give the device S/N to search for', background_color='#2a363b')],
					[sg.In(key='-DISKSN-')]]

InputFrameLayout = [[sg.Text('Choose the EVTX Log File to Analyse', background_color='#2a363b')],
					[sg.In(key='-IN-', readonly=True, background_color='#334147'), sg.FileBrowse(file_types=(('Windows Event Log', '*.evtx'),))]]

defaultOutputPathText = 'Default report location is the exe folder'
OutputSaveFrameLayout = [[sg.Text('Choose folder to save the report file', background_color='#2a363b')],
					[sg.In(key='-INSAVE-', readonly=True, background_color='#334147', default_text = defaultOutputPathText, text_color='grey'), sg.FolderBrowse(key='-SAVEBTN-', disabled=True, enable_events=True)]]

col_layout = [[sg.Frame('Device Serial Number', DiskSNFrameLayout, background_color='#2a363b')],
				[sg.Frame('Input EVTX File', InputFrameLayout, background_color='#2a363b', pad=((0,0),(0,65))) ],				
				[sg.Frame('Output Report File Location', OutputSaveFrameLayout, background_color='#2a363b')],
				[sg.Checkbox('HTML Report', background_color='#2a363b', enable_events=True, key='-HTMLCHK-'), sg.Checkbox('CSV Report', background_color='#2a363b', enable_events=True, key='-CSVCHK-')],
				[sg.Checkbox('Full report for all connected devices in the Event Log', background_color='#2a363b', enable_events=True, key='-FULLCHK-')],
				[sg.Button('Exit', size=(7,1)), sg.Button('Analyze', size=(7,1))]]

#---GUI Definition
layout = [[sg.Menu(menu_def, key='-MENUBAR-')],
			[sg.Column(col_layout, element_justification='c',background_color='#2a363b'), sg.Frame('Verbose Analysis',
			[[sg.Output(size=(70,25), background_color='#334147', text_color='#fefbd8')]], background_color='#2a363b')],
			[sg.Text('Partition%4DiagnosticParser Ver. 1.3.0', background_color='#2a363b', text_color='#b2c2bf')]]

window = sg.Window('Partition%4DiagnosticParser', layout, background_color='#2a363b') 

#---run
while True:
	event, values = window.read()
	# print(event, values)
	if event in (sg.WIN_CLOSED, 'Exit'):
		break
#---menu events
	if event == 'Documentation':
		try:
			webbrowser.open_new('https://github.com/theAtropos4n6')
		except:
			sg.PopupOK('Visit https://github.com/theAtropos4n6 for documentation', title='Documentation', background_color='#2a363b')
	if event == 'About':
		sg.PopupOK('EVTX Partition Diagnostic Parser Ver. 1.3.0 \n\n --DKats 2020', title='-About-', background_color='#2a363b')
#---checkbox events (to enable/disable the output save button)
	if event == '-HTMLCHK-' or event == '-CSVCHK-' or event == '-FULLCHK-':
		if values['-HTMLCHK-'] == False and values['-CSVCHK-'] == False and values['-FULLCHK-'] == False:
			window['-SAVEBTN-'].update(disabled=True)
			window['-INSAVE-'].update(value=defaultOutputPathText, text_color='grey')
		else:
			window['-SAVEBTN-'].update(disabled=False)
			window['-INSAVE-'].update(text_color='#000000')			
#---analyze event
	if event == "Analyze":
		if values['-IN-'] == '':			
			sg.PopupOK('Please choose an EVTX file to parse!', title='!', background_color='#2a363b')
			window['-DISKSN-'].update('')			
		else:
			if values['-DISKSN-'] == '': # den exei valei S/N na dei
				if values['-FULLCHK-']: # an exei checkarei na dei to full report kai epomenos den exei valei disk S/N to dexomai k parsarw mono to evtx gia oles tis plugged in syskeues
					fullparse()
				else:
					sg.PopupOK('Please give a device S/N to search for!', title='Warning', background_color='#2a363b')
					window['-IN-'].update('')
			else:
				try:
					filename = values['-IN-']
					# onlyname = fullpath[fullpath.rfind('/')+1:]					
					parser = PyEvtxParser(filename)					
					records_dict = {}
					serial = values['-DISKSN-']
					isDiskPlugedin = False
					isDiskMBR = False
					DataListsList = []
					AllSNs = []
					CSVTicked = False
					HTMLTicked = False
					IsPartitionDiagnosticEVTX = True
					html_code ='''<!DOCTYPE html>
					<html>
					<head>
					<meta charset="utf-8" />
					<title> Report </title>
					<meta name="viewport" content="width=device-width, initial-scale=1">
					<link rel="stylesheet" type="text/css" media="screen" href="style.css" />
					</head>
					<body>
					<div class="wrapper">
					<div class="header">                        
					<H1> Report </H1>
					</div>'''

					html_code += f'\n<table align="center"> \n <caption><h2><b>Partition%4Diagnostic.evtx ANALYSIS REPORT for device with S/N: {serial}</b></h2></caption>'
					html_code += '\n<tr style="background-color:DarkGrey"> \n <th>EventRecordID</th> \n <th>connected Timestamp (UTC)</th> \n <th>Manufacturer</th> \n <th>Model</th> \n <th>Volume 1 Serial Number</th> \n <th>Volume 2 Serial Number</th> \n <th>Volume 3 Serial Number</th> \n <th>Volume 4 Serial Number</th> \n <th>Flag</th>'
					
					for record in parser.records_json():
						data = json.loads(record['data'])	
						records_dict[(data['Event']['System']['EventRecordID'])] = data #add gia na sortaro meta me event id kai na exo xronologika sosti seira
						if data['Event']['System']['EventID'] != 1006:
							IsPartitionDiagnosticEVTX = False
					if IsPartitionDiagnosticEVTX: #parsarw giati einai to PartitionDiagnostic evtx afou ola ta EventID einai 1006
						print('Initializing parsing')
						print('.................')
						LogStartTime = records_dict[1]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
						LogEndTime = records_dict[len(records_dict)]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
						print('Parsing complete')
						print('.................')
						print('Initializing analysis')
						print('.................')

						for i in sorted(records_dict.keys()):
							if serial == records_dict[i]['Event']['EventData']['SerialNumber']:
								isDiskPlugedin = True
								if records_dict[i]['Event']['EventData']['Vbr0'] != '': #prwth fora pou egine plugged in ws MBR scheme disk
									isDiskMBR = True			
									SN0ToCheck = volumeSNParser(records_dict[i], 'Vbr0', records_dict[i]['Event']['EventData']['Vbr0'][6:18])
									if records_dict[i]['Event']['EventData']['Vbr1'] != '':
										SN1ToCheck = volumeSNParser(records_dict[i], 'Vbr1', records_dict[i]['Event']['EventData']['Vbr1'][6:18])
									else:
										SN1ToCheck = '-'
									if records_dict[i]['Event']['EventData']['Vbr2'] != '':
										SN2ToCheck = volumeSNParser(records_dict[i], 'Vbr2', records_dict[i]['Event']['EventData']['Vbr2'][6:18])
									else:
										SN2ToCheck = '-'
									if records_dict[i]['Event']['EventData']['Vbr3'] != '':
										SN3ToCheck = volumeSNParser(records_dict[i], 'Vbr3', records_dict[i]['Event']['EventData']['Vbr3'][6:18])
									else:
										SN3ToCheck = '-'
									break
						print('Analysis complete')
						print('.................')
						print('---------------------------')
						print('Analysis Results')
						print('---------------------------')
						print()
						print()
						print(f'Log timeline:')
						print(f'Start: {LogStartTime}')
						print(f'End: {LogEndTime}')
						print()
						if isDiskPlugedin == False:
							print(f'Device with serial number {serial} was never connected to the computer.')
							if values['-HTMLCHK-'] or values['-CSVCHK-']:
								print('No report was created for the specific S/N')
							if values['-FULLCHK-']:
								fullparse()
						elif isDiskMBR == False:
							print(f'Device with serial number {serial} uses a GPT partitioning scheme throughout the log! No volume info available in the evtx log.')
							if values['-HTMLCHK-'] or values['-CSVCHK-']:
								print('No report was created for the specific S/N')
							if values['-FULLCHK-']:
								fullparse()
						else:							
							for i in sorted(records_dict.keys()):
								if serial == records_dict[i]['Event']['EventData']['SerialNumber']:
									if records_dict[i]['Event']['EventData']['PartitionStyle'] == 1: # o diskos exei metatrapei se GPT scheme kai exei ginei plugged in (sto plug out to partitionStyle ginetai pali 0)
										DataList = []
										DataList.append(records_dict[i]['Event']['System']['EventRecordID'])	# event Record = integer
										DataList.append(records_dict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')) # timestamp = string
										DataList.append(records_dict[i]['Event']['EventData']['Manufacturer']) # Manufacturer = string
										DataList.append(records_dict[i]['Event']['EventData']['Model']) # Model = string
										DataList.append('GPT partitioning scheme - No VSN info')
										DataList.append('GPT partitioning scheme - No VSN info')
										DataList.append('GPT partitioning scheme - No VSN info')
										DataList.append('GPT partitioning scheme - No VSN info')
										DataListsList.append(DataList)
									else: # an diskos einai MBR
										if records_dict[i]['Event']['EventData']['Vbr0'] == '' and records_dict[i]['Event']['EventData']['Vbr1'] == '' and records_dict[i]['Event']['EventData']['Vbr2'] == '' and records_dict[i]['Event']['EventData']['Vbr3'] == '': # diladi exei ginei unplugged opote next iteration
											continue # giati einai unplugged ws MBR
										else:
											DataList = []			
											DataList.append(records_dict[i]['Event']['System']['EventRecordID'])	# event Record = integer
											DataList.append(records_dict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')) # timestamp = string
											DataList.append(records_dict[i]['Event']['EventData']['Manufacturer']) # Manufacturer = string
											DataList.append(records_dict[i]['Event']['EventData']['Model']) # Model = string
											SN0 = volumeSNParser(records_dict[i], 'Vbr0', records_dict[i]['Event']['EventData']['Vbr0'][6:18])
											DataList.append(SN0) # Volume Serial Number = string
											if SN0 not in AllSNs: AllSNs.append(SN0) # lista pou krataei ola ta unique SNs olwn twn volumes gia na ta parousiasw sto telos se textbox ws quick result
											if records_dict[i]['Event']['EventData']['Vbr1'] == '':
												DataList.append('-')
											else:
												SN1 = volumeSNParser(records_dict[i], 'Vbr1', records_dict[i]['Event']['EventData']['Vbr1'][6:18])
												DataList.append(SN1)
												if SN1 not in AllSNs: AllSNs.append(SN1)
											if records_dict[i]['Event']['EventData']['Vbr2'] == '':
												DataList.append('-')
											else:
												SN2 = volumeSNParser(records_dict[i], 'Vbr2', records_dict[i]['Event']['EventData']['Vbr2'][6:18])
												DataList.append(SN2)
												if SN2 not in AllSNs: AllSNs.append(SN2)
											if records_dict[i]['Event']['EventData']['Vbr3'] == '':
												DataList.append('-')
											else:
												SN3 = volumeSNParser(records_dict[i], 'Vbr3', records_dict[i]['Event']['EventData']['Vbr3'][6:18])
												DataList.append(SN3)
												if SN3 not in AllSNs: AllSNs.append(SN3)
										DataListsList.append(DataList) # dhmiourgw mia lista me oles tis listes tou ka8e entry, to len ths einai poses fores o diskos egine plugged in sto pc
							
							for DList in DataListsList:
								if DList[4] == 'GPT partitioning scheme - No VSN info': # exw entry tou GPT partition
									html_code += f'\n<tr> \n <td>{DList[0]}</td> \n <td>{DList[1]}</td> \n<td>{DList[2]}</td> \n<td>{DList[3]} </td> \n <td>{DList[4]}</td> \n<td>{DList[5]}</td>\n <td>{DList[6]}</td> \n<td>{DList[7]}</td> \n<td> GPT media </td>'
								elif DList[4] == SN0ToCheck and DList[5] == SN1ToCheck and DList[6] == SN2ToCheck and DList[7] == SN3ToCheck: # an idia me to check grafw to row sto html kai synexizw
									html_code += f'\n<tr> \n <td>{DList[0]}</td> \n <td>{DList[1]}</td> \n<td>{DList[2]}</td> \n<td>{DList[3]} </td> \n <td>{DList[4]}</td> \n<td>{DList[5]}</td>\n <td>{DList[6]}</td> \n<td>{DList[7]}</td> \n<td> - </td>'
								else: # an den einai idia me to check tote exei ginei format kai kapoio apo ta volume SNs allakse opote kanw add grammi highlightened kai kanw update tis nees times SN gia check apo edw kai pera
									html_code += f'\n<tr style="background-color:#c23c53"> \n <td>{DList[0]}</td> \n <td>{DList[1]}</td> \n<td>{DList[2]}</td> \n<td>{DList[3]} </td> \n <td>{DList[4]}</td> \n<td>{DList[5]}</td>\n <td>{DList[6]}</td> \n<td>{DList[7]}</td> \n<td> <b>VSN Change - Possible Format Action</b> </td>'
									SN0ToCheck = DList[4]
									SN1ToCheck = DList[5]
									SN2ToCheck = DList[6]
									SN3ToCheck = DList[7]
							
							if values['-HTMLCHK-']: # an exei tickarei to checkbox gia html report
								HTMLTicked = True
								HTMLWritten = True
								html_code += f'\n</table> \n<br>\n<p align="left" style="color:white">\n <u style="font-size:20px"> Complete Log Timeline</u> <br>\n From: {FullParseLogStartTime} <br>\nTo: {FullParseLogEndTime} <br>\n</p> \n<br> <br> <br> \n<div class="push"></div> \n</div> \n <div class="footer">--Partition%4DiagnosticParser Ver. 1.3.0</div>\n</body>\n</html>'
								css_code = '''table{border-collapse:collapse;}
								th{text-align:center;background-color:#4a4343;color=white;}
								table,th,td{border:1px solid #000;}
								tr{text-align:center;background-color:#595555; color:white;}

								html, body {
								height: 100%;
								margin: 0;
								}

								.wrapper {
								min-height: 100%;
								background-color: #4a4349;
								/* Equal to height of footer */
								/* But also accounting for potential margin-bottom of last child */
								margin-bottom: -50px;
								font-family: "Courier New", sans-serif;
								color=white;

								}

								.header{
								background-color: dark grey;
								color=white;
								}
								.header h1 {
								text-align: center;
								font-family: "Courier New", sans-serif;
								color=red;
								}
								.push {
								height: 50px;
								background-color: #4a4349;
								}
								.footer {
								height: 50px;
								background-color: #4a4349;
								color=white;
								text-align: right;
								}		'''
								if values['-INSAVE-'] == defaultOutputPathText: #dhl tickare gia html report alla den epelekse fakelo gia save alla afise to default
									try:
										with open(f'{serial}_report.html', 'w', encoding='utf8') as fout:
											fout.write(html_code)
										with open('style.css', 'w', encoding= 'utf8') as cssout:
											cssout.write(css_code)
									except:
										HTMLWritten = False
								else: #dhl exei dwsei output save folder
									try:
										with open(f"{values['-INSAVE-']}/{serial}_report.html", 'w', encoding='utf8') as fout:
											fout.write(html_code)
										with open(f"{values['-INSAVE-']}/style.css", 'w', encoding= 'utf8') as cssout:
											cssout.write(css_code)
									except:
										HTMLWritten = False
							
							if values['-CSVCHK-']:
								CSVTicked = True
								CSVWritten = True
								if values['-INSAVE-'] == defaultOutputPathText: #dhl tickare gia html report alla den epelekse fakelo gia save alla afise to default
									try:
										with open(f'{serial}_report.csv', 'w', encoding='utf8') as fout:
											for DList in DataListsList:
												fout.write(f'{DList[0]},{DList[1]},{DList[2]},{DList[3]},{DList[4]},{DList[5]},{DList[6]},{DList[7]}\n')
									except:
										CSVWritten = False
								else:
									try:
										with open(f"{values['-INSAVE-']}/{serial}_report.csv", 'w', encoding='utf8') as fout:
											for DList in DataListsList:
												fout.write(f'{DList[0]},{DList[1]},{DList[2]},{DList[3]},{DList[4]},{DList[5]},{DList[6]},{DList[7]}\n')
									except:
										CSVWritten = False
							print(f'Device with serial number {serial} has been (possibly) connected {len(DataListsList)} times.')
							print(f'Device has had {len(AllSNs)} unique volume serial number(s) in total throughout the log:')
							for SN in AllSNs:
								print(SN)
							if HTMLTicked:
								if HTMLWritten:
									print('The html report has been created succesfully')
								else:
									print('The html report has been created succesfully but was unable to be written to disk\nCheck write permissions for the selected output folder!')
							if CSVTicked:
								if CSVWritten:
									print('The csv report has been created succesfully')
								else:
									print('The csv report has been created succesfully but was unable to be written to disk\nCheck write permissions for the selected output folder!')
							if values['-FULLCHK-']: # an exei checkarei full report
								fullparse()
							sg.PopupOK('ANALYSIS SUCCESSFUL!', title=':)', background_color='#2a363b')
					else:
						sg.PopupOK('Not a PartitionDiagnostic EVTX log file!', title='!!!', background_color='#2a363b')
						window['-IN-'].update('')
						window['-DISKSN-'].update('')
				except:
					sg.PopupOK('Error parsing the chosen EVTX log file!', title='Error', background_color='#2a363b')					
					window['-IN-'].update('')
					window['-DISKSN-'].update('')			
window.close()
