import time, os, sys, os.path
from selenium import webdriver
from xlwt import Workbook

options = []
website = []
results = []

driver = webdriver.Chrome("chromedriver/chromedriver.exe")

def Welcome():
	os.system('cls')
	del options[:]
	del website[:]
	del results[:]
	print '+=================================+'
	print '| WELKOM BIJ DE UI TESTER VAN SAM |'
	print '+=================================+'

	print '\n1: Stel je UI test op'

	choice = raw_input('\nUw keuze: ')

	if choice == '1':
		WebPage()
	else:
		print 'Kies nummer 1.'
		time.sleep(3)
		Welcome()

def WebPage():
	os.system('cls')
	print '+====================+'
	print '| STEL JE UI TEST OP |'
	print '+====================+'
	print '\nWelke website wilt u testen?'
	page = raw_input('\n> ')
	page = 'http://' + page
	website.append(page)
	Options()

def SearchInputOption(type, name, searchterm):
	if type == 'class':
		try:
			sbox = driver.find_element_by_class_name(name)
			sbox.clear()
			sbox.send_keys(searchterm)
			result = "Passed: Input field with class '" + name + "' and search term '" + searchterm + "'"
			results.append(result)
			time.sleep(3)
		except:
			result = "Failed: Input field with class '" + name + "' and search term '" + searchterm + "'"
			results.append(result)
	elif type == 'id':
		try:
			sbox = driver.find_element_by_id(name)
			sbox.clear()
			sbox.send_keys(searchterm)
			result = "Passed: Input field with ID '" + name + "' and search term '" + searchterm + "'"
			results.append(result)
			time.sleep(3)
		except:
			result = "Failed: Input field with ID '" + name + "' and search term '" + searchterm + "'"
			results.append(result)

def SearchInput():
	data = 'searchInput'
	os.system('cls')
	print '+======================================+'
	print '| VIA WAT WILT U HET INPUT VELD ZOEKEN |'
	print '+======================================+'
	print '1: Class name'
	print '2: ID'

	type = raw_input('\nUw keuze: ')

	if type == '1':
		className = raw_input('Op welke class name wilt u zoeken: ')
		data = data + '|' + 'class'
		data = data + '|' + className
	elif type == '2':
		id = raw_input('Op welk ID wilt u zoeken: ')
		data = data + '|' + 'id'
		data = data + '|' + id
	else:
		print 'Gelieve 1 of 2 te typen.'
		time.sleep(2)
		os.system('cls')
		SearchInput()
	searchterm = raw_input('Wat wilt u in het input veld typen: ')
	data = data + '|' + searchterm
	options.append(data)
	StartScreen()

def TextOption(type, name, searchterm):
	if type == 'tag':
		try:
			text = driver.find_element_by_tag_name(name).text;
			if searchterm in text:
				result = "Passed: text '" + searchterm + "' is in tag '" + name + "'"
				results.append(result)
				time.sleep(2)
			else:
				result = "Failed: text '" + searchterm + "' is not in tag '" + name + "'"
				results.append(result)
		except:
			result = "Failed: text '" + searchterm + "' is not in tag '" + name + "'"
			results.append(result)

def Text():
	data = 'text'
	os.system('cls')
	print '+================================+'
	print '| VIA WAT WILT U DE TEKST ZOEKEN |'
	print '+================================+'
	print '1: Tag name'

	type = raw_input('\nUw keuze: ')

	if type == '1':
		tagName = raw_input('Op welke tag name wilt u zoeken: ')
		data = data + '|' + 'tag'
		data = data + '|' + tagName
	else:
		print 'Gelieve 1 te typen.'
		os.system('cls')
		Text()

	searchterm = raw_input('Welke tekst moet zeker in de tag staan: ')
	data = data + '|' + searchterm
	options.append(data)
	StartScreen()

def RunOptions():
	driver.get(website[0])
	time.sleep(2)
	for option in options:
		optionsList = option.split('|')
		if optionsList[0] == 'searchInput':
			SearchInputOption(optionsList[1], optionsList[2], optionsList[3])
		elif optionsList[0] == 'click':
			ClickOption(optionsList[1], optionsList[2])
		elif optionsList[0] == 'text':
			TextOption(optionsList[1], optionsList[2], optionsList[3])
	Summary()

def Options():
	os.system('cls')
	print '+================+'
	print '| KIES EEN OPTIE |'
	print '+================+'
	print '\n1: Klik op button/link'
	print '2: Type iets in input field'
	print '3: Check of tekst present is'
	choice = raw_input('\nUw keuze: ')
	if choice == '1':
		Click()
	elif choice == '2':
		SearchInput()
	elif choice == '3':
		Text()
	else:
		print 'Gelieve 1 of 2 te typen.'
		time.sleep(2)
		os.system('cls')
		Options()

def StartScreen():
	os.system('cls')
	print '+====================+'
	print '| WAT WILT U NU DOEN |'
	print '+====================+'
	print '\n1: nog een test toevoegen'
	print '2: begin de test'
	print '3: ingevoerde data aanpassen'

	choice = raw_input('\nUw keuze: ')

	if choice == '1':
		Options()
	elif choice == '2':
		RunOptions()
	elif choice == '3':
		Edit()
	else:
		print 'Gelieve 1 of 2 te typen.'
		time.sleep(2)
		os.system('cls')
		StartScreen()

def Click():
	data = 'click'
	os.system('cls')
	print '+=================================+'
	print '| VIA WAT WILT U DE BUTTON ZOEKEN |'
	print '+=================================+'
	print '1: Class name'
	print '2: ID'

	type = raw_input('\nUw keuze: ')

	if type == '1':
		className = raw_input('Op welke class name wilt u zoeken: ')
		data = data + '|' + 'class'
		data = data + '|' + className
	elif type == '2':
		id = raw_input('Op welk ID wilt u zoeken: ')
		data = data + '|' + 'id'
		data = data + '|' + id
	else:
		print 'Gelieve 1 of 2 te typen.'
		time.sleep(2)
		os.system('cls')
		Click()
	options.append(data)
	StartScreen()

def ClickOption(type, name):
	if type == 'class':
		try:
			submit = driver.find_element_by_class_name(name)
			submit.click()
			result = "Passed: Succesfully clicked on element with class name '" + name + "'"
			results.append(result)
			time.sleep(3)
		except:
			result = "Failed: Unsuccesfully clicked on element with class name '" + name + "'"
			results.append(result)
	elif type == 'id':
		try:
			submit = driver.find_element_by_id(name)
			submit.click()
			result = "Passed: Succesfully clicked on element with ID '" + name + "'"
			results.append(result)
			time.sleep(3)
		except:
			result = "Failed: Unsuccesfully clicked on element with ID '" + name + "'"
			results.append(result)

def Summary():
	os.system('cls')
	print '+==========================+'
	print '| OVERZICHT VAN DE UI TEST |'
	print '+==========================+\n'
	for result in results:
		print result
	print '\n+=====================+'
	print '| WAT WILT U NU DOEN? |'
	print '+=====================+\n'

	print '1: Voer dezelfde test opnieuw uit'
	print '2: Pas de data aan van deze test'
	print '3: Maak een nieuwe test (huidige test wordt verwijderd)'
	print '4: Zet de resultaten om naar een xls bestand'
	print '5: Exporteer test naar .py bestand'
	print '6: Stop het programma'

	choice = raw_input('\nUw keuze: ')

	if choice == '1':
		del results[:]
		RunOptions()
	elif choice == '2':
		Edit()
	elif choice == '3':
		Welcome()
	elif choice == '4':
		wb = Workbook()
		sheet1 = wb.add_sheet('Sheet 1')
		sheet1.row(0).height_mismatch = True
		sheet1.row(0).height = 600
		sheet1.write(0,0,'Gebruikte functie')
		sheet1.col(0).width = 7000
		sheet1.write(0,1,'Class name van element')
		sheet1.col(1).width = 7000
		sheet1.write(0,2,'ID van element')
		sheet1.col(2).width = 7000
		sheet1.write(0,3,'Tag name')
		sheet1.col(3).width = 7000
		sheet1.write(0,4,'Search term')
		sheet1.col(4).width = 7000
		sheet1.write(0,5,'Resultaat')
		sheet1.col(5).width = 7000

		for index, option in enumerate(options):
			optionsList = option.split('|')
			if optionsList[0] == 'searchInput':
				sheet1.write(index+1,0,optionsList[0])
				if optionsList[1] == 'class':
					sheet1.write(index+1,1,optionsList[2])
					sheet1.write(index+1,2,'NVT')
				elif optionsList[1] == 'id':
					sheet1.write(index+1,2,optionsList[2])
					sheet1.write(index+1,1,'NVT')
				sheet1.write(index+1,3,'NVT')				
				sheet1.write(index+1,4,optionsList[3])

			elif optionsList[0] == 'click':
				sheet1.write(index+1,0,optionsList[0])
				if optionsList[1] == 'class':
					sheet1.write(index+1,1,optionsList[2])
					sheet1.write(index+1,2,'NVT')
				elif optionsList[1] == 'id':
					sheet1.write(index+1,2,optionsList[2])
					sheet1.write(index+1,1,'NVT')
				sheet1.write(index+1,3,'NVT')
				sheet1.write(index+1,4,'NVT')

			elif optionsList[0] == 'text':
				sheet1.write(index+1,0,optionsList[0])
				if optionsList[1] == 'tag':
					sheet1.write(index+1,3,optionsList[2])
					sheet1.write(index+1,1,'NVT')
					sheet1.write(index+1,2,'NVT')
				sheet1.write(index+1,4,optionsList[3])
		for index, result in enumerate(results):
			sheet1.write(index+1,5,result)		
		
		SaveXls(wb)
		
		Summary()
		
	elif choice == '5':
		ExportPython(options, website)

	elif choice == '6':
		print 'exiting'
		sys.exit()
	else:
		print 'Gelieve 1, 2, 3, 4 of 5 te typen.'
		time.sleep(2)
		Summary()

def SaveXls(wb):
	print '\nAls Welke naam wilt u het xls bestand opslaan?'
	fileName = raw_input('\n> ')
	fileName = fileName + ".xls"
	if os.path.isfile(fileName): 
		print '\nDe naam bestaat al, wilt u het bestand overschrijven?'
		print '\n1: Ja'
		print '2: Nee'

		choice = raw_input('\n> ')

		if choice == '1':
			try:
				wb.save(fileName)
				print '\nSuccesvol opgeslaan.'
				time.sleep(3)
			except:
				print '\nEr is iets fout gegaan met het opslaan.'
				time.sleep(3)
		elif choice == '2':
			SaveXls(wb)
	else:
		try:
			wb.save(fileName)
			print '\nSuccesvol opgeslaan.'
			time.sleep(3)
		except:
			print '\nEr is iets fout gegaan met het opslaan.'
			time.sleep(3)

def Edit():
	os.system('cls')
	print '+==============================+'
	print '| WELKE OPTIE WILT U AANPASSEN |'
	print '+==============================+\n'

	print '1: Ga terug'
	print '2: Verander website'
	for index, option in enumerate(options):
		print str(index+3) + ": " + option

	choice = raw_input('\nUw keuze: ')

	if choice == '1':
		StartScreen()
	if choice == '2':
		EditWebsite()
	for index, option in enumerate(options):
		if choice == str(index+3):
			EditOptions(index)

def EditWebsite():
	print "De website die u nu hebt is '" + website[0] + "'"
	print 'Verander naar:'

	page = raw_input('\n> ')
	try:
		website[0] = 'http://' + page
		print 'Succesvol opgeslagen.'
		time.sleep(2)
	except:
		print 'Er is iets fout gegaan met het opslagen.'
		time.sleep(2)
	Edit()
	

def EditOptions(index):
	option = options[index]
	optionsList = option.split('|')
	if optionsList[0] == 'click':
		print '1(type): ' + optionsList[1]
		print '2(naam): ' + optionsList[2]
		print '3: Ga terug'
		choice = raw_input('\nUw keuze: ')
		if choice == '1':
			EditType(optionsList, index)
		elif choice == '2':
			EditName(optionsList, index)
		elif choice == '3':
			Edit()
		else:
			print 'Gelieve 1, 2 of 3 te typen'
			time.sleep(2)
			EditOptions(index)
	elif optionsList[0] == 'searchInput':
		print '1(type): ' + optionsList[1]
		print '2(naam): ' + optionsList[2]
		print '3(searchterm): ' + optionsList[3]
		print '4: Ga terug'
		choice = raw_input('\nUw keuze: ')
		if choice == '1':
			EditType(optionsList, index)
		elif choice == '2':
			EditName(optionsList, index)
		elif choice == '3':
			EditSearchterm(optionsList, index)
		elif choice == '4':
			Edit()
		else:
			print 'Gelieve 1, 2, 3 of 4 te typen'
			time.sleep(2)
			EditOptions(index)
	elif optionsList[0] == 'text':
		print '1(naam): ' + optionsList[2]
		print '2(searchterm): ' + optionsList[3]
		print '3: Ga terug'
		choice = raw_input('\nUw keuze: ')
		if choice == '1':
			EditName(optionsList, index)
		elif choice == '2':
			EditSearchterm(optionsList, index)
		elif choice == '3':
			Edit()
		else:
			print 'Gelieve 1, 2 of 3 te typen.'
			time.sleep(2)
			EditOptions(index)

def EditName(optionsList, index):
	print 'U wilt ' + optionsList[2] + " veranderen naar:"
	name = raw_input('\n> ')

	try:
		optionsList[2] = name		
		options[index] = '|'.join(optionsList)
		print 'Succesvol opgeslagen.'
		time.sleep(2)
	except:
		print 'Er is iets fout gegaan met het opslagen.'
		time.sleep(2)
	EditOptions(index)

def EditType(optionsList, index):
	print 'U wilt ' + optionsList[1] + " veranderen naar:"
	print '\n1: Class name'
	print '2: ID'

	choice = raw_input('\nUw keuze: ')

	if choice == '1':
		try:
			optionsList[1] = 'class'
			options[index] = '|'.join(optionsList)
			print 'Succesvol opgeslagen.'
			time.sleep(2)
		except:
			print 'Er is iets fout gegaan met het opslagen.'
			time.sleep(2)
		EditOptions(index)
	elif choice == '2':
		try:
			optionsList[1] = 'id'
			options[index] = '|'.join(optionsList)
			print 'Succesvol opgeslagen.'
			time.sleep(2)
		except:
			print 'Er is iets fout gegaan met het opslagen.'
			time.sleep(2)
		EditOptions(index)
	else:
		print 'Gelieve 1 of 2 te typen.'
		time.sleep(2)
		EditType(optionsList, index)
		
def EditSearchterm(optionsList, index):
	print 'U wilt ' + optionsList[3] + " veranderen naar:"

	searchterm = raw_input('\n> ')

	try:
		optionsList[3] = searchterm
		options[index] = '|'.join(optionsList)
		print 'Succesvol opgeslagen.'
		time.sleep(2)
	except:
		print 'Er is iets fout gegaan met het opslagen.'
		time.sleep(2)
	EditOptions(index)

def ExportPython(options, website):
	print '\nAls Welke naam wilt u het py bestand opslaan?'
	fileName = raw_input('\n> ')
	fileName = fileName + ".py"
	if os.path.isfile(fileName):
		print '\n de naam bestaat al, wilt u het bestand overschrijven?'
		print '\n1: Ja'
		print '2: Nee'
		choice = raw_input('\n> ')

		if choice == '1':
			python_file = open(fileName, 'w+')
		elif choice == '2':
			print 'Geef een nieuwe naam in.'
			time.sleep(2)
			ExportPython(options, website)
		else:
			'Gelieve 1 of 2 te typen.'
			sleep.time(2)
			ExportPython(options, website)
	else:
		python_file = open(fileName, 'w+')

	python_file.write('import time, os\nfrom selenium import webdriver\n\n')
	python_file.write('driver = webdriver.Chrome("chromedriver/chromedriver.exe")\n\n')
	python_file.write("driver.get('" + website[0] + "')\n\n")
	for option in options:
		option = option.split('|')
		if option[0] == 'searchInput':
			if option[1] == 'class':
				python_file.write("try:\n	sbox = driver.find_element_by_class_name('" + option[2] + "')\n")
			elif option[1] == 'id':
				python_file.write("try:\n	sbox = driver.find_element_by_id('" + option[2] + "')\n")
			python_file.write("	sbox.send_keys('" + option[3] + "')\n 	print 'Passed'\n")
			python_file.write("except:\n	print 'Failed'\n")
		elif option[0] == 'click':
			if option[1] == 'class':
				python_file.write("try:\n	submit = driver.find_element_by_class_name('" + option[2] + "')\n")
			elif option[1] == 'id':
				python_file.write("try:\n	submit = driver.find_element_by_id('" + option[2] + "')\n")
			python_file.write("	submit.click()\n 	print 'Passed'\n")
			python_file.write("except:\n	print 'Failed'\n")
		elif option[0] == 'text':
			if option[1] == 'tag':
				python_file.write("try:\n	text = driver.find_element_by_tag_name('" + option[2] + "').text\n")
				python_file.write("	if '" + option[3] + "' in text:\n")
				python_file.write("		print 'Passed'\n	else:\n		print 'Failed'\n")
				python_file.write("except:\n		print 'Failed'")

	try:
		python_file.close()
		print '\n Succesvol opgeslagen.'
		time.sleep(2)
	except:
		print 'Er liep iets fout bij het opslagen.'
		time.sleep(2)
	
	Summary()

Welcome()