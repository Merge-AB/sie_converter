from pathlib import Path
import os
import re
import csv
import openpyxl
import datetime


# Sätter directory


path = os.getcwd()
print(path)
directory = path
directory2 = path + "\\Output-filer"
#os.chdir(directory2) för att sätta outputfilerna i rätt mapp
os.chdir(directory2)
os.getcwd() 

print("initiating...")

#Skapar en CSV fil
outputFil1 = open("TRANSACTIONS.csv", "w", newline="", encoding='cp437')
outputWriter1 = csv.writer(outputFil1, delimiter=",")
#Skriver toppraderna
trans_topp = ['serie', 'ver_nr', 'ver_datum', 'ver_text', 'reg_datum', 'konto', 'belopp', 'kostnadsställe', 'trans_typ', 'belopp_2']
outputWriter1.writerow(trans_topp)

#Skapar en CSV fil
outputFil_UB = open("UB.csv", "w", newline="", encoding='ANSI')
outputWriter_UB = csv.writer(outputFil_UB, delimiter=",")
#Skriver toppraderna
ub_topp = ['DEL', 'DEL', 'konto','UB', 'DEL']
outputWriter_UB.writerow(ub_topp)

#Skapar en CSV fil
outputFil_IB = open("IB.csv", "w", newline="", encoding='ANSI')
outputWriter_IB = csv.writer(outputFil_IB, delimiter=",")
#Skriver toppraderna
ib_topp = ['DEL', 'DEL', 'konto', 'IB']
outputWriter_IB.writerow(ib_topp)

#Skapar en CSV fil
outputFil_RES = open("RES.csv", "w", newline="", encoding='ANSI')
outputWriter_RES = csv.writer(outputFil_RES, delimiter=",")
#Skriver toppraderna
res_topp = ['DEL', 'DEL', 'konto','resultat', 'DEL']
outputWriter_RES.writerow(res_topp)

#Skapar en CSV fil
outputFil_PSALDO = open("PSALDO.csv", "w", newline="", encoding='ANSI')
outputWriter_PSALDO = csv.writer(outputFil_PSALDO, delimiter=",")
#Skriver toppraderna
psaldo_topp = ['DEL', 'DEL', 'datum', 'konto', 'UB', 'DEL']
outputWriter_PSALDO.writerow(psaldo_topp)

#Skapar en CSV fil
outputFil_KONTO = open("KONTO.csv", "w", newline="", encoding='ANSI')
outputWriter_KONTO = csv.writer(outputFil_KONTO, delimiter=",")
#Skriver toppraderna
konto_topp = ['kontonummer','kontonamn']
outputWriter_KONTO.writerow(konto_topp)

print('output files created...')

class Master:
	def __init__(self, line, outputWriter_UB, outputWriter_IB, outputWriter_RES, outputWriter_PSALDO, outputWriter_KONTO, outputWriter1):
		self.line = line
		self.outputWriter_UB = outputWriter_UB
		self.outputWriter_IB = outputWriter_IB
		self.outputWriter_RES = outputWriter_RES
		self.outputWriter_PSALDO = outputWriter_PSALDO
		self.outputWriter_KONTO = outputWriter_KONTO
		self.outputWriter1 = outputWriter1

	def UB(self):
		line = self.line
		outputWriter_UB = self.outputWriter_UB
		#Hittar alla Utgående Balans och sammanställer dem i en .csv
		if line.startswith("#UB 0"):
			line = line.replace(",","")
			splitted_UB_line = line.split()
			splitted_UB_line[-1] = splitted_UB_line[-1].rstrip()
			outputWriter_UB.writerow(splitted_UB_line)

	def IB(self):
		line = self.line
		outputWriter_IB = self.outputWriter_IB
		#Hittar alla Ingående Balans och sammanställer dem i en .csv
		if line.startswith("#IB 0"):
			line = line.replace(",","")
			splitted_IB_line = line.split()
			splitted_IB_line[-1] = splitted_IB_line[-1].rstrip()
			outputWriter_IB.writerow(splitted_IB_line)

	def RES(self):
		line = self.line
		outputWriter_RES = self.outputWriter_RES
		#Hittar alla Resultat per konto och sammanställer dem i en .csv
		if line.startswith("#RES"):
			line = line.replace(",","")
			splitted_RES_line = line.split()
			splitted_RES_line[-1] = splitted_RES_line[-1].rstrip()
			outputWriter_RES.writerow(splitted_RES_line)

	def PSALDO(self):
		line = self.line
		outputWriter_PSALDO = self.outputWriter_PSALDO
		#Hittar alla PSALDO och sammanställer dem i en .csv
		if line.startswith("#PSALDO 0"):
			line = line.replace(",","")
		# Hitta allt som är inom brackets och ta bort
			line2 = re.findall("{(.*?)}", line)
			str1 = ''
			for x in line2:
			   str1 += x
			str1 = '{'+str1+'}'
			line = line.split(str1)
			
			# Denna finns till så att elementen i fixed_trans-listan läggs till i korrekt ordning. Note: Det som heter new_string är egentligen en lista 
			fixed_line = []
			räknare = 0
			for x in line:
				new_string = x.split()
				fixed_line.append(new_string)

			# Sammanställer en lista som är korrekt formaterad. Denna konverteras till csv
			final_lista = []

			for x in fixed_line[0]:
				final_lista.append(x) 
			for x in fixed_line[1]:
				final_lista.append(x)

			outputWriter_PSALDO.writerow(final_lista)

	def KONTO(self):
		line = self.line
		outputWriter_KONTO = self.outputWriter_KONTO
		#Hittar alla kontonummer och kontobeskrivningar och sammanställer dem i en .csv
		if line.startswith("#KONTO"):
			line = line.replace(",","")
			splitted_KONTO_line = line.split()
			splitted_KONTO_line[-1] = splitted_KONTO_line[-1].rstrip()
			kontonamn = " "
			kontonamn = kontonamn.join(splitted_KONTO_line[2:])
			kontonamn = kontonamn.strip('"')
			splitted_KONTO_line = splitted_KONTO_line[:2]
			splitted_KONTO_line.append(kontonamn)
			outputWriter_KONTO.writerow(splitted_KONTO_line[1:])

	def VER(self):
		line = self.line
		outputWriter1 = self.outputWriter1
		#Hittar alla Verifikationer och sammanställer dem i en .csv
		if "#VER" in line:
			ver = line
			# Här formaterar vi en lista så att allt blir lättåtkomligt i kommande funktioner
			ver = ver.replace(",","")
			ver = ver.replace('\\',"") 
			#ver = ver.replace('"','')
			ver = ver.replace("'",'')

			lista = []
		
			ver = ver.split()
			for x in ver[1:]:
				lista.append(x)

			# Undersöker om sista elementet i listan är ett datum. Detta är centralt!
			date_string = lista[-1]
			format = "%Y%m%d"
			last_date = False

			# Kolla om sista elementet i listan är ett datum eller ej. Detta är avgörande för hur listan delas upp. Om man inte inför argumentet att längden måste vara 8 (eg. YYYYMMDD = len 8) samt att det ska börja med 20 finns risk att den plockar datum som inte är datum. Nu minimerar vi risken för detta
			try:
				if len(date_string) == 8 and date_string.startswith('20') == True:
					datetime.datetime.strptime(date_string, format)
					last_date = True
			except ValueError:
				pass


			date_list = []
			# Detta finns för att undersöka vart det första datumet ligger i listan. Detta har betydelse alldeles strax
			try:
				third_ele = lista[2]
				fourth_ele = lista[3]
				fifth_ele = lista[4]
			except:
				pass

			# Detta är bara en indikator för att ta reda på vilken plats datumet ligger på
			third = False
			fourth = False
			fifth = False

			# Kolla var det första datumet i listan ligger i listan. Detta är också avgörande för hur listan ska delas up

			try:
				if len(third_ele) == 8 and third_ele.startswith('20') == True:
					datetime.datetime.strptime(third_ele, format)
					third = True
			except ValueError:
				try:
					if len(fourth_ele) == 8 and fourth_ele.startswith('20') == True:
						datetime.datetime.strptime(fourth_ele, format)
						fourth = True
				except ValueError:
					try:
						if len(fifth_ele) == 8 and fifth_ele.startsiwth('20') == True:
							datetime.datetime.strptime(fifth_ele, format)
							fifth = True
					except ValueError:
						pass

			final_ver = []
			# Detta är hur listan kommer vara formaterad i olika fall då det sista elementet i listan har ett datum
			if last_date == True:
				# Om datumet är på tredje plats (vilket det är oftast) så gör vi följande
				if third == True:
					ver_descr = lista[3:-1]
					ver_descr_string = ''
					for x in ver_descr:
						ver_descr_string += " "+x
					ver_descr_string = ver_descr_string[1:]
					series = lista[0]
					identification = lista[1]
					date_1 = lista[2]
					date_2 = lista[-1]
					# Skapa listan
					final_ver.append(series)
					final_ver.append(identification)
					final_ver.append(date_1)
					final_ver.append(ver_descr_string)
					final_ver.append(date_2)
				# Om datumet är på fjärde plats (sällan) så gör vi följande
				elif fourth == True:
					ver_descr = lista[4:-1]
					ver_descr_string = ''
					for x in ver_descr:
						ver_descr_string += " "+x
					ver_descr_string = ver_descr_string[1:]
					series = lista[0]
					identification = lista[1] + " " + lista[2]
					date_1 = lista[3]
					date_2 = lista[-1]
					# Skapa listan
					final_ver.append(series)
					final_ver.append(identification)
					final_ver.append(date_1)
					final_ver.append(ver_descr_string)
					final_ver.append(date_2)

				# Om datumet är på femte plats (extremt sällan) så gör vi följande
				elif fifth == True:
					ver_descr = lista[5:-1]
					ver_descr_string = ''
					for x in ver_descr:
						ver_descr_string += " "+x
					ver_descr_string = ver_descr_string[1:]
					series = lista[0]
					identification = lista[1] + " " + lista[2] + " " + lista[3]
					date_1 = lista[4]
					date_2 = lista[-1]
					# Skapa listan
					final_ver.append(series)
					final_ver.append(identification)
					final_ver.append(date_1)
					final_ver.append(ver_descr_string)
					final_ver.append(date_2)

			# Detta är hur listan kommer vara formaterad i olika fall då det sista elementet i listan INTE är ett datum
			if last_date == False:

				# Om datumet är på tredje plats (vilket det är oftast) så gör vi följande
				if third == True:
					ver_descr = lista[3:]
					ver_descr_string = ''
					for x in ver_descr:
						ver_descr_string += " "+x
					ver_descr_string = ver_descr_string[1:]
					series = lista[0]
					identification = lista[1]
					date_1 = lista[2]
					# Skapa listan
					final_ver.append(series)
					final_ver.append(identification)
					final_ver.append(date_1)
					final_ver.append(ver_descr_string)

				# Om datumet är på fjärde plats (sällan) så gör vi följande 
				elif fourth == True:
					ver_descr = lista[4:]
					ver_descr_string = ''
					for x in ver_descr:
						ver_descr_string += " "+x
					ver_descr_string = ver_descr_string[1:]
					series = lista[0]
					identification = lista[1] + " " + lista[2]
					date_1 = lista[3]
					# Skapa listan
					final_ver.append(series)
					final_ver.append(identification)
					final_ver.append(date_1)
					final_ver.append(ver_descr_string)

				# Om datumet är på femte plats (extremt sällan) så gör vi följande
				elif fifth == True:
					ver_descr = lista[5:]
					ver_descr_string = ''
					for x in ver_descr:
						ver_descr_string += " "+x
					ver_descr_string = ver_descr_string[1:]
					series = lista[0]
					identification = lista[1] + " " + lista[2] + " " + lista[3]
					date_1 = lista[4]
					# Skapa listan
					final_ver.append(series)
					final_ver.append(identification)
					final_ver.append(date_1)
					final_ver.append(ver_descr_string)

			# Korrigering för att vi inte vill att det fjärde elementet i listan ska vara ett datum!
			if len(final_ver[3]) == 8 and final_ver[3].startswith('20') == True:
				try:
					datetime.datetime.strptime(final_ver[3], format)
					final_ver.insert(3,'')
				except ValueError:
					pass

			ver = final_ver
			# Detta steg är väldigt nödvändigt så att rätt variabel blir returnad. Annars blir det massa "None"s. Nu accessar vi rätt variabel i TRANS()
			if ver != None:
				global globver
				globver = ver

		return globver

	def TRANS(self):
		line = self.line
		outputWriter1 = self.outputWriter1
		# Här hämtar vi in den globala variabeln
		ver = m.VER()

		# Hittar alla Transaktioner och sammanställer dem i en .csv
		if "TRANS " in line:

			trans = line
			# Måste ta bort komman (,) då de riskerar att export till CSV förstörs. Ö,Ä,Å och ö,ä,å kan ej läsas för ngn anledning. Dessa måste konverteras
			trans = trans.replace(",","")
			trans = trans.replace('\\',"")

			# Vi måste se om det är en särskild sorts transaktion. Det finns t.ex. #RTRANS och "#BTRANS". När fallet är #TRANS endast är det N som gäller
			trans_type = re.findall("#(.*?)TRANS", line)[0]
			if trans_type == '':
				trans_type = 'N'

			# Hittar den data som är innanför {} och tar bort den då den är irrelevant och gör listan svårhanterad annars. Här står ibland Dim tror jag.
			trans2 = re.findall("{(.*?)}", trans)
			str1 = ''
			for x in trans2:
			   str1 += x
			within_brackets = str1
			str1 = '{'+str1+'}'
			trans = trans.split(str1)
			
			# Denna finns till så att elementen i fixed_trans-listan läggs till i korrekt ordning. Note: Det som heter new_string är egentligen en lista 
			fixed_trans = []
			check = False
			for x in trans:
				new_string = x.split()

			# Lägger till belopp
				if check == True:
					fixed_trans.append(new_string[0])

			# Lägger till kontonr	
				if check == False:
					fixed_trans.append(new_string[-1])
					check = True

			# Sammanställer en lista som är korrekt formaterad. Denna konverteras till csv
			final_lista = []
			# Ibland finns ingen beskrivning av verifikationen. Då hamnar datum på fel ställe och allt blir förskjutet
			if len(ver) == 4:
				ver.append('')
			try:
				for x in ver:
					final_lista.append(x)

				for x in fixed_trans:
					final_lista.append(x) 
			except:
				print("NÅGOT GICK SNETT")
				pass

			#Lista med {xxx}
			brackets_list = final_lista
			brackets_list.append(within_brackets)
			#Lägger till trans_typ
			brackets_list.append(trans_type)

			# Måste även lägga till belopp_2, som ska vara reversat belopp om kontot startar med 3, 4, 5, 6, 7, 8 eller (9)
			# brackets_list[6] är alltså beloppet
			belopp_2 = brackets_list[6]
			# Här reversas beloppet om kontot är kontogrupp 3-9, dvs intäkter eller kostnader.
			# brackets_list[5] är alltså kontot
			konto = brackets_list[5]
			if konto.startswith('3') or konto.startswith('4') or konto.startswith('5') or konto.startswith('6') or konto.startswith('7') or konto.startswith('8') or konto.startswith('9'):
				# belopp_2 = '-' + belopp_2
				belopp_2 = belopp_2
			
			brackets_list.append(belopp_2)

			# Skriver över listan till outputfilen
			outputWriter1.writerow(brackets_list)

# Vi måste införa en global variabel som hjälper oss komma åt ver i funktionen VER när vi behöver
global globver
globver = ""
counter = 0

print('beginning to loop through input files...')
print('---------------------------------------------------------------')

# Loopar genom alla rader
# Måste loopa igenom rätt directory. Om vi inte sätter os.chdir(directory) så loopar den igenom directory2.
os.chdir(directory)
for file in os.listdir(directory):
	#print(file)
	if file.endswith(".se") or file.endswith(".SE") or file.endswith(".SI") or file.endswith(".si") or file.endswith(".SIE"):
		print(file)
		# Öppnar filen
		with open(file, "r", encoding='cp437') as f:
			content = f.readlines()
			for line in content:
				m = Master(line, outputWriter_UB, outputWriter_IB, outputWriter_RES, outputWriter_PSALDO, outputWriter_KONTO, outputWriter1)
				# Detta finns endast till för om man behöver experimentera med koden. Annars loopar den igenom en miljon rader. Nu bara 5750 (elr så många man vill). Dvs bara uncommenta nedan för att stoppa hundra tusentals rader från att skrivas varje gång.
				#if counter == 5750:
				#	break
				# counter += 1
				m.UB()
				m.IB()
				m.RES()
				m.PSALDO()
				m.KONTO()
				m.VER()
				m.TRANS()

#stäng alla filer
outputFil1.close()
outputFil_KONTO.close()
outputFil_UB.close()
outputFil_IB.close()
outputFil_RES.close()
outputFil_UB.close()

print('---------------------------------------------------------------')
print('closing files - output ready')