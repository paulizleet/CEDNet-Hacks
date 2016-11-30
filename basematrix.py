from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import coordinate_to_tuple
from openpyxl.formula import Tokenizer
from decimal import *
import time


productList = []
writeList = []
target = "BASE "
path = "C:\\PaulScripts\\Pricing Matrices\\"

TOP_CUSTOMER = .9
REG_CUSTOMER = .85


def writeMatrix(fn, productList, margin, margins):
	"""whut"""
	#print(productList[0])
	#print(writing)

	f = open(path + fn, 'w')
	i = 0
	j=0
	for line in productList:
		if line[2] == "MILW " or line[2] == "KLEIN" or line[2] == "SOWIR":
			i+=1
			continue
		j+=1
		wline = ""

		wline = ""
		wline+=line[0]
		wline+=line[1]
		wline+=line[2]
		wline+=line[3]
		wline+=formatPrice(line[4], line[2], line[5], margins)
		wline+=line[5]
		wline+=line[6]
		wline+=line[7]
		wline+=line[8]
		wline+=line[9]
		wline+=line[10]
		wline+=line[11]
		#print(wline)
		wline+="\n"
		f.write(wline)		
		#quit()
	
	print(str(i) + " Items skipped")
	print(str(j) + " Items given prices")
	f.close()

def formatPrice(price, mfr, uom, margins):
	multiplier = 0
	margin = 0
	zeroes = "000000000000000"
	
	if price is None:
		return zeroes
	
	print(mfr)
	price = round((float(price) * get_multiplier(uom) ) / get_margin(mfr, margins),2)
		
	price=int(price*100)
	
	price = str(price)
	
	price = zeroes[:-price.__len__()]+price
	
	return price

def readLPF(fn):
	"""whut"""
	
	f = open(fn, 'r')
	pl = []
	i = 0
	for line in f:
		pl.append(parseLPFLine(line))
		i+=1

	f.close()
	return pl
	
def readMatrix(fn):
	"""whut"""
	
	f = open(path + fn, 'r')
	pl = []
	i = 0
	for line in f:
		pl.append(parseLine(line))
		i+=1

	f.close()
	return pl
	
def parseLPFLine(line):
		"""whut
		
		"""
		pl = []
		
		line=line.split("|")
		
		i = 0
		#for each in line:
		#	print(str(i) + ". " + each)
		#	i+=1
		if "PS" in line[1]:
			print(line[1:3])
	
		pl.append(target) #Target matrix
		pl.append("5")#Level
		pl.append(line[1] + "     "[:-line[1].__len__()])#Mfr
		pl.append(line[2] + "                      "[:-line[2].__len__()])#Name
		pl.append(line[10])#Price  "000000000000000"[:-price.__len__()]+price
		pl.append(line[12])#UOM
		pl.append("01012020")#Exp
		pl.append("   ")#Pad 3
		pl.append("00000000000"[:-line[3].__len__()]+line[3])	#UPC
		pl.append("      ") #pad 6
		pl.append("000000000000000")#Zeroes
		pl.append(line[1] + "     "[:-line[1].__len__()]) #matrix


		return pl

def parseLine(line):
	"""whut
	
	"""
	pl = []
	pl.append(line[:11])
	pl.append(line[11:33])
	pl.append(float(line[33:48]) /100)
	pl.append(line[48:97])
	return pl
	
def compare(pl, cm):
	
	print(pl.__len__())
	for cmeach in cm:
		popUs = []
		print(cmeach[1])
		for i in range(0, pl.__len__()):
			if cmeach[1] == pl[i][3]:
				#print("will pop: " + str(i))
				popUs.append(i)
		
		
		for each in popUs:
			pl.pop(each)
	
		
	print(pl.__len__())
	
	print(pl)
	return pl

def get_margin(mfr, margins):
	
	for each in margins:
		if each[0] == mfr.strip():
		#	input(each)
			return float(each[1])
	
	return float(margins[0][1])

def get_multiplier(uom):
			
	if uom == "E":	#each
		return 1
	if uom == "C":	#per-hundred
		return 100
	if uom == "M": #per-thousand
		return 1000

def readMargins():
	f = open("C:\PaulScripts\Pricing Matrices\mfr_margins.txt")
	lns = []
	
	for line in f:
		lns.append(line.split())
		
	return lns
	
def run():
	#read all products in
	try:
		productList = readLPF("C:\Invsys\Algorithm\SPKPRDDT.lsq")
	except FileNotFoundError:
		print("CEDNet data files not found.  \nPlease run the Data Files export function in CEDNet")
		print("It can be found at Maintainance > Product > Data Files")
		return
		
	margins = readMargins()
	i=0

	#print("got list")
	#currentMatrix = readMatrix("base_export.lsq")
	
	#currentMatrix = compare(productList, currentMatrix)
		
		
	writeMatrix("Import Base.lsq", productList, REG_CUSTOMER, margins)
	
	print(str(productList.__len__()) + " Items in total.")
	
if __name__ == "__main__":
	run()
	
	print("Base matrix written to \"C:\\PaulScripts\\Pricing Matrices\\Import Base.lsq\"")
	
else:
	print("imported basematrix.py")