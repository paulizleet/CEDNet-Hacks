import os
import string




def main():
	
	csdat = get_customers()
	cust= []
	for each in csdat:
		cust.append(each[0])
		print(each[0])
	f=open("C:\\Users\\pgallagherjr\\Desktop\\spkout.txt", 'w')

	mfr = get_mfr()
	spk = run_speaks(cust,"C:\\Users\\pgallagherjr\\Desktop\\check.txt", mfr)
	
	chop(spk, cust)
	

def get_customers():


	cs = open("C:\\Invsys\\Algorithm\\SPKCUSDT.lsq")
		
	customers = []
	cline = []
	for line in cs:
		cline = []
		l = line.split("|")
		cline.append(l[3].strip()) #Name
		cline.append((l[5]+ " " + l[6]).strip()) #Addr
		cline.append(l[9]) #city
		cline.append(l[7]) #zip
		cline.append(l[8]) #State
		
		customers.append(cline)
		
	print(customers[0])

		
	return customers

def get_mfr():

	mf = open("C:\\Invsys\\Algorithm\\SPKMFRDT.lsq")
		
	mfrs = []
	for line in mf:
		ln = line.split("|")
		print(ln[1].strip())
		mfrs.append(ln[1].strip())	
		
		
	return mfrs
		
def run_speaks(customers, fp, mfr):
	f = open(fp)
	lines = []

	for line in f:
	
		l = line.split("\t")
		
		if l[0][:5].strip() in mfr:        
			lines.append(l)
			continue
		try:

			if l[1].strip() in customers:
				lines.append(l)
			
		except:
			continue                  
		
	for each in lines:
		dd = []
		for each2 in each:
			dd.append(each2.replace('\n', ''))

	asdf= []
	
	cs = ""
	
	for each in splits:
		ap = each
		try:
			if each[1].strip() in customers:
				cs = each[1]
				continue
		except:
			pass
	l = []	
	for each in asdf:

		if each.__len__() <= 3:
			continue
		
		ln = ""
		for p in each:
			ln=ln+p + "\t\t\t"
			
		ln=ln+"\n"

	return final

def chop(spk, cust):
	path = "C:\\Users\\pgallagherjr\\Desktop\\out\\"
	
	ind = 0
	print(spk[0])
	num_orders = 0
	for each in cust:
		#check for customer directory	
		e = each
		try:
			os.mkdir(path+each)
		except FileExistsError:
			print("dir already exists")
		except OSError:
			e = ""
			for char in each:
				if char not in string.punctuation:
					e=e+char
			try:
				os.mkdir(path+e)
			except:
				pass

		ind = 0
		cslines = []
		for j in range(0, spk.__len__()):
			if spk[j-ind][0] != each:
				break
			cslines.append(spk.pop(j-ind))
			ind+=1
		
		while cslines.__len__() > 0:
		
			invlines = []
			
			index = 0
			csinv = cslines[0][-4]
			num_orders += 1
			for i in range(0, cslines.__len__()):
				if cslines[i-index][-4] == csinv:
					invlines.append(cslines.pop(i-index))
					index += 1
			
		#	print(invlines[0])
			f = open(path+e+"\\"+csinv + ".txt", "w")
		#	print(path+e+"\\"+csinv)
			
			for ln in invlines:
				
				outline = ""
				for pce in ln:
					outline = outline + "\t" + pce
				
				outline = outline +"\n"
				f.write(outline)
				
			f.close()
	
	print(num_orders)
			

	
main()