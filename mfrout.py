f = open("C:\INVSYS\Algorithm\SPKMFRDT.lsq")
of = open("C:\PaulScripts\mfrout.txt", 'w')


mfrs = []

for line in f:
	of.write(line.split("|")[1].strip() + "|.85\n")
	
of.close()
	

	