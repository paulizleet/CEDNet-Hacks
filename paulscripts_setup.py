# -*- coding: utf-8 -*-
"""
Created on Tue Sep 13 10:31:55 2016

@author: Paul Gallagher

Setup script for my CEDNet Utilities
"""

import os
import pip
from shutil import copytree
import logging


    
from openpyxl import Workbook

import cycle
import nobins


path = "C:\\"

def make_dir(dir):
    print(path + dir)
    try:
        os.mkdir(path + dir)
    except FileExistsError:
        print("Directory {d} already exists.".format(d=dir))

        
def _logpath(path, names):
    logging.info('Working in %s' % path)
    return []   # nothing will be ignored       

if __name__ == "__main__":
    
    if os.getcwd()[0] != 'C':
        print("I noticed you are running this from a flash drive. \n\n"\
        "This program is fully portable, but will run very slowly on a flash drive."\
        "\nConsider installing this utility to your hard disk to make it faster.")
        while True:
            inp=input("Would you like this program to copy itself to your hard drive? Y/N\n")
            if inp.upper() == 'Y':
            
                print("This is going to take a while.  Just don't close this program before it's finished.")
                copytree(os.getcwd()[:3], "C:\\Python", ignore=_logpath)
                print("\nFinished Copying files.")
                break
            elif inp.upper()=='N':
                print("Be prepared to wait a while.\n\n")
                break
                
                
                
    #Installing the required libraries
    try:
        pip.main(['install', 'openpyxl'])
    except:
        print("pip failed to install openpyxl.  Halting.")
        quit() 
    try:
        pip.main(['install', 'pypiwin32'])
    except:
        print("pip failed to install pywin32.  Halting.")
        quit()  
    
    #Making all of the needed directories
    make_dir("PaulScripts")
    make_dir("PaulScripts\\Inventory Checking")
    make_dir("PaulScripts\\Pricing Matrices")
    make_dir("PaulScripts\\Shelving Addresses")
    make_dir("PaulScripts\\Solar Reporting")
    make_dir("PaulScripts\\Solar Reporting\\Weekly")
    make_dir("PaulScripts\\Solar Reporting\\Quarterly")
    make_dir("PaulScripts\\Solar Reporting\\Monthly")
    make_dir("PaulScripts\\Solar Reporting\\Monthly\\SMA")
    make_dir("PaulScripts\\Solar Reporting\\Monthly\\LG")
    make_dir("PaulScripts\\Speaks Exports")
    make_dir("PaulScripts\\Wire Matrix")
    make_dir("PaulScripts\\Wire Matrix\\wirebooks")
	make_dir("PaulScripts\\sys\\configs")
	make_dir("PaulScripts\\sys\\logs")
	
	f=open("C:\\PaulScripts\\configs\\customer_info.txt",'w')
	
	f.write("2		Account Number\n".strip())
	f.write("3		Customer Name\n".strip())
	f.write("5		Address Line 1\n".strip())
	f.write("6		Address Line 2\n".strip())
	f.write("9		Customer city\n".strip())
	f.write("8		Customer zip\n".strip())
	f.write("7		Customer state\n".strip())
	f.close()
    wb = Workbook()
    
    
    #Creating three empty workbooks before initializing them
    wb.save(path + "PaulScripts\\Inventory Checking\\cycle.xlsx")
    wb.save(path + "PaulScripts\\Inventory Checking\\cycle_out.xlsx")
    wb.save(path + "PaulScripts\\This Week's Stock Status.xlsx")

    #Initializing spreadsheets which need it
    nobins.run(skip_prompt=True)
    print("Nobins ran for the first time.")
    cycle.run(skip_prompt=True)
    print("cycle ran for the first time")
    
    f = open("C:\\PaulScripts\\.init.txt", 'w')
    f.write("Scripts setup OK")
    f.close()