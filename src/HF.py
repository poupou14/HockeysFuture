#!/usr/bin/python 
import string, sys
from HFParser import HFParser
#from HFWriter import HFWriter

def main():
	myHF = HFParser()
#	myHFWriter = HFWriter("../INPUT/HF_Players.xls")
	sourceFile_l = "/home/lili/Developpement/HockeysFuture/INPUT/Draft-2014-LNHV2-02-09-14.xls"
	targetFile_l = "/home/lili/Developpement/HockeysFuture/OUTPUT/Draft-2014-LNHV2-02-09-14-out.xls"
	myHF.readHF(sourceFile_l, targetFile_l)

main()
