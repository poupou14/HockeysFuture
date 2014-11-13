#!/usr/bin/python 
from HTMLParser import HTMLParser
import HFDataFormat
import os,string, sys
import urllib
import copy
import chardet
#### SPECIFIC IMPORT #####
sys.path.append("/home/lili/Developpement/HockeysFuture/Import/xlrd-0.7.1")
sys.path.append("/home/lili/Developpement/HockeysFuture/Import/xlwt-0.7.2")
sys.path.append("/home/lili/Developpement/HockeysFuture/Import/pyexcelerator-0.6.4.1")

from pyExcelerator import *
import xlwt
from xlrd import open_workbook
from xlwt import Workbook,easyxf,Formula,Style
#from lxml import etree
import xlrd

currentPlayer_g = ''

def onlyascii(char):
    if ord(char) <= 0 or ord(char) > 127: 
	return ''
    else: 
	return char

def isnumber(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

class HFParser():

	def __init__(self): 
		self.__status = False
		self.__counter = 0
		self.__rootAdd = "http://www.hockeysfuture.com/prospects/"

	def readHF(self, inputFile_p, outpuFile_p):
		# On va lire le fichier Input
		global currentPlayer_g
		self.getNames(inputFile_p)
		myHFParserSummary = HFParserSummary()
		myGoogleForHFParser = GoogleForHFParser()
		len_l = len(HFDataFormat.listPlayer)
		for playerName_l in HFDataFormat.listPlayer.keys() :
			self.__counter+=1
			print "%d" % self.__counter 
			print "%s" % playerName_l 
			myGoogleForHFParser.reset()
			hfPlayerUrl_l = myGoogleForHFParser.getHfPlayerUrl(playerName_l)
			try :
				myHFParserSummary.reset()
				print "try to read %s" % hfPlayerUrl_l
				url = urllib.urlopen(hfPlayerUrl_l)
				myHFParserSummary.html = url.read()
				url.close()
				currentPlayer_g = HFDataFormat.listPlayer[playerName_l]['Name']
				myHFParserSummary.html = filter(onlyascii, myHFParserSummary.html)
				myHFParserSummary.feed(myHFParserSummary.html)	
				HFDataFormat.listPlayer[playerName_l]['html'] = hfPlayerUrl_l
			except IOError:
				print "%s not readable" % hfPlayerUrl_l 
			
		self.__status = True
		self.writeOuput(outpuFile_p)	
		print HFDataFormat.listPlayer



	def getNames(self, inputFileName) :

		try:
			inputBook_l = open_workbook(inputFileName, encoding_override="cp1252")
		except:            
			print "Error: Impossible to open file %s" % inputFileName                       
			sys.exit()
    
		#read the sheet
		try:
			worksheet_l = inputBook_l.sheet_by_index(0)
		except:
			print  "Error: no sheet 0 found in %s" % xlsfile
			sys.exit()
		
		numRows_l = worksheet_l.nrows - 1
		numCells_l = worksheet_l.ncols - 1
		currRow_l = -1
		while currRow_l < numRows_l:
			currRow_l += 1
			row_l = worksheet_l.row(currRow_l)
			# Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
			name_l = worksheet_l.cell_value(currRow_l, 0)
			print name_l
			HFDataFormat.currentPlayer = dict(HFDataFormat.emptyPlayer)
			HFDataFormat.currentPlayer['Name'] = name_l
			HFDataFormat.listPlayer[name_l] = HFDataFormat.currentPlayer
			teamLNHV2_l = worksheet_l.cell_value(currRow_l, 1)
			HFDataFormat.listPlayer[name_l]['TeamLNHV2'] = teamLNHV2_l
			print "Team : %s" % HFDataFormat.listPlayer[name_l]['TeamLNHV2'] 


	def writeOuput(self, outputFile_p):
#		Book to read xls file (output of main_CSVtoXLS)
    		workbook1_l = Workbook()
    		newPlayerSheet_l = workbook1_l.add_sheet("Prospects", cell_overwrite_ok=True)

		newPlayerIndex_l = 1
		title = False

		for playerName_l in HFDataFormat.listPlayer.keys() :
			player_l = HFDataFormat.listPlayer[playerName_l]
			if not(title) :
				self.addClmnTitle(newPlayerSheet_l, player_l)
				title = True
				self.addPlayer(newPlayerSheet_l, player_l, newPlayerIndex_l) 
				newPlayerIndex_l += 1
			else:
				self.addPlayer(newPlayerSheet_l, player_l, newPlayerIndex_l) 
				newPlayerIndex_l += 1
		workbook1_l.save(outputFile_p)

	def addClmnTitle(self, sheet, player) :
		indexClmn_l = 0
		for key in player.keys() :
			value_l = filter(onlyascii, key)
			sheet.write(0, indexClmn_l, value_l)
			indexClmn_l += 1
 	
	def addPlayer(self, sheet, player, index):
		indexClmn_l = 0
		style = easyxf('font: underline single')
#		for key in player.keys() :
		for key in HFDataFormat.keyListPlayer :
			if isnumber(player[key]) :
				try :
					if player[key].find(".") != -1 :
						value_l = float(player[key])
					else:
						value_l = player[key]
					sheet.write(index, indexClmn_l, value_l)
				except :
					value_l = player[key]
					sheet.write(index, indexClmn_l, value_l)
					
			else :
				if key == "html" :
					link = 'HYPERLINK("%s";"%s")' % (player[key], player['Name'])
					value_l = Formula(link)
					sheet.write(index, indexClmn_l, value_l, style)
				else :
					value_l = filter(onlyascii, player[key])
					sheet.write(index, indexClmn_l, value_l)
			indexClmn_l += 1
				
			
class GoogleForHFParser(HTMLParser): 

	def __init__(self): 
		self.__playerUrlFound = False
		self.__playerUrl = ""

	def getHfPlayerUrl(self, playerName_p) :
		self.__playerUrlFound = False
		self.__playerUrl = ""
		query_l = ''.join((playerName_p, " hockeysfuture"))
		filename = 'http://www.google.com/search?' + urllib.urlencode({'q': query_l })
		cmd = os.popen("lynx -dump %s | grep www.hockeysfuture.com\/prospects\/" % filename)
		output = cmd.read()
		lineResult_l = output.splitlines()
		cmd.close()
		for line_l in lineResult_l :
			wordsResult_l = line_l.split(" ")
			for word_l in wordsResult_l :
				if word_l.find("www.hockeysfuture.com") != -1 :
					self.__playerUrlFound = True
					self.__playerUrl = word_l
					break
			if self.__playerUrlFound : 
				break
		if self.__playerUrl.find("http") == -1 :
			self.__playerUrl = ''.join(("http://", self.__playerUrl))
		return self.__playerUrl	
		

class HFParserSummary(HTMLParser): 

	def __init__(self): 
		self.__readOK = False 
		self.__getTitle = False 
		self.__betweenTag = False
		self.__title = False 
		self.__playerUrl = False
		self.__nextScore = 0
		self.__nextSuccess = 0
		self.__nextTalent = 0
		self.__nextFuture = 0
		self.__nextPosition = 0

	def setReadNext(self, val) :
		self.__readNext = val

	def readNext(self) :
		return self.__readNext	

	def getNextPage(self) :
		return self.__nextPage	

	def handle_starttag(self, tag, attrs):
		if tag == "title" :
			self.__getTitle = True
		elif tag == "span" and len(attrs) == 1 :
			if attrs[0][1] == "numerical" : 
				self.__nextScore = 2
			else :
				self.__nextScore = 0
		elif tag == "li" :
			if self.__nextScore == -1 and self.__nextSuccess != -1:	
				self.__nextSuccess = 2
			else :
				self.__nextSuccess = -1
		elif tag == "p" :
			if self.__nextTalent == 1:
				self.__nextTalent = 2
			elif self.__nextFuture == 1:
				self.__nextFuture = 2
			elif self.__nextPosition == 1:
				self.__nextPosition = 2
				
		self.__betweenTag = True

	def handle_data(self, data):
		if self.__nextTalent == 2 :
			HFDataFormat.listPlayer[currentPlayer_g]['talent'] = data
			print "talent : %s" % data
			self.__nextTalent = -1
		elif self.__nextFuture == 2 :
			HFDataFormat.listPlayer[currentPlayer_g]['future'] = data
			print "future : %s" % data
			self.__nextFuture = -1
		elif self.__getTitle  :
			if data.find("Page not found") != -1 :
				raise IOError 
			else :
				self.__getTitle = False
		elif self.__nextScore == 2 :
			HFDataFormat.listPlayer[currentPlayer_g]['score'] = data
			self.__nextScore = -1
		elif self.__nextSuccess == 2 :
			HFDataFormat.listPlayer[currentPlayer_g]['success'] = data
			self.__nextSuccess = -1
		elif self.__nextPosition == 2 :
			HFDataFormat.listPlayer[currentPlayer_g]['Pos'] = data
			print "Position : %s" % data
			self.__nextPosition = -1
			
		elif self.__betweenTag :
			if data.find("Prospect Talent Score") == 0 :
				self.__nextScore = 1
			elif data.find("Probability of Success") == 0 :
				self.__nextSuccess = 1
			elif data.find("Talent Analysis") == 0 :
				self.__nextTalent = 1
			elif data.find("Future") == 0 :
				self.__nextFuture = 1
			elif data.find("Position") == 0 :
				self.__nextPosition = 1
		
			

	def handle_endtag(self, tag):
		self.__betweenTag = False



class HFWriter() :

	def __init__(self):
		self.__workingSheet = ""


	def fileUpdate(self, inputFileName, outputFileName):
    		workbook1_l = Workbook()
    		worksheet1_l = workbook1_l.add_sheet("Sheet1", cell_overwrite_ok=True)


#		 Remplissage des champs pour chaque article
		indexRaw_l = 0
		for article_l in listeArticle :
#			print article_l
			indexClmn_l = 0
			if indexRaw_l == 0 :
#		 Remplissage de la premiere ligne
				for key_l in keyListeArticle:
					worksheet1_l.write(0, indexClmn_l, article_l[key_l])
					indexClmn_l+=1
			else:
				for key_l in keyListeArticle:
					if isinstance(article_l[key_l], str):
						article_l[key_l] = cleanString(article_l[key_l])
					worksheet1_l.write(indexRaw_l, indexClmn_l, article_l[key_l])
					indexClmn_l+=1
#				print article_l
			indexRaw_l += 1
    		workbook1_l.save(name)



def open_excel_sheet():
    """ Opens a reference to an Excel WorkBook and Worksheet objects """
    workbook = Workbook()
    worksheet = workbook.add_sheet("Sheet 1")
    return workbook, worksheet

def write_excel_header(worksheet, title_cols):
    """ Write the header line into the worksheet """
    cno = 0
    for title_col in title_cols:
        worksheet.write(0, cno, title_col)
        cno = cno + 1
    return

def write_excel_row(worksheet, rowNumber, columnNumber):
    """ Write a non-header row into the worksheet """
    cno = 0
    for column in columns:
        worksheet.write(lno, cno, column)
        cno = cno + 1
    return

def save_excel_sheet(workbook, output_file_name):
    """ Saves the in-memory WorkBook object into the specified file """
    workbook.save(output_file_name)
    return

