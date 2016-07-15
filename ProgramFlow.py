#coding: utf-8
import shlex
import string 

class ProgramFlow(object):
	
	def __init__(self):

		self.cmdLine=""					#cmd command for hansoftexport.exe
		self.fileToBeTransfered=""		#Name of the file to transfer (to ericoll preferably)
		self.finalFileName=""			#Desired name for the transfered file
		self.finalFolder=""				#Desired location (ericoll preferably)
		self.latestExport=""			#folder path + file name for latest hansoft export
		self.hansoftDatabase=""			#name of the hansoft database to connect to
		self.hansoftServer=""			#name of the hansoft server to connect to
		self.hansoftPort=""			#port number to connect to 
		self.sdkUser=""
		self.sdkPass=""
		self.hansoftProjectName=""
		self.hansoftProjectView=""
		self.reportName=""
		self.reportAuthor=""
		self.query=""
		self.configLine=""
		self.macroName=""
		self.hansoftExtractFolder=""
		self.alteredExportsFolder=""
		self.exhaustFolder=""
		self.macroWorkbook=""
		self.macroModule=""
		self.configFile="C:\Users\ejoanyk\Desktop\ExHaUST\Config.txt" #just here for now to be able to test the methods 
		self.adminFile=""
		self.fullConfigFile=[]
		
		self.filesToKeep=[]
		
		self.currentRow=0
		self.lastRow=0
		self.nrConAttemptsAllowed=0

		self.runMacro=False
		self.reportRequested=False
		self.queryRequested=False

	def SetAdminVars(self):
		return 

	def ReadLineConfigFile(self):
		
		openConfigFile=open(configFile, 'r')	 		#Opens config.txt in read-only mode
		
		self.fullConfigFile = openConfigFile.readlines() 	#Creates a list made up of the lines in config files as strings, 
									#e.g. one full line at each index, and sets this list to the fullConfigFile attribute
		return

	def LineOK(self)
		return	

	def BuildCMDLine(self):
		return

	def UpdateVars(self):
		
		self.ReadLineConfigFile(self.configFile) 

		roughConfigLine= shlex.split(self.fullConfigFile[self.currentRow], posix=False)


		for j in roughConfigLine:
		    roughConfigLine[roughConfigLine.index(j)] = j.strip("\"")

		self.hansoftServer=roughConfigLine[0]
		self.hansoftPort=roughConfigLine[1]
		self.hansoftDatabase=roughConfigLine[2]
		self.sdkUser=roughConfigLine[3]
		self.sdkPass=roughConfigLine[4]
		self.hansoftProjectName=roughConfigLine[5]
		roughConfigLine[6]=roughConfigLine[6].strip("-")
		self.hansoftProjectView=roughConfigLine[6]

		def HandleMacroFields(indexOfRunMacro):
			if roughConfigLine[indexOfRunMacro]=="-m":
				self.runMacro=True
				self.macroWorkbook=roughConfigLine[indexOfRunMacro+1]
				self.macroModule=roughConfigLine[indexOfRunMacro+2]
				self.macroName=roughConfigLine[indexOfRunMacro+3]
				if len(roughConfigLine)==roughConfigLine[indexOfRunMacro+4]:   #checks if the fileToBeTransfered field is filled in
					self.fileToBeTransfered=roughConfigLine[indexOfRunMacro+4]
			return

		if roughConfigLine[7]=="-r":
			self.reportRequested=True
			self.reportName=roughConfigLine[8]
			self.reportAuthor=roughConfigLine[9]
			self.finalFolder=roughConfigLine[10]
			self.finalFileName=roughConfigLine[11]
			indexOfRunMacro=12	
			HandleMacroFields(12)

		elif roughConfigLine[7]=="-q":
			self.queryRequested=True
			self.query=roughConfigLine[8]
			self.finalFolder=roughConfigLine[9]
			self.finalFileName=roughConfigLine[10]
			indexOfRunMacro=11
			HandleMacroFields(11)


		print self.hansoftServer
		print self.hansoftPort
		print self.hansoftDatabase
		print self.sdkUser
		print self.sdkPass
		print self.hansoftProjectView
		print self.hansoftProjectName
		print self.runMacro
		print self.macroWorkbook
		print self.macroModule
		print self.macroName
		print self.fileToBeTransfered
		print self.reportRequested
		print self.reportName
		print self.reportAuthor
		print self.queryRequested
		print self.query
		print self.finalFolder
		print self.finalFileName
		
		return

	def ControlFlow(self):
		return

	def ShouldFileBeTransfered(self):
		return 


g=ProgramFlow()
g.UpdateVars(0)
