#coding: utf-8
import shlex
import string 

class ProgramFlow(object):
	
	def __init__(self):

		self.cmdLine=""				#cmd command for hansoftexport.exe
		self.fileToBeTransfered=""		#Name of the file to transfer (to ericoll preferably)
		self.finalFileName=""			#Desired name for the transfered file
		self.finalFolder=""			#Desired location (ericoll preferably)
		self.latestExport=""			#Folder path + file name for latest hansoft export
		self.hansoftDatabase=""			#Name of the hansoft database to connect to
		self.hansoftServer=""			#Name of the hansoft server to connect to
		self.hansoftPort=""			#Port number to connect to 
		self.sdkUser=""				#Hansoft SDK username
		self.sdkPass=""				#Hansoft SDK password
		self.hansoftProjectName=""		#Name of the project to look for
		self.hansoftProjectView=""		#Backlog view, sprint-view etc
		self.reportName=""			#Name of the report to fetch
		self.reportAuthor=""			#Author of the report
		self.query=""				#Hansoft query
		self.configLine=""			#Will this be used? 
		self.macroName=""			#Name of the macro to be run (if there is a macro that will be run)
		self.hansoftExtractFolder=""		#Folder where HansoftExtract.exe lives
		self.alteredExportsFolder=""		#Altered Exports folder
		self.exhaustFolder=""			#Folder where ExHaUST.exe, Config.txt, README.txt, etc live
		self.macroWorkbook=""			#Workbook where the macro (if there is one) is found
		self.macroModule=""			#Module of the macro
		self.configFile="C:\Users\ejoanyk\Desktop\ExHaUST\Config.txt" #path just here for now to be able to test the methods 
		self.adminFile=""			#Full path to admin.txt
		
		self.fullConfigFile=[]			#All of the text in Config.txt, formatted into a list seperating the rows
		self.filesToKeep=[]			#Will this be used? 
		
		self.currentRow=0			#Counter allowing to loop through rows in Config.txt
		self.lastRow=0				#Index of the last row in Config.txt
		self.nrConAttemptsAllowed=0		#Numberr of attempts to connect to Hansoft that we allow

		self.runMacro=False			#Indicates whether a macro will be run or not
		self.reportRequested=False		#True if the export is a report
		self.queryRequested=False		#True if the export is a query

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
		
		self.ReadLineConfigFile(self.configFile) 		#Calls ReadLineConfigFile to collect all of the information
									#from Config.txt.
		
		roughConfigLine= shlex.split(self.fullConfigFile[self.currentRow], posix=False)	#Splits the current row into
												#a list with one field at 
												#each index.

		for j in roughConfigLine:						#Loops through the row to filter out
		    roughConfigLine[roughConfigLine.index(j)] = j.strip("\"")		#quotations marks.

		self.hansoftServer=roughConfigLine[0]					#Sets the first basic attributes
		self.hansoftPort=roughConfigLine[1]					#of ProgramFlow
		self.hansoftDatabase=roughConfigLine[2]
		self.sdkUser=roughConfigLine[3]
		self.sdkPass=roughConfigLine[4]
		self.hansoftProjectName=roughConfigLine[5]
		roughConfigLine[6]=roughConfigLine[6].strip("-")			#Removes the dash from Project View field
		self.hansoftProjectView=roughConfigLine[6]				#since this is the syntax used for HansoftExport.

		def HandleMacroFields(indexOfRunMacro):				#Function for handling the fields concerning macros, and the way they
										#might be at different indices depending on the report and query fields.
										#Takes the index of the runMacro bool in roughCOnfigLine as input.
			
			if roughConfigLine[indexOfRunMacro]=="-m":		
										
				self.runMacro=True
				self.macroWorkbook=roughConfigLine[indexOfRunMacro+1]
				self.macroModule=roughConfigLine[indexOfRunMacro+2]
				self.macroName=roughConfigLine[indexOfRunMacro+3]
				if len(roughConfigLine)==roughConfigLine[indexOfRunMacro+4]:   		#Checks if  fileToBeTransfered field is filled in,
					self.fileToBeTransfered=roughConfigLine[indexOfRunMacro+4]	
			return

		if roughConfigLine[7]=="-r":				#if report
			self.reportRequested=True
			self.reportName=roughConfigLine[8]
			self.reportAuthor=roughConfigLine[9]
			self.finalFolder=roughConfigLine[10]
			self.finalFileName=roughConfigLine[11]
			indexOfRunMacro=12	
			HandleMacroFields(12)

		elif roughConfigLine[7]=="-q":				#if query
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
		while currentRow <= lastRow:
	            UpdateVars()    #update variable values for a new user input in config file
	
	            ##check that the input-row in config file is ok.
	            if LineOK() == False:
	                print("Could not process row " + str(currentRow) + " in " + configFile + ". Please check the input for errors, moving on to next row..")
	                currentRow = currentRow + 1
	                continue
	
	            ##run hansoft export.
	            RunHansoftExport(hansoftExtractFolder, cmdLine)
	            if wasExportSuccessful(latestExport) == False:
	                print("Could not succeed retrieving data from Hansoft. Please check row " + str(currentRow) + " in " + configFile + " for errors, moving on to next row..") 
	                currentRow = currentRow + 1
	                continue
	
	            ##run macros if specified.
	            if runMacro == True:
	                try:
	                    RunExcelMacro(macrosFolder, macroWorkbook, macroModule, macroName)
	                except:
	                    print("Error with running macro:" + macroName + ". Check that the parameters specified in line " + str(currentRow) + " in " + configFile + " are correct.") 
	                    currentRow = currentRow + 1
	                    continue
	
	            ##transfer the file to its final location if the specified folder is different than the altered exports folder.
	            if finalFolder != alteredExportsFolder:
	                try:
	                    TransferFile(alteredExportsFolder, fileToBeTransferred, finalFolder, finalFileName)
	                    print("File" + fileToBeTransferred + " successfully moved to " + finalFolder + " as " + finalFileName + "."
	                except:
	                    print("Error with transferring file: " + alteredExportsFolder + fileToBeTransferred + " to " + finalFolder + " as " + finalFileName + ". Check that the parameters specified in line " +str(currentRow) + " in " + configFile + " are correct.") 
	
	            ##move on to the next row in config file.
	            currentRow = currentRow + 1

	        ##do the clean-up                  
	        RemoveFiles(alteredExportsFolder)
	
	        ##exit program
	        print("Reached the end of + " configFile + ". Shutting down..."
	        return

	def ShouldFileBeTransfered(self):
		return 


g=ProgramFlow()
g.UpdateVars(0)
