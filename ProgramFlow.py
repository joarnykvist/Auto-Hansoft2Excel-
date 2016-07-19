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
		self.configFile="" 			#0719  -------- location for config file
		self.configFile="" 			#path just here for now to be able to test the methods 
		self.adminFile=""			#Full path to admin.txt
		self.macrosFolder=""			#Full path to folder containing macro workbooks.
		
		self.configLine=[] 			#one line in config file, contained as a list where each index contains one "field"
		self.fullConfigFile=[]			#All of the text in Config.txt, formatted into a list seperating the rows
		
		self.currentRow=0			#Counter allowing to loop through rows in Config.txt
		self.lastRow=0				#Index of the last row in Config.txt
		self.nrConAttemptsAllowed=0		#Numberr of attempts to connect to Hansoft that we allow
		
		self.adminFileOK = True 		#0719 added as a flag to indicate whether the program can be run at all
		self.lineOK = True 			#Indicates whether the information from config line is okay. 
		self.runMacro=False			#Indicates whether a macro will be run or not
		self.reportRequested=False		#True if the export is a report
		self.queryRequested=False		#True if the export is a query

	def SetAdminVars(self):
		
		a=open(self.adminFile, 'r')	
		aList = a.readlines() 				#Opens Admin text file and reads every line, creating a list with the lines as elements

		
		strippedList=[]

		for el in aList:
			el= el.split(" - ")				#Filters out everything that is not information that will be used in the program
			strippedList.append(el[1].strip())

		for fileRef in strippedList[0:5]:			#0719 loops through file references to see if they can be accessed 
			if os.path.exists(fileRef)==False:
				self.adminFileOK = False
				break

		print(self.adminFileOK) #0719

		self.exhaustFolder=strippedList[0]
		self.configFile=strippedList[1]
		self.hansoftExtractFolder=strippedList[2]
		self.macrosFolder=strippedList[3]
		self.alteredExportsFolder=strippedList[4]
		self.nrConAttemptsAllowed=strippedList[5]
		self.latestExport=strippedList[6]

		self.ReadLineConfigFile(self.configFile) #Imports all the text from config.txt - this is done here so that we only have to do this once every run.
	
		return 

	def ReadLineConfigFile(self):
		
		openConfigFile=open(configFile, 'r')	 		#Opens config.txt in read-only mode
		
		self.fullConfigFile = openConfigFile.readlines() 	#Creates a list made up of the lines in config files as strings, 
									#e.g. one full line at each index, and sets this list to the fullConfigFile attribute
		return

	def CheckLine(self):
		
		def CheckDirectories(Path):	#Tries changing directories to a given path
			try:
				os.chdir(Path)
			except:
				return False

		def DoesFileExist(FilePath):			#Checks if a given file exists 
			return os.path.exists(FilePath)

		self.lineOK=True #Resets the lineOK variable
		
		if len(self.configLine) < 11: 
			print("Some of the necessary fields in Config were not filled in. (At least 10 fields must be filled in for a basic export)")
			self.lineOK = False #Minimum number of fields filled in is 11
			return

		if self.configLine[6] not in ["-a","-s","-b","-q"]: #a, s, b, q are the allowed choices for the Project View option
			print("Project view option was not filled in properly.")
			self.lineOK = False
			return

		if self.configLine[7]=="-r":
				nextIndex = 10 
														#since exporting a report requires one more argument than exporting a query, the indices get out of sync here
		elif self.configLine[7]=="-q":
				nextIndex = 9
		else: 											#There are only two ways in which this option can be specified. 
			print("Report/query option was not filled in properly")
			self.lineOK = False
			return

		if  CheckDirectories(self.configLine[nextIndex])==False: 
			print("Final Folder could not be accessed.")
			self.lineOK = False 	#Makes sure that the destination folder can be accessed.
			return

		if not self.configLine[nextIndex+1][-5:]==".xlsx" and not self.configLine[nextIndex+1][-4:]==".xls" and not self.configLine[nextIndex+1][-5:]==".xlsm":
			print("The extension for Final FileName was not properly written.")  #0719 changed both 'or' to 'and'
			self.lineOK = False
			return														#Makes sure that an extension is included in the final file name.

		if len(self.configLine)>nextIndex+2: 
			if self.configLine[nextIndex+2]=="-m": 							#Checks that the macro option has been filled in properly, if it has been filled in.
				if len(self.configLine)<nextIndex+5+1: 
					print("Not enough fields were filled in for a macro to be run - worbook, module and macro name are all required.")
					self.lineOK = False			#If a macro is to be run, the module and macro name must be specified.
					return

				if  DoesFileExist(self.macrosFolder + self.configLine[nextIndex + 3])==False: 
					print("Could not access the macro workbook in macros folder.")
					self.lineOK = False 		#Checks that the Workbook with the macro exists in the macros folder.

					return

				if len(self.configLine)>nextIndex+5+1:
					if DoesFileExist(self.alteredExportsFolder + self.configLine[nextIndex+6])==False: 
						print("The resaved workbook could not be accessed in altered exports folder.")
						self.lineOK = False																			#If the Workbook has been saved differently,
						return																						#we here check that the resaved workbook is found.

			else: 
				print("The macro option was not filled in properly, must be filled in as '-m'")
				self.lineOK = False	#Returns False if the macro has been filled in in any other way than "-m".
				return
		return  	#If we get here we consider the line to be okay.
			

	def BuildCMDLine(self):				#Adjusts the input from Config.txt so that it fits the syntax for HansoftExport.
		
		cmdToBe=""                                  #Using cmdToBe as a  local variable to which self.cmdLine is later set. 
		cmdToBe += ( 
		 "-c" + self.hansoftServer + ":" 
		 + self.hansoftPort + ":" 
		 +  self.hansoftDatabase + ":"
		 +  self.sdkUser + ":" 
		 + self.sdkPass 
		 + " -p" + self.hansoftProjectName + ":"
		 + self.hansoftProjectView 				 
		 ) 

		if self.reportRequested == True:     #0719 added '== True'
			cmdToBe += (
				" -r" 
				+ self.reportName + ":"
				+ self.reportAuthor 
				)

		elif self.queryRequested == True: #0719 same as above
			cmdToBe += (
				" -f"
				+ self.query
				) 

		cmdToBe += " -o" + self.latestExport
		self.cmdLine=cmdToBe
		
		return

	def UpdateVars(self):
		
	
		
		self.configLine= shlex.split(self.fullConfigFile[self.currentRow], posix=False)	#Splits the current row into
												#a list with one field at 
												#each index.
		self.CheckLine()
		
		if self.lineOK != True: 
			print("Line from Config.txt not okay.")
			return  
		
		self.cmdLine="" 						#empties the command line between runs  
		
		self.hansoftServer=self.configLine[0]					#Sets the first basic attributes
		self.hansoftPort=self.configLine[1]					#of ProgramFlow
		self.hansoftDatabase=self.configLine[2]
		self.sdkUser=self.configLine[3]
		self.sdkPass=self.configLine[4]
		self.hansoftProjectName=self.configLine[5]
		self.hansoftProjectView=self.configLine[6].strip("-")		#since this is the syntax used for HansoftExport.
		self.lastRow=len(self.fullConfigFile)-1 #0718			#-1 because we call the first row 0.
		
		def HandleMacroFields(indexOfRunMacro):				#Function for handling the fields concerning macros, and the way they
										#might be at different indices depending on the report and query fields.
										#Takes the index of the runMacro bool in self.configLine as input.
			
			if self.configLine[indexOfRunMacro]=="-m":		
										
				self.runMacro=True
				self.macroWorkbook=self.configLine[indexOfRunMacro+1]
				self.macroModule=self.configLine[indexOfRunMacro+2]
				self.macroName=self.configLine[indexOfRunMacro+3]
				if len(self.configLine)==self.configLine[indexOfRunMacro+4]:   		#Checks if  fileToBeTransfered field is filled in,
					self.fileToBeTransfered=self.configLine[indexOfRunMacro+4]	
				else: self.fileToBeTransfered=""
			
			else:
				self.runMacro=False			#resets macro variables if macro field is not filled in.
				self.macroWorkbook=""
				self.macroModule=""
				self.macroName=""
			
			return

		if self.configLine[7]=="-r":				#if report
			self.reportRequested=True
			
			self.queryRequested=False #0719 resets query options if report
			self.query=""				#0719
			
			self.reportName=self.configLine[8]
			self.reportAuthor=self.configLine[9]
			self.finalFolder=self.configLine[10]
			self.finalFileName=self.configLine[11]
			if len(self.configLine)>12:
				indexOfRunMacro=12	
				HandleMacroFields(12)

		elif self.configLine[7]=="-q":				#if query
			
			self.queryRequested=True
			
			self.reportRequested=False #0719 resets report options if query
			self.reportName=""			#0719
			self.reportAuthor=""      #0719
			
			self.query=self.configLine[8]
			self.finalFolder=self.configLine[9]
			self.finalFileName=self.configLine[10]
			if len(self.configLine)>11:
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
