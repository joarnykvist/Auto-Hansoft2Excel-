#coding: utf-8


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
		return

	def LineOK(self)
		return	

	def BuildCMDLine(self):
		return

	def UpdateVars(self):
		return

	def ControlFlow(self):
		return

	def ShouldFileBeTransfered(self):
		return 


print("tobias härjar")
g=ProgramFlow()
g.SetAdminVars()
