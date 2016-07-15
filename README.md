# HE

HE is a script used for extracting information from a Hansoft database to an Excel workbook saved at a specified location, it also provides the option of running VBA macros in the export workbook.

HE uses HansoftExport (1) for the extraction from Hansoft. HansoftExport in turn uses EPPlus (2), SimpleLogging (3) and ObjectWrapper (4). All of these programs can be found on github and need to be built succesfully for HE to run.

(1) https://github.com/Hansoft/HansoftExport
(2) http://epplus.codeplex.com/
(3) https://github.com/Hansoft/Hansoft-SimpleLogging
(4) https://github.com/Hansoft/Hansoft-ObjectWrapper

**************************************************************************************************************************************
CONFIG.TXT

WARNING: Files saved in the sourceFolder path will be removed by the script after it is run.



If your maco resaves the file, this must be specified in ________________.
_____________________________________________________________________________________________________________________________________
ADDING SEVERAL EXPORTS AS SHEETS IN THE SAME WORKBOOK:

For every export that you want saved as a sheet in this workbook, except the last one, set finalFolder=sourceFolder.

Create a macro that copies export content to sheets in the workbook, the way you want it to. Specify that this macro will be run on the export.

For the final export that you want saved as a sheet in this workbook, set finalFolder as the destination file where you want this workbook to be saved. 


_____________________________________________________________________________________________________________________________________
**************************************************************************************************************************************

**************************************************************************************************************************************
MACRO REQUIREMENTS

- Only ONE macro can be used for each export, however, this macro in turn can call other macros if desired.
- It must be specifies in the macro that it is to be run on LatestReport.xlsx in sourceFolder path.
- It is required to specify in config.txt the final name of the export after the macros have been run.
- The macro itself needs to save the workbook in sourceFolder path.
- 


**************************************************************************************************************************************



