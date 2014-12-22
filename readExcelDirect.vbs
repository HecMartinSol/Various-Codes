'On Error Resume Next
'------------------------------------------------------------------------------------------------------------
fileName = "Filtro 2014-12-14_ES_Computers.xlsx"
'fileName = "Filtro 2014-12-14_ES_Users.xlsx"
'fileName = "Filtro 2014-12-14_ES_Groups.xlsx"

'sheetNames = Array("windows xp")
'sheetNames = Array("Windows 7")
sheetNames = Array("WindowsXP","Windows7")

collumns = Array("D","I")
'collumns = Array("A","B","C","D","E","F","G","H","I","J","K")

beginRow = 2
'------------------------------------------------------------------------------------------------------------


Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
filePath = WshShell.CurrentDirectory & "\" & fileName

Set objApp = CreateObject("Excel.Application")
Set objWbs = objApp.WorkBooks
objApp.Visible = False
Set objWorkbook = objWbs.Open(filePath)


	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(WshShell.CurrentDirectory & "\dump.txt",2,true)
	lineData = ""



' Iterates every sheet declared
for each sheet_name in sheetNames

	Set objSheet = objWorkbook.Sheets(sheet_name)
	
	row = beginRow	
	
	' Iterates every row until an empty one
	do
		' Iterates every collumn declared
		for each collumn_letter in collumns
		
					valueAt_c_r =  objSheet.Range(collumn_letter & row).Value
					if valueAt_c_r <> "" then
						
						' Manage the data
						linedata = linedata & valueAt_c_r & " "
						
					end if
		next
		
		objFileToWrite.WriteLine(lineData)
		lineData = ""

		
		row = row + 1
	loop until valueAt_c_r = ""
	
next




objFileToWrite.Close
Set objFileToWrite = Nothing
 
objWorkbook.Close False
objWbs.Close 
objApp.Quit 
 
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objWbs = Nothing
Set objApp = Nothing


MsgBox "DONE!",, "Success"
