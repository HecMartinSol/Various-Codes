Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

' CALL SCRIPT WITH USING FILENAME AS FIRST ARGUMENT
fileName = WScript.Arguments.Item(0)


filePath = WshShell.CurrentDirectory & "\" & fileName

' Check if file exists in the current path
Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(filePath)) Then

	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' IMAGE FILES 																																													[OK]
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	if Instr(1, fileName, ".png") > 0 or Instr(1, fileName, ".jpg") > 0 or Instr(1, fileName, ".jpeg") or Instr(1, fileName, ".gif") or Instr(1, fileName, ".bmp") then

		Set objExplorer = CreateObject("InternetExplorer.Application")
			
		Set oImage = CreateObject("WIA.ImageFile")
		oImage.LoadFile filePath

		With objExplorer
			.Navigate "about:blank"
			.Visible = 1
			.Fullscreen = False
			.Document.Title = fileName
			.Toolbar=False
			.Statusbar=False
			.Top=0
			.Left=0
			.Height=oImage.Height+100
			.Width=oImage.Width+100
			.Document.Body.InnerHTML = "<style type='text/css'> img { width:100%;} </style>" & _
														"<img src='" & filePath & "' max-width = '100%'>"
		End With

		
		
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' PDF FILES 																																														[OK]
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	elseif Instr(1, fileName, ".pdf") > 0  then
	
		'*************Launch IExplorer (IE10 & ¿IE11?) with pdf embedded*************
		Set objExplorer = CreateObject("InternetExplorer.Application")
			
		With objExplorer
			.Navigate "about:blank"
			.Visible = 1
			.Fullscreen = False
			.Document.Title = fileName
			.Toolbar=False
			.Statusbar=False
			.Top=0
			.Left=0
			.Height=1024
			.Width=720
			.Document.Body.InnerHTML = "" & _ 
				"<html>" & vbCrLf & _
				"<body>" & vbCrLf & _
					"<object data='" & filePath & "' type='application/pdf' width = '100%' height = '100%'>" & vbCrLf & _
						"<embed src='" & filePath & "' type='application/pdf' />" & vbCrLf & _
					"</object>" & vbCrLf & _
				"</body>" & vbCrLf & _
				"</html>"

		End With
		'***********************************************************************

			
		
		'****************************Other launch mode****************************
		' Just runs file with the default program, maximized
		'WshShell.run fileName, 3
		
		' Toggles fullscreen mode in PDF files, sending KeyStroke
		'WScript.Sleep 500
		'WshShell.Sendkeys "^l"
		'***********************************************************************
		
		
		
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' TEXT FILES 																																														[OK]
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	elseif Instr(1, fileName, ".txt") > 0 then
	
		'*************Launch IExplorer (IE10 & ¿IE11?) with txt embedded*************
		Set objExplorer = CreateObject("InternetExplorer.Application")
					
		With objExplorer
			.Navigate "about:blank"
			.Visible = 1
			.Fullscreen = False
			.Document.Title = fileName
			.Toolbar=False
			.Statusbar=False
			.Top=0
			.Left=0
			.Height=1024
			.Width=720
			.Document.Body.InnerHTML = "" & _ 
			"<div>" & vbCrLf & _
				 "<object data='" & filePath & "' type='text/plain' width = '100%' height = '100%'>" & vbCrLf & _
					"<a href='" & filePath & "'></a>" & vbCrLf & _
				 "</object>" & vbCrLf & _
			 "</div>" 
			 
		End With
		'***********************************************************************
		
		
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' PRESENTATION FILES 																																										[OK]
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	elseif Instr(1, fileName, ".pps") > 0 or Instr(1, fileName, ".ppsx") > 0 then
		'Just runs file with the default program, maximized
		WshShell.run fileName, 3

	

	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' MICROSOFT POWERPOINT  FILES																																						[OK]
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	elseif  Instr(1, fileName, ".ppt") > 0  or Instr(1, fileName, ".pptx") > 0 then 
			
		'********************Open and set auto-advance slideshow********************		
		Const ppAdvanceOnTime = 2	' Run according to timings (not clicks)
		Const ppShowTypeKiosk = 3	' Run in "Kiosk" mode (fullscreen)
		Const ppAdvanceTime = 5     ' Show each slide for 5 seconds

		
		Set objPPT = CreateObject("PowerPoint.Application")
		objPPT.Visible = True

		' Open presentation in Read-Only mode
		Set objPresentation = objPPT.Presentations.Open(filePath,True)
		
		' Calculates the seconds to be displayed, according to the number of slides
		seconds = objPresentation.Slides.Count * ppAdvanceTime
		
		' Apply powerpoint settings
		objPresentation.Slides.Range.SlideShowTransition.AdvanceOnTime = True
		objPresentation.Slides.Range.SlideShowTransition.AdvanceOnClick = False
		objPresentation.SlideShowSettings.AdvanceMode = ppAdvanceOnTime 
		objPresentation.SlideShowSettings.ShowType = ppShowTypeKiosk
		objPresentation.Slides.Range.SlideShowTransition.AdvanceTime = ppAdvanceTime
		objPresentation.SlideShowSettings.LoopUntilStopped = True

		' Run the slideshow
		Set objSlideShow = objPresentation.SlideShowSettings.Run.View
		
		' Waits until te slideshow ends
		WScript.Sleep seconds*1000'milliseconds
		
		objPresentation.Saved = True
		objPresentation.Close
		objPPT.Quit
		'***********************************************************************

		
		'***************************Open standard mode***************************
		'Set powerpointObject = WScript.CreateObject("Powerpoint.Application")
		'powerpointObject.Visible = True
		
		'Call powerpointObject.Presentations.Open(filePath,,True)
		'***********************************************************************
		
	
		'********************Open and set fullscreen mode **************************
		'***************************Test mode***********************************
		'Set powerpointObject = CreateObject("PowerPoint.Application")
		'	powerpointObject.Visible = True
		'
		'Set objPresentation = powerpointObject.Presentations.Open(filePath, True)
		'
		'Set objSlideShow = objPresentation.SlideShowSettings.Run.View
		'***********************************************************************

		
		
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' MICROSOFT WORD FILES																																									[OK]
	'---------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	elseif  Instr(1, fileName, ".doc") > 0  or Instr(1, fileName, ".docx") > 0  then
		
		'***********************************************************************
		' Based on: http://www.robvanderwoude.com/vbstech_automation_word.php#SaveAsHTML
		' Save current doc file as HTML and displays it within IExplorer
		
		Const wdFormatFilteredHTML= 10
		
		
		Set wordObject = WScript.CreateObject("Word.Application")
		wordObject.Visible = False
		Call wordObject.Documents.Open(filePath)
					
				
		if not fso.FileExists(filePath&".html") then
			strHTML = fso.BuildPath( WshShell.CurrentDirectory, fileName & ".html" )
			wordObject.ActiveDocument.SaveAs strHTML, wdFormatFilteredHTML
		end if
		
		
		Set dict = CreateObject("Scripting.Dictionary")
		Set file = fso.OpenTextFile (WshShell.CurrentDirectory & "\" & fileName & ".html", 1)
		row = 0
		Do Until file.AtEndOfStream
		  line = file.Readline
		  dict.Add row, line
		  row = row + 1
		Loop

		file.Close

		htmlEmbed = ""
		'Loop over it
		For Each line in dict.Items
		   htmlEmbed = htmlEmbed & line
		Next
		
		
		wordObject.ActiveDocument.Close
		wordObject.Quit	
		fso.DeleteFile fileName & ".html"

		
		Set objExplorer = CreateObject("InternetExplorer.Application")
					
		With objExplorer
			.Navigate "about:blank"
			.Visible = 1
			.Fullscreen = False
			.Toolbar=False
			.Statusbar=False
			.Top=0
			.Left=0
			.Height=1024
			.Width=720
			.Document.Body.InnerHTML = "" & htmlEmbed
		End With
		'***********************************************************************
		
		
		
		
		'***************Open Word fullscreen. Dispose when not-fullsceen***************		
		'Set wordObject = WScript.CreateObject("Word.Application")
		'wordObject.Visible = True
		'Call wordObject.Documents.Open(filePath,,True)
		'		
		'wordObject.ActiveWindow.View.Fullscreen = True
		'
		'do while wordObject.ActiveWindow.View.Fullscreen
		'	WScript.sleep 1
		'loop
		'
		'wordObject.Quit		
		'***********************************************************************
		
	end if

'else
'	WScript.Echo "The file '" & filePath & "' does not exists"
end if
		
