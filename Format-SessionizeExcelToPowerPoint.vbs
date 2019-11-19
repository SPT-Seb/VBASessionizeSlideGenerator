' Author : Sébastien Paulet (@SP_twit) - aOS Community
' Based on an Excel file exported from sessionize.com, populate a PowerPoint 
' template to generate picture to announce each conference based on first PowerPoint slide.
' args0 - source Excel file fullpath
' args1 - source PowerPoint file fullpath
'---------------------------------

	Dim args, excelFileFullPath, powerPointFileFullPath
	Set args = WScript.Arguments
	excelFileFullPath = args(0)
	powerPointFileFullPath = args(1)
	
'create fileSystem object
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	
' Source Richard Bendall https://gist.github.com/Richienb/51021a1c16995a07478dfa20a6db725c 	
Sub HTTPDownload( myURL, myPath )
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
	'MsgBox "Download " & myURL & " in " myPath
    If objFileSystem.FolderExists( myPath ) Then
        strFile = objFileSystem.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
    ElseIf objFileSystem.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
        strFile = myPath
    Else
        WScript.Echo "ERROR: Target folder not found."
        Exit Sub
    End If
    Set objFile = objFileSystem.OpenTextFile( strFile, ForWriting, True )
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
    objHTTP.Open "GET", myURL, False
    objHTTP.Send
    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next
    objFile.Close( )
	'MsgBox "Download done"
End Sub	

Sub ProcessExcelRowToPowerPointSlide( rowIndex, sh, objExcel, objPresentation)
	Dim sessionId, sessionTitle, sessionSpeakerNameList, sessionLanguage, sessionSpeakersIDList
	Dim sessionSpeaker1Name, sessionSpeaker1ID, sessionSpeaker1PictureUrl, sessionSpeaker1TagLine
	Dim sessionSpeaker2Name, sessionSpeaker2ID, sessionSpeaker2PictureUrl, sessionSpeaker2TagLine

	'Read Excel Data
	sessionId = sh.Cells(rowIndex, 1) 'Id column
	sessionTitle = sh.Cells(rowIndex, 2) 'Title column
	sessionSpeakerNameList = Split(sh.Cells(rowIndex, 4), ",") 'Speakers column
	sessionLanguage = sh.Cells(rowIndex, 10) 'Language column
	sessionSpeakersIDList = Split(sh.Cells(rowIndex, 14), ",") 'Speaker Ids column
	
	sessionSpeaker1Name = sessionSpeakerNameList(0)
	sessionSpeaker1ID = sessionSpeakersIDList(0)
	sessionSpeaker1PictureUrl = objExcel.WorksheetFunction.VLookup(sessionSpeaker1ID, objExcel.Range("'All Speakers'!A1:I200"), 9, False) 'Profile Picture column
	sessionSpeaker1TagLine = objExcel.WorksheetFunction.VLookup(sessionSpeaker1ID, objExcel.Range("'All Speakers'!A1:I200"), 5, False) 'Tagline column
	
	If Ubound(sessionSpeakersIDList) > 0 Then
		sessionSpeaker2Name = Trim(sessionSpeakerNameList(1))
		sessionSpeaker2ID = Trim(sessionSpeakersIDList(1))
		sessionSpeaker2PictureUrl = objExcel.WorksheetFunction.VLookup(sessionSpeaker2ID, objExcel.Range("'All Speakers'!A1:I200"), 9, False) 'Profile Picture column
		sessionSpeaker2TagLine = objExcel.WorksheetFunction.VLookup(sessionSpeaker2ID, objExcel.Range("'All Speakers'!A1:I200"), 5, False) 'Tagline column
	End If
	
	'MsgBox sessionTitle
	'download speaker picture		
	speaker1PicFullPath = objPresentation.Path + "\tmpSpeakersPics\" + sessionSpeaker1ID + "." + objFileSystem.GetExtensionName(sessionSpeaker1PictureUrl)
	HTTPDownload sessionSpeaker1PictureUrl, speaker1PicFullPath
	If Ubound(sessionSpeakersIDList) > 0 Then
		speaker2PicFullPath = objPresentation.Path + "\tmpSpeakersPics\" + sessionSpeaker2ID + "." + objFileSystem.GetExtensionName(sessionSpeaker2PictureUrl)
		HTTPDownload sessionSpeaker2PictureUrl, speaker2PicFullPath
	End If
	
	'Duplicate Slide
	newSlideTemplateIndex = Ubound(sessionSpeakersIDList) + 1 
	Set newSlide = objPresentation.Slides(newSlideTemplateIndex).Duplicate
	newSlide.MoveTo(objPresentation.Slides.Count)

	'Fill Slide
	For Each oShape In newSlide.Shapes
		If oShape.Type = 17 Then 
			If oShape.TextFrame.TextRange.Text = "*Title*" Then
				oShape.TextFrame.TextRange.Text = sessionTitle
			ElseIf oShape.TextFrame.TextRange.Text = "*Speaker1*" Then
				oShape.TextFrame.TextRange.Text = sessionSpeaker1Name
			ElseIf oShape.TextFrame.TextRange.Text = "*TagLineSpeaker1*" Then
				oShape.TextFrame.TextRange.Text = sessionSpeaker1TagLine
			ElseIf oShape.TextFrame.TextRange.Text = "*Speaker2*" Then
				oShape.TextFrame.TextRange.Text = sessionSpeaker2Name
			ElseIf oShape.TextFrame.TextRange.Text = "*TagLineSpeaker2*" Then
				oShape.TextFrame.TextRange.Text = sessionSpeaker2TagLine
			End If 
		ElseIf oShape.Type = 1 Then 'msoAutoShape
			If oShape.TextEffect.Text = "*SpeakerPic1*" Then
				oShape.TextEffect.Text = ""
				oShape.Fill.UserPicture speaker1PicFullPath
			ElseIf oShape.TextEffect.Text = "*SpeakerPic2*" Then
				oShape.TextEffect.Text = ""
				oShape.Fill.UserPicture speaker2PicFullPath
			End If
		End If
	Next
	
	'Export Slide as JPG
	sImagePath = objPresentation.Path
	sPrefix = Split(objPresentation.Name, ".")(0)
	sImageName = sPrefix & "-" & sessionId & ".jpg"
	newSlide.Export sImagePath & "\sessionsOut\" & sImageName, "JPG"
End Sub	
	
'create the excel object
	Set objExcel = CreateObject("Excel.Application") 
	objExcel.Visible = True 
	Set objWorkbook = objExcel.Workbooks.Open(excelFileFullPath)

'create the PowerPoint object
	Set objPowerPoint = CreateObject("Powerpoint.Application") 
	objPowerPoint.Visible = True 
	Set objPresentation = objPowerPoint.Presentations.Open(powerPointFileFullPath)

'create needed folders on fileSystem
	If objFileSystem.FolderExists( objPresentation.Path + "\tmpSpeakersPics" ) Then
		objFileSystem.DeleteFolder(objPresentation.Path + "\tmpSpeakersPics")
	End If
	objFileSystem.CreateFolder(objPresentation.Path + "\tmpSpeakersPics")
	
	If objFileSystem.FolderExists( objPresentation.Path + "\sessionsOut" ) Then
		objFileSystem.DeleteFolder(objPresentation.Path + "\sessionsOut")
	End If
	objFileSystem.CreateFolder(objPresentation.Path + "\sessionsOut")
	
'go through Excel rows	
	objWorkbook.Sheets("Accepted Sessions").Select
	
	Set sh = objWorkbook.ActiveSheet
	
	For Each rw In sh.Rows
		If rw.Row <> 1 Then
            If sh.Cells(rw.Row, 1).Value = "" Then
                Exit For
            End If
            
			On Error Resume Next 
			ProcessExcelRowToPowerPointSlide rw.Row, sh, objExcel, objPresentation
			
			If err.Number <> 0 Then : MsgBox "Erreur lors du traitement de la ligne " & rw.Row : End If
        End If
    Next

'save Powerpoint
	objPresentation.Save

'close Office objects
	objWorkbook.Close 
	objPresentation.Close

'exit Office programs
	objExcel.Quit
	objPowerPoint.Quit

'release Office objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing
	Set objPowerPoint = Nothing
	Set objPresentation = Nothing
	Set objFileSystem = Nothing