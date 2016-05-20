Option explicit
On Error Resume Next
' PT
' Desenvolvido por André Nunes - andre.castro.nunes@gmail.com
' Este script é disponibilizado assim, "as is", sem qualquer tipo de garantia.
' Pode ser modificado à vontade.
'
' EN
' Developed by André Nunes - andre.castro.nunes@gmail.com
' This script is offered as is without any kind of warranty.
' Feel free to modify it at your own will.
'
'
'Auxiliar routines
Sub DisplayUsage
  WScript.Echo "Usage:" & vbCrLf
  WScript.Echo "cscript " & WScript.ScriptName & " -file|-f <path\upalert.log> [-smtp <SMTP server> " & _
  " -to|-t <semicolon separated e-mail addresses> -from <e-mail address>] [-verbose|-v]" & vbCrLf
  WScript.Echo  Wscript.ScriptName & " [-help|-?]" & vbCrLf
  WScript.Echo ""
  WSCript.Quit
End Sub

Sub AppendLine(filespec,Line)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(filespec, ForAppending,True)
	objTextFile.WriteLine(Line)
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO= Nothing
End Sub

Sub OverWriteLine(filespec,Line)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.CreateTextFile(filespec,True)
	objTextFile.Write(Line)
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO= Nothing
End Sub

Function ReadLastLine(filespec)
	Dim retString, objFSO, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(filespec, ForReading, False)
	Do While objTextFile.AtEndOfStream <> True 
      retString = objTextFile.ReadLine 
	Loop 
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO= Nothing
	ReadLastLine = retString 
End Function 

Function FileStatus(filespec)
   Dim objFSO, msg
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If (objFSO.FileExists(filespec)) Then
      msg = True
   Else
      msg = False
   End If
   FileStatus = msg
End Function

Function SendMail(strSMTPServer, strFrom, strTo, strSubject, strBody)
	' Send by connecting to port 25 of the SMTP server.
	Dim iMsg 
	Dim iConf 
	Dim Flds 
	Dim strHTML

	Const cdoSendUsingPort = 2
	set iMsg = CreateObject("CDO.Message")
	set iConf = CreateObject("CDO.Configuration")

	Set Flds = iConf.Fields

	' Set the CDOSYS configuration fields to use port 25 on the SMTP server.
	With Flds
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		.Update
	End With
	
	' Build HTML for message body.
	strHTML = "<HTML>"
	strHTML = strHTML & "<HEAD>"
	strHTML = strHTML & "<BODY>"
	strHTML = strHTML & "<b> " & strBody & "</b></br>"
	strHTML = strHTML & "</BODY>"
	strHTML = strHTML & "</HTML>"

	' Apply the settings to the message.
	With iMsg
		Set .Configuration = iConf
		.To = strTo
		.From = strFrom
		.Subject = strSubject
		.HTMLBody = strHTML
		'.AddAttachment "C:\temp\file.txt"
		.Send
	End With
	
	' Clean up variables.
	Set iMsg = Nothing
	Set iConf = Nothing
	Set Flds = Nothing
End Function

'Main routine
Dim objFSO, objTextFile, strPathtoFileTXT, strPathtoFile, filespec, _
strAvailable, arrLastStatusMsg, strStatusAux, transport_order, system_id, client, u_switch, profile, _
strSMTPserver, strFrom, strTo, strSubject, strBody, oArgs, ArgNum, verboseoutput

Const ForReading = 1 
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001

Set oArgs = WScript.Arguments
ArgNum = 0

While ArgNum < oArgs.Count
  Select Case LCase(oArgs(ArgNum))
	 Case "-file","-f":
		ArgNum = ArgNum + 1
		strPathtoFile = oArgs(ArgNum)
	 Case "-smtp":
		ArgNum = ArgNum + 1
		strSMTPServer = oArgs(ArgNum)
	 Case "-to","-t":
		ArgNum = ArgNum + 1
		strTo = oArgs(ArgNum)
	 Case "-from":
		ArgNum = ArgNum + 1
		strFrom = oArgs(ArgNum)		
	 Case "-verbose","-v":
		verboseoutput = True
	 Case "-help","-?":
		Call DisplayUsage
	 Case Else:
		Call DisplayUsage
  End Select
  ArgNum = ArgNum + 1
Wend

If oArgs.Count=0 Or strPathtoFile="" Then Call DisplayUsage
If strSMTPServer <> "" and strTo = "" Then Call DisplayUsage
If strSMTPServer <> "" and strFrom = "" Then Call DisplayUsage

strPathtoFileTXT = strPathtoFile & ".TXT"

If FileStatus(strPathtoFile) Then 'File exists 
	If Not FileStatus(strPathtoFileTXT) Then
		Call AppendLine(strPathtoFileTXT,"SUM stopped;" & Now())
		strSubject = "SUM stopped"
		strBody = "SUM waiting for input since " & Now() & "..."
		If strSMTPserver <> "" Then
			SendMail strSMTPserver, strFrom, strTo, strSubject, strBody
		End If
		If verboseoutput Then Wscript.Echo "SUM stopped;" & Now()
		WScript.Quit
	Else	
		strStatusAux = ReadLastLine(strPathtoFileTXT)
		arrLastStatusMsg = Split(strStatusAux, ";")
		If arrLastStatusMsg(0) = "SUM stopped" Then 
			'Warning already sent: do nothing
			'Msgbox "Warning already sent: do nothing",64,"SUM alert"
			WScript.Quit
		Else
			'Save SUM state and send e-mail
			Call AppendLine(strPathtoFileTXT,"SUM stopped;" & Now())
			strSubject = "SUM stopped"
			strBody = "SUM waiting for input since " & Now() & "..."
			If strSMTPserver <> "" Then
				SendMail strSMTPserver, strFrom, strTo, strSubject, strBody
			End If
			If verboseoutput Then Wscript.Echo "SUM stopped;" & Now()
		End If
	End If
Else
	strStatusAux = ReadLastLine(strPathtoFileTXT)
	arrLastStatusMsg = Split(strStatusAux, ";")
	If arrLastStatusMsg(0) = "SUM running" Then 
		'Warning already sent: do nothing
		'Msgbox "SUM status didn't change",64,"SUM alert"
		WScript.Quit
	Else
		Call AppendLine(strPathtoFileTXT,"SUM running;" & Now())
		'Msgbox "SUM running again",64,"SUM alert"
		If verboseoutput Then Wscript.Echo "SUM running;" & Now()
	End If
End If

Set strPathtoFile = Nothing
Set strPathtoFileTXT = Nothing
Set strStatusAux = Nothing
Set arrLastStatusMsg = Nothing
Set verboseoutput = Nothing
