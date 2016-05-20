# SAP-checkSUM
Script to monitor the SAP upgrade (whenever SUM is waiting for input).
Usage:

cscript checkSUM.vbs -file|-f <path\upalert.log> [-smtp <SMTP server>  -to|-t <semicolon separated e-mail addresses> -from <e-mail address>] [-verbose|-v]

checkSUM.vbs [-help|-?]


The file parameter  is the upalert.log that is created whenever SUM stops waiting for input.
For example, X:\usr\sap\SUM\abap\tmp\upalert.log

A log called X:\usr\sap\SUM\abap\tmp\upalert.log.TXT is created with all the stops and starts.
Example:
SUM stopped;18-05-2016 10:45:10

SUM running;18-05-2016 10:45:46

SUM stopped;18-05-2016 10:47:27

SUM running;18-05-2016 10:54:39

SUM stopped;18-05-2016 14:51:49

This script should run at certain intervals in background.

To avoid a command prompt window from showing up, the following script can be used:

Const HIDDEN_WINDOW = 12
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objStartup = objWMIService.Get("Win32_ProcessStartup")

Set objConfig = objStartup.SpawnInstance_
objConfig.ShowWindow = HIDDEN_WINDOW
Set objProcess = GetObject("winmgmts:root\cimv2:Win32_Process")
errReturn = objProcess.Create("C:\Program Files\Java\j2re1.4.2\bin\java.exe -classpath C:\SAPDownloadManager\DLManager.jar dlmanager.Application", null, objConfig, intProcessID)



