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

To avoid a command prompt window from showing up, the attached script runHidden.vbs can be used.



