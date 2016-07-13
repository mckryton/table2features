Attribute VB_Name = "basConstants"
'------------------------------------------------------------------------
' Description  : contains all global constants
'------------------------------------------------------------------------

'Options
Option Explicit

'log level (range is 1 to 100)
Global Const cLogDebug = 100
Global Const cLogInfo = 90
Global Const cLogWarning = 50
Global Const cLogError = 30
Global Const cLogCritical = 1

'current log level - decreasing log level means decreasing amount of messages
Global Const cCurrentLogLevel = 90

