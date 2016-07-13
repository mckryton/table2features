Attribute VB_Name = "basSystem"
'------------------------------------------------------------------------
' Description  : extends system related functions
'------------------------------------------------------------------------

' Copyright 2016 Matthias Carell
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

'Declarations

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : checks if item exists in a collection object
' Parameter     : pvarKey           - item name
'                 pcolACollection   - collection object
' Returnvalue   : true if item exits, false if not
'-------------------------------------------------------------
Public Function existsItem(pvarKey As Variant, pcolACollection As Collection) As Boolean
                    
    Dim varItemValue As Variant
                     
    On Error GoTo NOT_FOUND
    varItemValue = pcolACollection.Item(pvarKey)
    On Error GoTo 0
    existsItem = True
    Exit Function
                     
NOT_FOUND:
    existsItem = False
End Function
'-------------------------------------------------------------
' Description   : prints log messages to direct window
' Parameter     :   pstrLogMsg      - log message
'                   pintLogLevel    - log level for this message
'-------------------------------------------------------------
Public Sub log(pstrLogMsg As String, Optional pintLogLevel)

    Dim intLogLevel As Integer      'aktueller Loglevel
    Dim strLog As String            'auszugebender Text
    
    'default log level is cLogInfo
    If IsMissing(pintLogLevel) Then
        intLogLevel = cLogInfo
    Else
        intLogLevel = pintLogLevel
    End If
   
    'print log message only if given log level is lower or equal then
    ' log level set in module basConstants
    If intLogLevel <= cCurrentLogLevel Then
        'start with current time
        strLog = Time
        'add log level
        Select Case intLogLevel
            Case cLogDebug
                strLog = strLog & " debug:"
            Case cLogInfo
                strLog = strLog & " info:"
            Case cLogWarning
                strLog = strLog & " warning:"
            Case cLogError
                strLog = strLog & " error:"
            Case cLogCritical
                strLog = strLog & " critical:"
            Case Else
                strLog = strLog & " custom(" & intLogLevel & "):"
        End Select
        'add log message
        strLog = strLog & " " & pstrLogMsg
        Debug.Print strLog
    End If
End Sub
'-------------------------------------------------------------
' Description   : function print error messages to the direct window
' Parameter     : pstrFunctionName  - name of the calling function
'                 pstrLogMsg        - optional: custom error message
'-------------------------------------------------------------
Public Sub log_error(pstrFunctionName As String, Optional pstrLogMsg As Variant)

    Dim intLogLevel As Integer      'current log level
    Dim strLog As String            'complete log messages
    Dim strError As String          'system error message from Err object
    
    strError = Err.Description
    'start log messages with time
    strLog = Time
    'log level = error
    strLog = strLog & " error:"
    'add caller name
    strLog = strLog & "error in " & pstrFunctionName & ": "
    'if given add custom log message
    If Not IsMissing(pstrLogMsg) Then
        strLog = strLog & " " & pstrLogMsg
    Else
        'use message from Err object
        On Error Resume Next
        strLog = strLog & " " & strError
    End If
    Debug.Print strLog
    'reset cursor, screen update status and statusbar
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub
'-------------------------------------------------------------
' Description   : alias to log function with cLogDebug level
' Parameter     :   pstrLogMsg      - log message
'-------------------------------------------------------------
Public Sub logd(pstrLogMsg As String)

    On Error GoTo error_handler
    basSystem.log pstrLogMsg, cLogDebug
    Exit Sub

error_handler:
    basSystem.log_error "basSystem.logd"
End Sub
'-------------------------------------------------------------
' Description   : save source code as text files
'-------------------------------------------------------------
Private Sub exportCode()

    Dim vcomSource As VBComponent
    Dim strPath As String
    Dim strSeparator As String
    Dim strSuffix As String

    On Error GoTo error_handler
    #If Mac Then
        strSeparator = ":"
    #Else
        strSeparator = "\"
    #End If
    strPath = ThisWorkbook.Path & strSeparator & "source"
    For Each vcomSource In Application.VBE.VBProjects("table2features").VBComponents
        Select Case vcomSource.Type
            Case vbext_ct_StdModule
                strSuffix = "bas"
            Case vbext_ct_ClassModule
                strSuffix = "cls"
            Case vbext_ct_Document
                strSuffix = "cls"
            Case vbext_ct_MSForm
                strSuffix = "frm"
            Case Else
                strSuffix = "txt"
        End Select
        vcomSource.Export strPath & strSeparator & vcomSource.Name & "." & strSuffix
        basSystem.log "export code to " & strPath & strSeparator & vcomSource.Name & "." & strSuffix
    Next
    Exit Sub

error_handler:
    basSystem.log_error "basSystem.exportCode"
End Sub
