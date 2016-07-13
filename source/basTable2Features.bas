Attribute VB_Name = "basTable2Features"
'------------------------------------------------------------------------
' Description  : exports features into feature files
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
' Description   : main routine for generating feature files from a table
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Sub exportFeatures()

    Dim colFeatures As Collection

    On Error GoTo error_handler
    Set colFeatures = basTable2Features.readFeatureFromTable()
    basTable2Features.writeFeaturesToFiles colFeatures
    Exit Sub
    
error_handler:
    basSystem.log_error "basTable2Features.exportFeatures"
End Sub
'-------------------------------------------------------------
' Description   : reads data from a atbel into a collection object
' Parameter     :
' Returnvalue   : collection object containing features as items
'-------------------------------------------------------------
Private Function readFeatureFromTable()

    Dim colFeatures As New Collection
    Dim colSingleFeature As Collection
    Dim strFeatureName As String
    Dim strDomain As String
    Dim strAggregate As String
    Dim strScenario As String
    Dim colScenarios As Collection
    
    Dim rngCurrent As Range
    
    On Error GoTo error_handler
    Set rngCurrent = Selection
    While Trim(rngCurrent.Text) <> ""
        Application.StatusBar = "reading table"
        'read current feature
        strDomain = Replace(rngCurrent.Text, " ", "_")
        strAggregate = Replace(rngCurrent.Offset(, 1).Text, ":", " ")
        strFeatureName = Replace(rngCurrent.Offset(, 2).Text, ":", " ")
        If Trim(strFeatureName) = "" Then
            strFeatureName = "undefined"
        End If
        strScenario = rngCurrent.Offset(, 3).Text
        basSystem.log "read feature " & strAggregate & " - " & strFeatureName
        'try to get the existing feature
        On Error GoTo create_new_feature
        'if not found create a new one
        Set colSingleFeature = colFeatures(strDomain & "-" & strAggregate & "-" & strFeatureName)
        On Error GoTo error_handler
        Set colScenarios = colSingleFeature.Item("scenarios")
        If Trim(strScenario) <> "" Then
            colScenarios.Add strScenario
        End If
        Set rngCurrent = rngCurrent.Offset(1)
    Wend
    Application.StatusBar = False
    Set readFeatureFromTable = colFeatures
    Exit Function
    
create_new_feature:
    Set colSingleFeature = New Collection
    colSingleFeature.Add strFeatureName, "name"
    colSingleFeature.Add strDomain, "domain"
    colSingleFeature.Add strAggregate, "aggregate"
    colSingleFeature.Add New Collection, "scenarios"
    colFeatures.Add colSingleFeature, strDomain & "-" & strAggregate & "-" & strFeatureName
    Resume Next
error_handler:
    basSystem.log_error "basTable2Features.exportFeatures"
End Function

'-------------------------------------------------------------
' Description   : writes data from collection into feature files
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub writeFeaturesToFiles(pcolFeatures As Collection)

    Dim strTargetDir As String
    Dim AppleScript As String
    Dim strFileName As String
    Dim colSingleFeature As Collection
    Dim colScenarios As Collection
    Dim lngFeatureId As Long
    Dim strFeatureName As String
    Dim strAggregateName As String
    Dim strFeatureText As String
    Dim strFullFileName As String

    On Error GoTo error_handler
    strTargetDir = basTable2Features.getTargetDir()
    lngFeatureId = 1
    For Each colSingleFeature In pcolFeatures
        strFeatureName = colSingleFeature.Item("name")
        strAggregateName = colSingleFeature.Item("aggregate")
        strFileName = basTable2Features.getFileName(lngFeatureId, strAggregateName, strFeatureName)
        strFeatureText = basTable2Features.getFeatureTextFromCollection(colSingleFeature)
        strFeatureText = Replace(strFeatureText, """", "#")
        basSystem.logd vbCr & vbLf & strFeatureText
        strFullFileName = strTargetDir & strFileName
        Application.StatusBar = "writing file " & strFileName & " to folder " & strTargetDir
        #If Mac Then
            AppleScript = "set theFeatureFile to a reference to file """ & strFullFileName & """" & vbLf & _
            "try" & vbLf & _
                "set fileRef to (open for access theFeatureFile with write permission)" & vbLf & _
            "on error errMsg number errNum" & vbLf & _
                "display dialog (""Open for Access, Error Number: "" & errNum as string) & return & errMsg" & vbLf & _
            "end try" & vbLf & _
             vbLf & _
            "set dataOut to """ & strFeatureText & """" & vbLf & _
             vbLf & _
            "try" & vbLf & _
                "write dataOut to fileRef as Çclass utf8È" & vbLf & _
            "on error errMsg number errNum" & vbLf & _
                "display dialog (""Write, Error Number: "" & errNum as string) & return & errMsg" & vbLf & _
            "end try" & vbLf & _
             vbLf & _
            "try" & vbLf & _
                "close access fileRef" & vbLf & _
            "on error errMsg number errNum" & vbLf & _
                "display dialog (""Close, Error Number: "" & errNum as string) & return & errMsg" & vbLf & _
            "end try"
            MacScript AppleScript
        #Else
            Dim fsStream As Object
            
            Set fsStream = CreateObject("ADODB.Stream")
            With fsStream
                'stream type is text/string data
                .Type = 2
                .Charset = "utf-8"
                .Open
                .WriteText strFeatureText
                .SaveToFile strFullFileName, 2
            End With
        #End If
        lngFeatureId = lngFeatureId + 1
    Next
    Application.StatusBar = False
    Exit Sub
    
error_handler:
    basSystem.log_error "basTable2Features.writeFeaturesToFiles"
End Sub
'-------------------------------------------------------------
' Description   : translates feature collection into text
' Parameter     :
' Returnvalue   : content of a feature file as string
'-------------------------------------------------------------
Private Function getFeatureTextFromCollection(pcolSingleFeature As Collection) As String
    
    Dim strFeatureName As String
    Dim strDomain As String
    Dim strAggregate As String
    Dim colScenarios As Collection
    Dim strScenario As Variant
    Dim strFeatureText As String
    
    On Error GoTo error_handler
    strDomain = pcolSingleFeature.Item("domain")
    strFeatureText = "@d-" & strDomain & vbLf
    strAggregate = pcolSingleFeature.Item("aggregate")
    strFeatureName = pcolSingleFeature.Item("name")
    strFeatureText = strFeatureText & "Feature: " & strAggregate & " - " & strFeatureName & vbLf & vbLf & vbLf
    Set colScenarios = pcolSingleFeature.Item("scenarios")
    For Each strScenario In colScenarios
        strFeatureText = strFeatureText & vbLf & "  Scenario: " & strScenario & vbLf & vbLf & vbLf & vbLf
    Next
    getFeatureTextFromCollection = strFeatureText
    Exit Function
error_handler:
    basSystem.log_error "basTable2Features.getFeatureText"
End Function

'-------------------------------------------------------------
' Description   : convert feature name into a valid file name
' Parameter     :
' Returnvalue   : feature file name as string
'-------------------------------------------------------------
Private Function getFileName(plngFeatureId, pstrAggregateName, pstrFeatureName) As String
    
    Dim strFileName As String
    Dim varSpecialChars As Variant
    Dim intReplacement As Integer
    
    On Error GoTo error_handler
    'list special chars and theire replacments like array(char, replacement, char, replacement)
    varSpecialChars = Array("""", "", "(", "#", ")", "#", " ", "-", ":", "_")
    strFileName = Trim(pstrAggregateName) & "---" & Trim(pstrFeatureName)
    For intReplacement = 0 To UBound(varSpecialChars) Step 2
        strFileName = Replace(strFileName, varSpecialChars(intReplacement), varSpecialChars(intReplacement + 1))
    Next
    strFileName = plngFeatureId & "-" & strFileName & ".feature"
    basSystem.log ("set filename to " & strFileName)
    getFileName = strFileName
    Exit Function
    
error_handler:
    basSystem.log_error "basTable2Features.getFileName"
End Function
'-------------------------------------------------------------
' Description   : ask user for the target dir where to write feature files
' Parameter     :
' Returnvalue   : target dir as string
'-------------------------------------------------------------
Private Function getTargetDir() As String

    Dim strTargetDir As Variant
    Dim AppleScript As String
    
    #If Mac Then
    
    #Else
        Dim dlgChooseFolder As FileDialog
    #End If

    On Error GoTo error_handler
    #If Mac Then
        AppleScript = "(choose folder with prompt ""choose feature folder"" default location (path to the desktop folder from user domain)) as string"
        strTargetDir = MacScript(AppleScript)
    #Else
        Set dlgChooseFolder = Application.FileDialog(msoFileDialogFolderPicker)
        With dlgChooseFolder
            .Title = "Please choose a feature folder"
            .AllowMultiSelect = False
            '.InitialFileName = strPath
            If .Show <> False Then
                strTargetDir = .SelectedItems(1) & "\"
            End If
        End With
        Set dlgChooseFolder = Nothing
    #End If
    basSystem.log ("target dir is set to " & strTargetDir)
    getTargetDir = strTargetDir
    Exit Function
    
error_handler:
    basSystem.log_error "basTable2Features.getTargetDir"
End Function
