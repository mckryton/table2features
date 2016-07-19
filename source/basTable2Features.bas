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
Public Sub run_exportFeatures()

    Dim colFeatures As Collection

    On Error GoTo error_handler
    Set colFeatures = basTable2Features.readFeatureFromTable()
    basTable2Features.writeFeaturesToFiles colFeatures
    Exit Sub
    
error_handler:
    basSystem.log_error "basTable2Features.exportFeatures"
End Sub
'-------------------------------------------------------------
' Description   : reads data from a table into a collection object
' Parameter     :
' Returnvalue   : collection object containing features as items
'-------------------------------------------------------------
Private Function readFeatureFromTable()

    Dim colFeatures As New Collection
    Dim colSingleFeature As Collection
    Dim colScenarios As Collection
    Dim colDataTableSetup As Collection
    Dim colAvailableColumns As Collection
    Dim strColumnType As Variant
    Dim rngCurrent As Range
    Dim colCurrentDataRow As Collection
    Dim lngFeatureId As Long
    
    'On Error GoTo error_handler
    lngFeatureId = 1
    Set colDataTableSetup = getDataTableSetup()
    Set rngCurrent = colDataTableSetup("firstItem")
    Set colAvailableColumns = colDataTableSetup("ColTypes")
    Application.StatusBar = "reading table"
    While Trim(rngCurrent.Text) <> ""
        Set colCurrentDataRow = New Collection
        colCurrentDataRow.Add "", cColTypeDomain
        colCurrentDataRow.Add "", cColTypeAggregate
        colCurrentDataRow.Add "", cColTypeFeature
        colCurrentDataRow.Add "", cColTypeScenario
        For Each strColumnType In colAvailableColumns
            colCurrentDataRow.Remove strColumnType
            'read current feature
            If strColumnType = cColTypeDomain Then
                colCurrentDataRow.Add Trim(Replace(rngCurrent.Offset(, colDataTableSetup(strColumnType)).Text, " ", "_")), strColumnType
            Else
                'replace path delimiters, because this column data will be used for file name
                #If Mac Then
                    colCurrentDataRow.Add Trim(Replace(rngCurrent.Offset(, colDataTableSetup(strColumnType)).Text, ":", " ")), strColumnType
                #Else
                    colCurrentDataRow.Add Trim(Replace(rngCurrent.Offset(, colDataTableSetup(strColumnType)).Text, "\", " ")), strColumnType
                #End If
            End If
            If strColumnType = cColTypeFeature And colCurrentDataRow(strColumnType) = "" Then
                colCurrentDataRow.Remove strColumnType
                colCurrentDataRow.Add "undefined_" & lngFeatureId, strColumnType
            End If
        Next
        basSystem.log "read feature: " & colCurrentDataRow(cColTypeDomain) & " - " & colCurrentDataRow(cColTypeAggregate) & " - " & colCurrentDataRow(cColTypeFeature)
        'try to get the existing feature
        On Error GoTo create_new_feature
        'if not found create a new one
        Set colSingleFeature = colFeatures(colCurrentDataRow(cColTypeDomain) & "-" & colCurrentDataRow(cColTypeAggregate) & "-" & colCurrentDataRow(cColTypeFeature))
        On Error GoTo error_handler
        Set colScenarios = colSingleFeature.Item("scenarios")
        If Trim(colCurrentDataRow(cColTypeScenario)) <> "" Then
            colScenarios.Add colCurrentDataRow(cColTypeScenario)
        End If
        Set rngCurrent = rngCurrent.Offset(1)
        lngFeatureId = lngFeatureId + 1
        Set colCurrentDataRow = Nothing
    Wend
    Application.StatusBar = False
    Set readFeatureFromTable = colFeatures
    Exit Function
    
create_new_feature:
    Set colSingleFeature = New Collection
    colSingleFeature.Add lngFeatureId, "featureId"
    colSingleFeature.Add colCurrentDataRow(cColTypeFeature), "name"
    colSingleFeature.Add colCurrentDataRow(cColTypeDomain), "domain"
    colSingleFeature.Add colCurrentDataRow(cColTypeAggregate), "aggregate"
    colSingleFeature.Add New Collection, "scenarios"
    colFeatures.Add colSingleFeature, colCurrentDataRow(cColTypeDomain) & "-" & colCurrentDataRow(cColTypeAggregate) & "-" & colCurrentDataRow(cColTypeFeature)
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
    For Each colSingleFeature In pcolFeatures
        lngFeatureId = colSingleFeature.Item("featureId")
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
    basSystem.log_error "basTable2Features.getTargetDir"
End Function
'-------------------------------------------------------------
' Description   : detect column names and first data row or just
'                   just assume a default order if table header is missing
' Parameter     :
' Returnvalue   : collection containing table setup
'-------------------------------------------------------------
Private Function getDataTableSetup() As Collection

    Dim colTableSetup As New Collection
    Dim colColumnTypesFound As New Collection
    Dim rngFirstDataItem As Range
    Dim rngDataTable As Range
    Dim rngSelection As Range
    Dim intColumn As Integer
    Dim strColumnType As String

    'stop script if no range selected
    On Error GoTo stop_script
    Set rngSelection = Selection
    On Error GoTo error_handler
    Set rngDataTable = rngSelection.CurrentRegion
    For intColumn = 1 To rngDataTable.Columns.Count
        strColumnType = getDataColumnType(rngDataTable.Cells(1, intColumn).Text)
        colColumnTypesFound.Add strColumnType
        If strColumnType <> "" Then
            'save column index  - 1 because we want to use the index as an offset
            colTableSetup.Add intColumn - 1, strColumnType
        End If
    Next
    'if no header detected
    If colTableSetup.Count = 0 Then
        'assume table hasn't any header
        Select Case rngDataTable.Columns.Count
            Case 4
                'assume columns are domain, aggregate, feature and scenario
                colTableSetup.Add 1, cColTypeDomain
                colTableSetup.Add 2, cColTypeAggregate
                colTableSetup.Add 3, cColTypeFeature
                colTableSetup.Add 4, cColTypeScenario
                colColumnTypesFound.Add cColTypeDomain
                colColumnTypesFound.Add cColTypeAggregate
                colColumnTypesFound.Add cColTypeFeature
                colColumnTypesFound.Add cColTypeScenario
            Case 3
                'assume columns are domain, feature and scenario
                colTableSetup.Add 1, cColTypeDomain
                colTableSetup.Add 2, cColTypeFeature
                colTableSetup.Add 3, cColTypeScenario
                colColumnTypesFound.Add cColTypeDomain
                colColumnTypesFound.Add cColTypeFeature
                colColumnTypesFound.Add cColTypeScenario
            Case 2
                'assume columns are feature and scenario
                colTableSetup.Add 1, cColTypeFeature
                colTableSetup.Add 2, cColTypeScenario
                colColumnTypesFound.Add cColTypeFeature
                colColumnTypesFound.Add cColTypeScenario
            Case 1
                'assume it's the feature only column
                colTableSetup.Add 1, cColTypeFeature
                colColumnTypesFound.Add cColTypeFeature
            Case Else
                MsgBox "Sorry, found " & rngDataTable.Columns.Count & " columns in your table but no header. Expect 4 to 2, don't know how to map them." & vbCrLf & _
                    "Please use domain, aggregate, feature or scenario as table header!"
                End
        End Select
        'table starts in first row
        colTableSetup.Add rngDataTable.Cells(1, 1), "firstItem"
    Else
        'table starts in second row
        colTableSetup.Add rngDataTable.Cells(2, 1), "firstItem"
    End If
    'feature column is mandatory for creating feature files
    On Error GoTo missing_columns
    intColumn = colTableSetup(cColTypeFeature)
    colTableSetup.Add colColumnTypesFound, "ColTypes"
    Set getDataTableSetup = colTableSetup
    Exit Function

missing_columns:
    MsgBox "Sorry, could not find a feature column in your selected table!"
    End
stop_script:
    MsgBox "Please set the cursor into the table containing your feature data!"
    End
error_handler:
    basSystem.log_error "basTable2Features.getDataTableSetup"
End Function
'-------------------------------------------------------------
' Description   : try to recognize the columntype from it's heading
' Parameter     :
' Returnvalue   : domain, aggregate, feature or scenario
'-------------------------------------------------------------
Private Function getDataColumnType(pstrColumnHeader As String) As String

    Dim arrColumnTypes As Variant
    Dim intColumnType As Integer
        
    On Error GoTo error_handler
    arrColumnTypes = Array(cColTypeDomain, cColTypeAggregate, cColTypeFeature, cColTypeScenario)
    For intColumnType = 0 To UBound(arrColumnTypes)
        If Trim(LCase(pstrColumnHeader)) = arrColumnTypes(intColumnType) Or _
                Trim(LCase(pstrColumnHeader)) & "s" = arrColumnTypes(intColumnType) Then
            getDataColumnType = arrColumnTypes(intColumnType)
            Exit Function
        End If
    Next
    getDataColumnType = ""
    
error_handler:
    basSystem.log_error "basTable2Features.getDataTableSetup"
End Function
