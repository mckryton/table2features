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
    Dim colScenario As Collection
    
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
        colCurrentDataRow.Add "", cColTypeFeature & "Tags"
        colCurrentDataRow.Add "", cColTypeScenario & "Tags"
        For Each strColumnType In colAvailableColumns
            colCurrentDataRow.Remove strColumnType
            'read current feature
            If strColumnType = cColTypeDomain Then
                'replace spaces in domain names
                colCurrentDataRow.Add Trim(Replace(rngCurrent.Offset(, colDataTableSetup(strColumnType)).Text, " ", "_")), strColumnType
            ElseIf strColumnType = cColTypeFeature Then
                'replace path delimiters, because this column data will be used for file name
                #If Mac Then
                    colCurrentDataRow.Add Trim(Replace(rngCurrent.Offset(, colDataTableSetup(strColumnType)).Text, ":", " ")), strColumnType
                #Else
                    colCurrentDataRow.Add Trim(Replace(rngCurrent.Offset(, colDataTableSetup(strColumnType)).Text, "\", " ")), strColumnType
                #End If
            Else
                colCurrentDataRow.Add Trim(rngCurrent.Offset(, colDataTableSetup(strColumnType)).Text), strColumnType
            End If
            'set default name for feature, because feature name is mandatory
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
            Set colScenario = New Collection
            colScenario.Add colCurrentDataRow(cColTypeScenario), "name"
            colScenario.Add colCurrentDataRow(cColTypeScenario & "Tags"), "scenarioTags"
            colScenarios.Add colScenario, colCurrentDataRow(cColTypeScenario)
            Set colScenario = Nothing
        End If
        Set rngCurrent = rngCurrent.Offset(1)
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
    colSingleFeature.Add colCurrentDataRow(cColTypeFeature & "Tags"), "featureTags"
    colSingleFeature.Add New Collection, "scenarios"
    colFeatures.Add colSingleFeature, colCurrentDataRow(cColTypeDomain) & "-" & colCurrentDataRow(cColTypeAggregate) & "-" & colCurrentDataRow(cColTypeFeature)
    lngFeatureId = lngFeatureId + 1
    Resume Next
error_handler:
    basSystem.log_error "basTable2Features.readFeatureFromTable"
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
    strTargetDir = basTable2Features.chooseFeatureFolder()
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
        #If MAC_OFFICE_VERSION >= 15 Then
        
            Dim varScriptResult As Variant
        
            varScriptResult = AppleScriptTask("table2features.scpt", "writeFeatureToFile", strFullFileName & vbLf & strFeatureText)
            
            If LCase(CStr(varScriptResult)) = "cancel" Then
                en
            End If

        #ElseIf Mac Then
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
    Dim colScenario As Collection
    Dim strFeatureText As String
    Dim varFeatureTags As Variant
    Dim lngFeatureTag As Long
    Dim varScenarioTags As Variant
    Dim lngScenarioTag As Long
    
    On Error GoTo error_handler
    strFeatureText = ""
    strDomain = pcolSingleFeature.Item(cColTypeDomain)
    If Trim(strDomain) <> "" Then
        strFeatureText = "@d-" & strDomain
    End If
    varFeatureTags = Split(Trim(pcolSingleFeature.Item(cColTypeFeature & "Tags")))
    For lngFeatureTag = 0 To UBound(varFeatureTags)
        If Trim(varFeatureTags(lngFeatureTag)) <> "" Then
            strFeatureText = strFeatureText & " @" & Trim(varFeatureTags(lngFeatureTag))
        End If
    Next
    strFeatureText = strFeatureText & vbLf
    strAggregate = pcolSingleFeature.Item(cColTypeAggregate)
    If Trim(strAggregate) <> "" Then
        strFeatureText = strFeatureText & "Feature: " & strAggregate & " - "
    Else
        strFeatureText = strFeatureText & "Feature: "
    End If
    strFeatureName = pcolSingleFeature.Item("name")
    strFeatureText = strFeatureText & strFeatureName & vbLf & vbLf & vbLf
    Set colScenarios = pcolSingleFeature.Item("scenarios")
    For Each colScenario In colScenarios
        strFeatureText = strFeatureText & " "
        varScenarioTags = Split(Trim(colScenario.Item(cColTypeScenario & "Tags")))
        For lngScenarioTag = 0 To UBound(varScenarioTags)
            If Trim(varScenarioTags(lngScenarioTag)) <> "" Then
                strFeatureText = strFeatureText & " @" & Trim(varScenarioTags(lngScenarioTag))
            End If
        Next
        strFeatureText = strFeatureText & vbLf & _
                            "  Scenario: " & colScenario("name") & vbLf & _
                            "    GIVEN " & vbLf & _
                            "    WHEN " & vbLf & _
                            "    THEN " & vbLf & vbLf
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
    varSpecialChars = Array("""", "", "(", "#", ")", "#", " ", "-", ":", "_", "<", "#", ">", "#", "/", "-", "\", "-", "*", "#", "'", "")
    If Trim(pstrAggregateName) = "" Then
        strFileName = Trim(pstrFeatureName)
    Else
        strFileName = Trim(pstrAggregateName) & "---" & Trim(pstrFeatureName)
    End If
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
Private Function chooseFeatureFolder() As String

    Dim strFeatureFolder As Variant
    
    On Error GoTo error_handler
    #If MAC_OFFICE_VERSION >= 15 Then
       
        Dim varScriptResult As Variant
        
        varScriptResult = AppleScriptTask("table2features.scpt", "chooseFeatureFolder", "")
        strFeatureFolder = CStr(varScriptResult)

    #ElseIf Mac Then
        Dim AppleScript As String

        AppleScript = "(choose folder with prompt ""choose feature folder"" default location (path to the desktop folder from user domain)) as string"
        strFeatureFolder = MacScript(AppleScript)
    #Else
        Dim dlgChooseFolder As FileDialog

        Set dlgChooseFolder = Application.FileDialog(msoFileDialogFolderPicker)
        With dlgChooseFolder
            .Title = "Please choose a feature folder"
            .AllowMultiSelect = False
            '.InitialFileName = strPath
            If .Show <> False Then
                strFeatureFolder = .SelectedItems(1) & "\"
            End If
        End With
        Set dlgChooseFolder = Nothing
    #End If
    basSystem.log ("target dir is set to " & strFeatureFolder)
    chooseFeatureFolder = strFeatureFolder
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
        If strColumnType <> "" Then
            'save column index  - 1 because we want to use the index as an offset
            colTableSetup.Add intColumn - 1, strColumnType
            colColumnTypesFound.Add strColumnType
        End If
    Next
    'if no header detected
    If colTableSetup.Count = 0 Then
        'assume table hasn't any header
        Select Case rngDataTable.Columns.Count
            Case 4
                'assume columns are domain, aggregate, feature and scenario
                colTableSetup.Add 0, cColTypeDomain
                colTableSetup.Add 1, cColTypeAggregate
                colTableSetup.Add 2, cColTypeFeature
                colTableSetup.Add 3, cColTypeScenario
                colColumnTypesFound.Add cColTypeDomain
                colColumnTypesFound.Add cColTypeAggregate
                colColumnTypesFound.Add cColTypeFeature
                colColumnTypesFound.Add cColTypeScenario
            Case 3
                'assume columns are domain, feature and scenario
                colTableSetup.Add 0, cColTypeDomain
                colTableSetup.Add 1, cColTypeFeature
                colTableSetup.Add 2, cColTypeScenario
                colColumnTypesFound.Add cColTypeDomain
                colColumnTypesFound.Add cColTypeFeature
                colColumnTypesFound.Add cColTypeScenario
            Case 2
                'assume columns are feature and scenario
                colTableSetup.Add 0, cColTypeFeature
                colTableSetup.Add 1, cColTypeScenario
                colColumnTypesFound.Add cColTypeFeature
                colColumnTypesFound.Add cColTypeScenario
            Case 1
                'assume it's the feature only column
                colTableSetup.Add 0, cColTypeFeature
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
    'TODO: handle column names in languages other than English
    arrColumnTypes = Array(cColTypeDomain, cColTypeAggregate, cColTypeFeature, cColTypeScenario)
    For intColumnType = 0 To UBound(arrColumnTypes)
        'figure out if column is about standard data (e.g. features, scenarios)
        If Trim(LCase(pstrColumnHeader)) = arrColumnTypes(intColumnType) Or _
                Trim(LCase(pstrColumnHeader)) & "s" = arrColumnTypes(intColumnType) Then
            getDataColumnType = arrColumnTypes(intColumnType)
            Exit Function
        'figure out if the column contains feature or scenario tags (e.g. status tags)
        ElseIf arrColumnTypes(intColumnType) = cColTypeFeature Or arrColumnTypes(intColumnType) = cColTypeScenario Then
            If LCase(Left(Trim(pstrColumnHeader), Len(arrColumnTypes(intColumnType)))) = arrColumnTypes(intColumnType) _
                And (LCase(Right(Trim(pstrColumnHeader), 3)) = "tag" Or LCase(Right(Trim(pstrColumnHeader), 4)) = "tags") Then
                getDataColumnType = LCase(Left(Trim(pstrColumnHeader), Len(arrColumnTypes(intColumnType)))) & "Tags"
                Exit Function
            End If
        End If
    Next
    getDataColumnType = ""
    Exit Function
    
error_handler:
    basSystem.log_error "basTable2Features.getDataColumnType"
End Function
