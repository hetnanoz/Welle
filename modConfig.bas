Option Explicit

Private Const CLASS_NAME As String = "modConfig"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    appConfig As TAppConfig
' Returns:       ---
' Description:   Reads and validates all named configuration cells.
'-------------------------------------------------------------------------------
Public Sub LoadConfiguration(ByRef appConfig As TAppConfig)
    Const METHOD_NAME As String = "LoadConfiguration"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    appConfig.SaveLocal = ReadBooleanFlag(NAME_FLAG_SAVE_LOCAL)
    appConfig.SaveSharePoint = ReadBooleanFlag(NAME_FLAG_SAVE_SHAREPOINT)
    appConfig.UpdateHistorical = ReadBooleanFlag(NAME_FLAG_UPDATE_HISTORICAL)
    appConfig.OutlookMailbox = ReadRequiredText(NAME_OUTLOOK_MAILBOX)
    appConfig.OutlookSourceFolder = ReadRequiredText(NAME_OUTLOOK_SOURCE_FOLDER)
    appConfig.OutlookArchiveFolder = ReadRequiredText(NAME_OUTLOOK_ARCHIVE_FOLDER)
    appConfig.FondslisteFolder = ReadRequiredText(NAME_FONDSLISTE_FOLDER)
    appConfig.FondslistePattern = ReadRequiredText(NAME_FONDSLISTE_PATTERN)
    appConfig.FondslisteExtension = ReadRequiredText(NAME_FONDSLISTE_EXTENSION)
    appConfig.SupportedTeams = ReadRequiredText(NAME_SUPPORTED_TEAMS)

    If appConfig.SaveLocal Then appConfig.OutputLocalBase = ReadRequiredText(NAME_OUTPUT_LOCAL_BASE)
    If appConfig.SaveSharePoint Then appConfig.OutputSharePointBase = ReadRequiredText(NAME_OUTPUT_SHAREPOINT_BASE)

    Call ValidateConfiguration(appConfig)

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strName As String
' Returns:       Excel.Range
' Description:   Returns a one-cell named range located on Mapping.
'-------------------------------------------------------------------------------
Private Function GetNamedCellRange(ByVal strName As String) As Excel.Range
    Const METHOD_NAME As String = "GetNamedCellRange"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objName As Excel.Name
    Dim rngCandidate As Excel.Range
    Dim rngResult As Excel.Range
    Dim strSimpleName As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each objName In ThisWorkbook.Names
        strSimpleName = GetUnqualifiedName(objName.Name)

        If StrComp(strSimpleName, strName, vbTextCompare) = 0 Then
            Set rngCandidate = objName.RefersToRange

            If StrComp(CStr(rngCandidate.Parent.Name), SHEET_MAPPING, vbTextCompare) = 0 Then
                Set rngResult = rngCandidate
                Exit For
            End If
        End If
    Next objName

    If rngResult Is Nothing Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "Named cell was not found on Mapping: " & strName)
    If rngResult.Cells.CountLarge <> 1 Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "Named range must refer to exactly one cell: " & strName)

ExitPoint:
    Set rngCandidate = Nothing
    Set objName = Nothing
    If errNumber = 0 Then Set GetNamedCellRange = rngResult
    Set rngResult = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "name", strName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strQualifiedName As String
' Returns:       String
' Description:   Removes workbook or worksheet qualification from a name.
'-------------------------------------------------------------------------------
Private Function GetUnqualifiedName(ByVal strQualifiedName As String) As String
    Const METHOD_NAME As String = "GetUnqualifiedName"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPosition As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngPosition = InStrRev(strQualifiedName, "!")

    If lngPosition > 0 Then
        strResult = Mid$(strQualifiedName, lngPosition + 1)
    Else
        strResult = strQualifiedName
    End If

    strResult = Replace$(strResult, "'", vbNullString)

ExitPoint:
    If errNumber = 0 Then GetUnqualifiedName = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "qualifiedName", strQualifiedName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strName As String
' Returns:       String
' Description:   Reads a required text value from a named Mapping cell.
'-------------------------------------------------------------------------------
Private Function ReadRequiredText(ByVal strName As String) As String
    Const METHOD_NAME As String = "ReadRequiredText"
    Dim errDescription As String
    Dim errNumber As Long
    Dim rngValue As Excel.Range
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set rngValue = GetNamedCellRange(strName)
    strResult = Trim$(CStr(rngValue.Value2))

    If Len(strResult) = 0 Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "Required configuration value is blank: " & strName)

ExitPoint:
    Set rngValue = Nothing
    If errNumber = 0 Then ReadRequiredText = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "name", strName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strName As String
' Returns:       Boolean
' Description:   Reads a Boolean flag from a named Mapping cell.
'-------------------------------------------------------------------------------
Private Function ReadBooleanFlag(ByVal strName As String) As Boolean
    Const METHOD_NAME As String = "ReadBooleanFlag"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim rngValue As Excel.Range
    Dim strValue As String
    Dim varValue As Variant

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set rngValue = GetNamedCellRange(strName)
    varValue = rngValue.Value2

    If VarType(varValue) = vbBoolean Then
        blnResult = CBool(varValue)
        GoTo ExitPoint
    End If

    If IsNumeric(varValue) Then
        blnResult = CDbl(varValue) <> 0
        GoTo ExitPoint
    End If

    strValue = UCase$(Trim$(CStr(varValue)))
    strValue = Replace$(strValue, ChrW(321), "L")
    strValue = Replace$(strValue, ChrW(322), "L")

    Select Case strValue
        Case "TRUE", "PRAWDA", "YES", "Y", "1"
            blnResult = True
        Case "FALSE", "FALSZ", "NO", "N", "0"
            blnResult = False
        Case Else
            Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "Invalid Boolean value in named cell: " & strName)
    End Select

ExitPoint:
    Set rngValue = Nothing
    If errNumber = 0 Then ReadBooleanFlag = blnResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "name", strName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    appConfig As TAppConfig
' Returns:       ---
' Description:   Validates cross-field configuration requirements.
'-------------------------------------------------------------------------------
Private Sub ValidateConfiguration(ByRef appConfig As TAppConfig)
    Const METHOD_NAME As String = "ValidateConfiguration"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not appConfig.SaveLocal And Not appConfig.SaveSharePoint Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "At least one output flag must be TRUE.")
    If appConfig.SaveLocal And Len(Trim$(appConfig.OutputLocalBase)) = 0 Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "Local output path is required when FLAG_SAVE_LOCAL is TRUE.")
    If appConfig.SaveSharePoint And Len(Trim$(appConfig.OutputSharePointBase)) = 0 Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "SharePoint output path is required when FLAG_SAVE_SHAREPOINT is TRUE.")
    If StrComp(appConfig.OutlookSourceFolder, appConfig.OutlookArchiveFolder, vbTextCompare) = 0 Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "Outlook source and archive folders must be different.")

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Sub
