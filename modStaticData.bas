Option Explicit

Private Const CLASS_NAME As String = "modStaticData"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    dictFundTeams As Object
' Returns:       ---
' Description:   Updates Static_Data Team values from Open Fondsliste funds.
'-------------------------------------------------------------------------------
Public Sub UpdateStaticDataFromFondsliste(ByVal dictFundTeams As Object)
    Const METHOD_NAME As String = "UpdateStaticDataFromFondsliste"
    Dim arrFundCodes As Variant
    Dim arrTeams As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngDataRows As Long
    Dim lngLastRow As Long
    Dim lngRow As Long
    Dim strFundNumber As String
    Dim wksStatic As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dictFundTeams Is Nothing Then Call VBA.Err.Raise(ERROR_STATIC_DATA, METHOD_NAME, "Fondsliste Team dictionary is not available.")

    Set wksStatic = ThisWorkbook.Worksheets(SHEET_STATIC_DATA)
    lngLastRow = GetStaticDataLastRow(wksStatic)

    If lngLastRow < STATIC_DATA_FIRST_DATA_ROW Then Call VBA.Err.Raise(ERROR_STATIC_DATA, METHOD_NAME, "Static_Data does not contain any data rows.")

    lngDataRows = lngLastRow - STATIC_DATA_FIRST_DATA_ROW + 1
    arrFundCodes = wksStatic.Range(wksStatic.Cells(STATIC_DATA_FIRST_DATA_ROW, 3), wksStatic.Cells(lngLastRow, 3)).Value2
    ReDim arrTeams(1 To lngDataRows, 1 To 1)

    For lngRow = 1 To lngDataRows
        strFundNumber = NormalizeLookupKey(arrFundCodes(lngRow, 1))

        If Len(strFundNumber) > 0 And dictFundTeams.Exists(strFundNumber) Then
            arrTeams(lngRow, 1) = CStr(dictFundTeams(strFundNumber))
        Else
            arrTeams(lngRow, 1) = vbNullString
        End If
    Next lngRow

    wksStatic.Range(wksStatic.Cells(STATIC_DATA_FIRST_DATA_ROW, 5), wksStatic.Cells(lngLastRow, 5)).Value2 = arrTeams

ExitPoint:
    Set wksStatic = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "row", lngRow)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    ---
' Returns:       Object
' Description:   Builds one Account to Fonds, Master, and Team dictionary.
'-------------------------------------------------------------------------------
Public Function BuildAccountMappingDictionary() As Object
    Const METHOD_NAME As String = "BuildAccountMappingDictionary"
    Dim arrData As Variant
    Dim dictAccounts As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastRow As Long
    Dim lngRow As Long
    Dim strDisplayedIdentifier As String
    Dim strFund As String
    Dim strIdentifier As String
    Dim strIdentifierContext As String
    Dim strMaster As String
    Dim strTeam As String
    Dim wksStatic As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set wksStatic = ThisWorkbook.Worksheets(SHEET_STATIC_DATA)
    lngLastRow = GetStaticDataLastRow(wksStatic)

    If lngLastRow < STATIC_DATA_FIRST_DATA_ROW Then Call VBA.Err.Raise(ERROR_STATIC_DATA, METHOD_NAME, "Static_Data does not contain any account mappings.")

    arrData = wksStatic.Range(wksStatic.Cells(STATIC_DATA_FIRST_DATA_ROW, 2), wksStatic.Cells(lngLastRow, 5)).Value2
    Set dictAccounts = CreateObject("Scripting.Dictionary")
    dictAccounts.CompareMode = vbTextCompare

    For lngRow = 1 To UBound(arrData, 1)
        If IsNumeric(arrData(lngRow, 1)) Then
            strDisplayedIdentifier = CStr(wksStatic.Cells(STATIC_DATA_FIRST_DATA_ROW + lngRow - 1, 2).Text)
        ElseIf IsError(arrData(lngRow, 1)) Then
            strDisplayedIdentifier = vbNullString
        ElseIf IsNull(arrData(lngRow, 1)) Or IsEmpty(arrData(lngRow, 1)) Then
            strDisplayedIdentifier = vbNullString
        Else
            strDisplayedIdentifier = Trim$(CStr(arrData(lngRow, 1)))
        End If

        strIdentifierContext = "Static_Data row " & CStr(STATIC_DATA_FIRST_DATA_ROW + lngRow - 1) & "."
        strIdentifier = NormalizeAccountIdentifier(arrData(lngRow, 1), strDisplayedIdentifier, ERROR_STATIC_DATA, strIdentifierContext)

        If Len(strIdentifier) > 0 Then
            strFund = GetTextValue(arrData(lngRow, 2))
            strMaster = GetTextValue(arrData(lngRow, 3))
            strTeam = GetTextValue(arrData(lngRow, 4))
            dictAccounts(strIdentifier) = Array(strFund, strMaster, strTeam)
        End If
    Next lngRow

ExitPoint:
    Set wksStatic = Nothing
    If errNumber = 0 Then Set BuildAccountMappingDictionary = dictAccounts
    Set dictAccounts = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "row", lngRow)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    arrCombined As Variant; dictAccounts As Object
' Returns:       ---
' Description:   Calculates MF, Master, VD, TD, and Team in memory.
'-------------------------------------------------------------------------------
Public Sub EnrichTransactions(ByRef arrCombined As Variant, ByVal dictAccounts As Object)
    Const METHOD_NAME As String = "EnrichTransactions"
    Dim arrMapping As Variant
    Dim dtAddedTimestamp As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngRow As Long
    Dim strAccount As String
    Dim strFund As String
    Dim strMaster As String
    Dim strTeam As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dictAccounts Is Nothing Then Call VBA.Err.Raise(ERROR_STATIC_DATA, METHOD_NAME, "Account mapping dictionary is not available.")

    dtAddedTimestamp = Now

    For lngRow = 2 To UBound(arrCombined, 1)
        strAccount = NormalizeLookupKey(arrCombined(lngRow, 6))

        If Len(strAccount) > 0 And dictAccounts.Exists(strAccount) Then
            arrMapping = dictAccounts(strAccount)
            strFund = Trim$(CStr(arrMapping(0)))
            strMaster = Trim$(CStr(arrMapping(1)))
            strTeam = Trim$(CStr(arrMapping(2)))

            If Len(strFund) = 0 Then strFund = "UNKNOWN FUND"
            If Len(strTeam) = 0 Then strTeam = "UNKNOWN TEAM"

            arrCombined(lngRow, 1) = strFund
            arrCombined(lngRow, 2) = strMaster
            arrCombined(lngRow, 5) = strTeam
        Else
            arrCombined(lngRow, 1) = "UNKNOWN FUND"
            arrCombined(lngRow, 2) = "UNKNOWN FUND"
            arrCombined(lngRow, 5) = "UNKNOWN TEAM"
        End If

        arrCombined(lngRow, 3) = arrCombined(lngRow, 8)
        arrCombined(lngRow, 4) = arrCombined(lngRow, 9)
        arrCombined(lngRow, OUTPUT_TIMESTAMP_COLUMN) = dtAddedTimestamp
    Next lngRow

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "row;account", lngRow, strAccount)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    wksStatic As Excel.Worksheet
' Returns:       Long
' Description:   Finds the last populated Static_Data row across columns B:E.
'-------------------------------------------------------------------------------
Private Function GetStaticDataLastRow(ByVal wksStatic As Excel.Worksheet) As Long
    Const METHOD_NAME As String = "GetStaticDataLastRow"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngColumnLastRow As Long
    Dim lngResult As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngColumn = 2 To 5
        lngColumnLastRow = wksStatic.Cells(wksStatic.Rows.Count, lngColumn).End(xlUp).Row
        If lngColumnLastRow > lngResult Then lngResult = lngColumnLastRow
    Next lngColumn

ExitPoint:
    If errNumber = 0 Then GetStaticDataLastRow = lngResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    varValue As Variant
' Returns:       String
' Description:   Converts and trims a lookup key using CStr.
'-------------------------------------------------------------------------------
Private Function NormalizeLookupKey(ByVal varValue As Variant) As String
    Const METHOD_NAME As String = "NormalizeLookupKey"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(varValue) Then Call VBA.Err.Raise(ERROR_STATIC_DATA, METHOD_NAME, "An Excel error value cannot be used as a lookup key.")

    If IsNull(varValue) Or IsEmpty(varValue) Then
        strResult = vbNullString
    Else
        strResult = CStr(Trim$(CStr(varValue)))
    End If

ExitPoint:
    If errNumber = 0 Then NormalizeLookupKey = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    varValue As Variant
' Returns:       String
' Description:   Converts a Static_Data value to trimmed text.
'-------------------------------------------------------------------------------
Private Function GetTextValue(ByVal varValue As Variant) As String
    Const METHOD_NAME As String = "GetTextValue"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(varValue) Then Call VBA.Err.Raise(ERROR_STATIC_DATA, METHOD_NAME, "Static_Data contains an Excel error value.")

    If IsNull(varValue) Or IsEmpty(varValue) Then
        strResult = vbNullString
    Else
        strResult = Trim$(CStr(varValue))
    End If

ExitPoint:
    If errNumber = 0 Then GetTextValue = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function

