Option Explicit

Private Const CLASS_NAME As String = "modFondsliste"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    appConfig As TAppConfig
' Returns:       String
' Description:   Resolves the latest configured Fondsliste file.
'-------------------------------------------------------------------------------
Public Function GetLatestFondslistePath(ByRef appConfig As TAppConfig) As String
    Const METHOD_NAME As String = "GetLatestFondslistePath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFolderPath As String
    Dim strResult As String
    Dim strSearchPattern As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strFolderPath = ResolveConfiguredPath(appConfig.FondslisteFolder)
    strSearchPattern = BuildFondslisteSearchPattern(appConfig.FondslistePattern, appConfig.FondslisteExtension)
    strResult = GetLatestFile(strFolderPath, strSearchPattern)

    If Len(strResult) = 0 Then Call VBA.Err.Raise(ERROR_FONDSLISTE, METHOD_NAME, "No matching Fondsliste file was found.")

ExitPoint:
    If errNumber = 0 Then GetLatestFondslistePath = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "folder;pattern;extension", appConfig.FondslisteFolder, appConfig.FondslistePattern, appConfig.FondslisteExtension)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFilePattern As String; strFileExtension As String
' Returns:       String
' Description:   Combines a configured Fondsliste prefix or pattern with extension.
'-------------------------------------------------------------------------------
Private Function BuildFondslisteSearchPattern(ByVal strFilePattern As String, ByVal strFileExtension As String) As String
    Const METHOD_NAME As String = "BuildFondslisteSearchPattern"
    Dim blnContainsWildcard As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim strExtension As String
    Dim strPattern As String
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strPattern = Trim$(strFilePattern)
    strExtension = Trim$(strFileExtension)

    If Len(strPattern) = 0 Then Call VBA.Err.Raise(ERROR_FONDSLISTE, METHOD_NAME, "Fondsliste file pattern is blank.")
    If Len(strExtension) = 0 Then Call VBA.Err.Raise(ERROR_FONDSLISTE, METHOD_NAME, "Fondsliste file extension is blank.")
    If Left$(strExtension, 1) <> "." Then strExtension = "." & strExtension

    blnContainsWildcard = InStr(1, strPattern, "*", vbBinaryCompare) > 0 Or InStr(1, strPattern, "?", vbBinaryCompare) > 0

    If blnContainsWildcard Then
        strResult = strPattern
        If Not EndsWithText(strResult, strExtension) Then strResult = strResult & strExtension
    Else
        strResult = strPattern & "*" & strExtension
    End If

ExitPoint:
    If errNumber = 0 Then BuildFondslisteSearchPattern = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "filePattern;fileExtension", strFilePattern, strFileExtension)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strValue As String; strSuffix As String
' Returns:       Boolean
' Description:   Checks whether text ends with a suffix, ignoring case.
'-------------------------------------------------------------------------------
Private Function EndsWithText(ByVal strValue As String, ByVal strSuffix As String) As Boolean
    Const METHOD_NAME As String = "EndsWithText"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Len(strSuffix) <= Len(strValue) Then blnResult = StrComp(Right$(strValue, Len(strSuffix)), strSuffix, vbTextCompare) = 0

ExitPoint:
    If errNumber = 0 Then EndsWithText = blnResult
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
' Parameters:    strFondslistePath As String
' Returns:       Object
' Description:   Builds one Open Fund Number to Team dictionary in memory.
'-------------------------------------------------------------------------------
Public Function BuildOpenFundTeamDictionary(ByVal strFondslistePath As String) As Object
    Const METHOD_NAME As String = "BuildOpenFundTeamDictionary"
    Dim arrData As Variant
    Dim blnCloseRequired As Boolean
    Dim dictTeams As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim handlerErrDescription As String
    Dim handlerErrNumber As Long
    Dim lngLastRow As Long
    Dim lngRow As Long
    Dim strFundNumber As String
    Dim strStatus As String
    Dim strTeam As String
    Dim wkbAlreadyOpen As Excel.Workbook
    Dim wkbFonds As Excel.Workbook
    Dim wksFonds As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictTeams = CreateObject("Scripting.Dictionary")
    dictTeams.CompareMode = vbTextCompare

    Set wkbAlreadyOpen = GetOpenWorkbookByFullName(strFondslistePath, True)

    If wkbAlreadyOpen Is Nothing Then
        Set wkbFonds = Application.Workbooks.Open(Filename:=strFondslistePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
        blnCloseRequired = True
    Else
        Set wkbFonds = wkbAlreadyOpen
    End If

    If wkbFonds Is Nothing Then Call VBA.Err.Raise(ERROR_FONDSLISTE, METHOD_NAME, "The Fondsliste workbook could not be accessed.")
    If wkbFonds.Worksheets.Count = 0 Then Call VBA.Err.Raise(ERROR_FONDSLISTE, METHOD_NAME, "The Fondsliste workbook does not contain a worksheet.")

    Set wksFonds = wkbFonds.Worksheets(1)
    lngLastRow = wksFonds.Cells(wksFonds.Rows.Count, 4).End(xlUp).Row

    If lngLastRow < FONDSLISTE_FIRST_DATA_ROW Then Call VBA.Err.Raise(ERROR_FONDSLISTE, METHOD_NAME, "No fund records were found in Fondsliste column D.")

    arrData = wksFonds.Range("A" & CStr(FONDSLISTE_FIRST_DATA_ROW) & ":O" & CStr(lngLastRow)).Value2

    For lngRow = 1 To UBound(arrData, 1)
        strStatus = Trim$(CStr(arrData(lngRow, 2)))
        strFundNumber = Trim$(CStr(arrData(lngRow, 4)))
        strTeam = Trim$(CStr(arrData(lngRow, 15)))

        If StrComp(strStatus, "Open", vbTextCompare) = 0 Then
            If Len(strFundNumber) > 0 Then dictTeams(strFundNumber) = strTeam
        End If
    Next lngRow

ExitPoint:
    Set wksFonds = Nothing
    Set wkbAlreadyOpen = Nothing

    If blnCloseRequired Then
        blnCloseRequired = False
        Call wkbFonds.Close(SaveChanges:=False)
    End If

    Set wkbFonds = Nothing

    If errNumber = 0 Then Set BuildOpenFundTeamDictionary = dictTeams
    Set dictTeams = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    handlerErrNumber = VBA.Err.Number
    handlerErrDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, handlerErrNumber, handlerErrDescription, "fondslistePath;row", strFondslistePath, lngRow)

    If errNumber = 0 Then
        errNumber = handlerErrNumber
        errDescription = handlerErrDescription
    End If

    GoTo ExitPoint
End Function
