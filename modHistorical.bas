Option Explicit

Private Const CLASS_NAME As String = "modHistorical"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    arrCombined As Variant; colTeams As Collection; colDestinationRoots As Collection
' Returns:       ---
' Description:   Updates combined, unknown-fund, and per-team historical master workbooks.
'-------------------------------------------------------------------------------
Public Sub UpdateHistoricalFiles(ByRef arrCombined As Variant, ByVal colTeams As Collection, ByVal colDestinationRoots As Collection)
    Const METHOD_NAME As String = "UpdateHistoricalFiles"
    Dim arrTeamData As Variant
    Dim arrUnknownData As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngRoot As Long
    Dim lngTeam As Long
    Dim strHistoricalPath As String
    Dim strRootPath As String
    Dim strTeam As String
    Dim strTeamFilePart As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If UBound(arrCombined, 1) > 1 Then
        For lngRoot = 1 To colDestinationRoots.Count
            strRootPath = CStr(colDestinationRoots(lngRoot))
            strHistoricalPath = CombinePath(strRootPath, HISTORICAL_COMBINED_FILE)
            Call UpdateHistoricalWorkbook(strHistoricalPath, arrCombined)
        Next lngRoot
    End If

    arrUnknownData = FilterUnknownFundTransactions(arrCombined)

    If UBound(arrUnknownData, 1) > 1 Then
        For lngRoot = 1 To colDestinationRoots.Count
            strRootPath = CStr(colDestinationRoots(lngRoot))
            strHistoricalPath = CombinePath(strRootPath, HISTORICAL_UNKNOWN_FILE)
            Call UpdateHistoricalWorkbook(strHistoricalPath, arrUnknownData)
        Next lngRoot
    End If

    arrUnknownData = Empty

    For lngTeam = 1 To colTeams.Count
        strTeam = CStr(colTeams(lngTeam))
        strTeamFilePart = SanitizeFileNamePart(strTeam)
        arrTeamData = FilterTransactionsByTeam(arrCombined, strTeam)

        If UBound(arrTeamData, 1) > 1 Then
            For lngRoot = 1 To colDestinationRoots.Count
                strRootPath = CStr(colDestinationRoots(lngRoot))
                strHistoricalPath = CombinePath(strRootPath, HISTORICAL_TEAM_PREFIX & strTeamFilePart & ".xlsx")
                Call UpdateHistoricalWorkbook(strHistoricalPath, arrTeamData)
            Next lngRoot
        End If

        arrTeamData = Empty
    Next lngTeam

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "team;historicalPath", strTeam, strHistoricalPath)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strHistoricalPath As String; arrCurrent As Variant
' Returns:       ---
' Description:   Appends only unseen records to one historical workbook.
'-------------------------------------------------------------------------------
Private Sub UpdateHistoricalWorkbook(ByVal strHistoricalPath As String, ByRef arrCurrent As Variant)
    Const METHOD_NAME As String = "UpdateHistoricalWorkbook"
    Dim arrExisting As Variant
    Dim arrInclude() As Boolean
    Dim arrNewRows As Variant
    Dim blnCloseRequired As Boolean
    Dim blnHeaderUpgraded As Boolean
    Dim blnHttpPath As Boolean
    Dim blnNewWorkbook As Boolean
    Dim dictKeys As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim handlerErrDescription As String
    Dim handlerErrNumber As Long
    Dim lngColumn As Long
    Dim lngCurrentLastRow As Long
    Dim lngExistingLastRow As Long
    Dim lngNewCount As Long
    Dim lngOutputRow As Long
    Dim lngRow As Long
    Dim strKey As String
    Dim wkbAlreadyOpen As Excel.Workbook
    Dim wkbHistory As Excel.Workbook
    Dim wksHistory As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngCurrentLastRow = UBound(arrCurrent, 1)
    blnHttpPath = IsHttpPath(strHistoricalPath)
    Set wkbAlreadyOpen = GetOpenWorkbookByFullName(strHistoricalPath, False)

    If Not wkbAlreadyOpen Is Nothing Then
        If wkbAlreadyOpen.ReadOnly Then Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "Historical workbook is already open as read-only: " & strHistoricalPath)
        If Not wkbAlreadyOpen.Saved Then Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "Historical workbook has unsaved user changes. Save or close it before running the macro: " & strHistoricalPath)
        Set wkbHistory = wkbAlreadyOpen
    ElseIf blnHttpPath Then
        Set wkbHistory = Application.Workbooks.Open(Filename:=strHistoricalPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
        blnCloseRequired = True
    ElseIf FileExists(strHistoricalPath) Then
        Set wkbHistory = Application.Workbooks.Open(Filename:=strHistoricalPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
        blnCloseRequired = True
    Else
        Set wkbHistory = Application.Workbooks.Add(xlWBATWorksheet)
        blnCloseRequired = True
        blnNewWorkbook = True
    End If

    If wkbHistory Is Nothing Then Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "Historical workbook could not be accessed.")
    If wkbHistory.ReadOnly Then Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "Historical workbook is read-only: " & strHistoricalPath)

    Set wksHistory = GetHistoricalWorksheet(wkbHistory, blnNewWorkbook)

    If blnNewWorkbook Then
        Call InitializeHistoricalHeader(wksHistory)
    Else
        Call ValidateHistoricalHeader(wksHistory, strHistoricalPath, blnHeaderUpgraded)
    End If

    wksHistory.Columns("C:D").NumberFormat = "dd.mm.yyyy"
    wksHistory.Columns("H:I").NumberFormat = "dd.mm.yyyy"
    wksHistory.Columns("V:V").NumberFormat = OUTPUT_TIMESTAMP_NUMBER_FORMAT

    lngExistingLastRow = wksHistory.Cells(wksHistory.Rows.Count, 6).End(xlUp).Row
    If lngExistingLastRow < 1 Then lngExistingLastRow = 1

    Set dictKeys = CreateObject("Scripting.Dictionary")
    dictKeys.CompareMode = vbBinaryCompare

    If lngExistingLastRow >= 2 Then
        arrExisting = wksHistory.Range(wksHistory.Cells(1, 1), wksHistory.Cells(lngExistingLastRow, OUTPUT_COLUMN_COUNT)).Value2

        For lngRow = 2 To UBound(arrExisting, 1)
            strKey = BuildTransactionKey(arrExisting, lngRow)
            If Not dictKeys.Exists(strKey) Then dictKeys.Add strKey, True
        Next lngRow
    End If

    If lngCurrentLastRow >= 2 Then
        ReDim arrInclude(2 To lngCurrentLastRow)

        For lngRow = 2 To lngCurrentLastRow
            strKey = BuildTransactionKey(arrCurrent, lngRow)

            If Not dictKeys.Exists(strKey) Then
                dictKeys.Add strKey, True
                arrInclude(lngRow) = True
                lngNewCount = lngNewCount + 1
            End If
        Next lngRow
    End If

    If lngNewCount > 0 Then
        ReDim arrNewRows(1 To lngNewCount, 1 To OUTPUT_COLUMN_COUNT)
        lngOutputRow = 0

        For lngRow = 2 To lngCurrentLastRow
            If arrInclude(lngRow) Then
                lngOutputRow = lngOutputRow + 1

                For lngColumn = 1 To OUTPUT_COLUMN_COUNT
                    arrNewRows(lngOutputRow, lngColumn) = arrCurrent(lngRow, lngColumn)
                Next lngColumn
            End If
        Next lngRow

        wksHistory.Range(wksHistory.Cells(lngExistingLastRow + 1, 1), wksHistory.Cells(lngExistingLastRow + lngNewCount, OUTPUT_COLUMN_COUNT)).Value2 = arrNewRows
    End If

    If blnNewWorkbook Or lngNewCount > 0 Or blnHeaderUpgraded Then
        lngExistingLastRow = lngExistingLastRow + lngNewCount
        Call AutoFitColumnsWithPadding(wksHistory.Range(wksHistory.Cells(1, 1), wksHistory.Cells(lngExistingLastRow, OUTPUT_COLUMN_COUNT)).Columns, 2)
    End If

    If blnNewWorkbook Then
        Call ApplyWorkbookSensitivityLabel(wkbHistory)
        Call wkbHistory.SaveAs(Filename:=strHistoricalPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False, AddToMru:=False, Local:=True)
    ElseIf lngNewCount > 0 Or blnHeaderUpgraded Then
        Call ApplyWorkbookSensitivityLabel(wkbHistory)
        Call wkbHistory.Save
    End If

ExitPoint:
    Set dictKeys = Nothing
    Set wksHistory = Nothing
    Set wkbAlreadyOpen = Nothing

    If blnCloseRequired Then
        blnCloseRequired = False
        Call wkbHistory.Close(SaveChanges:=False)
    End If

    Set wkbHistory = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    handlerErrNumber = VBA.Err.Number
    handlerErrDescription = VBA.Err.Description

    If blnHttpPath And wkbHistory Is Nothing Then
        handlerErrDescription = "HTTPS historical workbook could not be opened. Upload/create the historical master in the configured SharePoint root before enabling historical updates, and confirm edit permission. Path: " & strHistoricalPath & ". Original error: " & handlerErrDescription
    End If

    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, handlerErrNumber, handlerErrDescription, "historicalPath;currentRows;newRows", strHistoricalPath, lngCurrentLastRow, lngNewCount)

    If errNumber = 0 Then
        errNumber = handlerErrNumber
        errDescription = handlerErrDescription
    End If

    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    wkbHistory As Excel.Workbook; blnNewWorkbook As Boolean
' Returns:       Excel.Worksheet
' Description:   Resolves the historical transaction worksheet.
'-------------------------------------------------------------------------------
Private Function GetHistoricalWorksheet(ByVal wkbHistory As Excel.Workbook, ByVal blnNewWorkbook As Boolean) As Excel.Worksheet
    Const METHOD_NAME As String = "GetHistoricalWorksheet"
    Dim errDescription As String
    Dim errNumber As Long
    Dim wksCandidate As Excel.Worksheet
    Dim wksResult As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If wkbHistory.Worksheets.Count = 0 Then Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "Historical workbook does not contain a worksheet.")

    If blnNewWorkbook Then
        Set wksResult = wkbHistory.Worksheets(1)
        wksResult.Name = SHEET_TRANSACTIONS
    Else
        For Each wksCandidate In wkbHistory.Worksheets
            If StrComp(wksCandidate.Name, SHEET_TRANSACTIONS, vbTextCompare) = 0 Then
                Set wksResult = wksCandidate
                Exit For
            End If
        Next wksCandidate

        If wksResult Is Nothing Then Set wksResult = wkbHistory.Worksheets(1)
    End If

ExitPoint:
    Set wksCandidate = Nothing
    If errNumber = 0 Then Set GetHistoricalWorksheet = wksResult
    Set wksResult = Nothing
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
' Parameters:    wksHistory As Excel.Worksheet
' Returns:       ---
' Description:   Initializes a new historical workbook header.
'-------------------------------------------------------------------------------
Private Sub InitializeHistoricalHeader(ByVal wksHistory As Excel.Worksheet)
    Const METHOD_NAME As String = "InitializeHistoricalHeader"
    Dim arrHeaders As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrHeaders = GetOutputHeaders()

    For lngColumn = 1 To OUTPUT_COLUMN_COUNT
        wksHistory.Cells(1, lngColumn).Value2 = arrHeaders(lngColumn - 1)
    Next lngColumn

    wksHistory.Rows(1).Font.Bold = True
    Call wksHistory.Range(wksHistory.Cells(1, 1), wksHistory.Cells(1, OUTPUT_COLUMN_COUNT)).AutoFilter
    wksHistory.Columns("C:D").NumberFormat = "dd.mm.yyyy"
    wksHistory.Columns("H:I").NumberFormat = "dd.mm.yyyy"
    wksHistory.Columns("R:S").NumberFormat = "#,##0.00"
    wksHistory.Columns("V:V").NumberFormat = OUTPUT_TIMESTAMP_NUMBER_FORMAT

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
' Parameters:    wksHistory As Excel.Worksheet; strHistoricalPath As String; blnHeaderUpgraded As Boolean
' Returns:       ---
' Description:   Validates the historical header and upgrades legacy A:U files with the timestamp column.
'-------------------------------------------------------------------------------
Private Sub ValidateHistoricalHeader(ByVal wksHistory As Excel.Worksheet, ByVal strHistoricalPath As String, ByRef blnHeaderUpgraded As Boolean)
    Const METHOD_NAME As String = "ValidateHistoricalHeader"
    Dim arrHeaders As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim strActual As String
    Dim strExpected As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrHeaders = GetOutputHeaders()
    blnHeaderUpgraded = False

    For lngColumn = 1 To OUTPUT_TIMESTAMP_COLUMN - 1
        strActual = Trim$(CStr(wksHistory.Cells(1, lngColumn).Value2))
        strExpected = CStr(arrHeaders(lngColumn - 1))

        If StrComp(strActual, strExpected, vbTextCompare) <> 0 Then Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "Historical header mismatch in column " & CStr(lngColumn) & ". Expected '" & strExpected & "' and found '" & strActual & "'.")
    Next lngColumn

    strActual = Trim$(CStr(wksHistory.Cells(1, OUTPUT_TIMESTAMP_COLUMN).Value2))
    strExpected = CStr(arrHeaders(OUTPUT_TIMESTAMP_COLUMN - 1))

    If Len(strActual) = 0 Then
        wksHistory.Cells(1, OUTPUT_TIMESTAMP_COLUMN).Value2 = strExpected
        blnHeaderUpgraded = True
    ElseIf StrComp(strActual, strExpected, vbTextCompare) <> 0 Then
        Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "Historical header mismatch in column " & CStr(OUTPUT_TIMESTAMP_COLUMN) & ". Expected '" & strExpected & "' and found '" & strActual & "'.")
    End If

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "historicalPath;column", strHistoricalPath, lngColumn)
    GoTo ExitPoint
End Sub


'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-02
' Parameters:    rngColumns As Excel.Range; dblPadding As Double
' Returns:       ---
' Description:   AutoFits historical output columns and adds extra width.
'-------------------------------------------------------------------------------
Private Sub AutoFitColumnsWithPadding(ByVal rngColumns As Excel.Range, ByVal dblPadding As Double)
    Const METHOD_NAME As String = "AutoFitColumnsWithPadding"
    Const MAX_EXCEL_COLUMN_WIDTH As Double = 255
    Dim errDescription As String
    Dim errNumber As Long
    Dim rngColumn As Excel.Range
    Dim dblNewWidth As Double

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Call rngColumns.AutoFit

    For Each rngColumn In rngColumns.Columns
        dblNewWidth = rngColumn.ColumnWidth + dblPadding

        If dblNewWidth > MAX_EXCEL_COLUMN_WIDTH Then
            dblNewWidth = MAX_EXCEL_COLUMN_WIDTH
        End If

        rngColumn.ColumnWidth = dblNewWidth
    Next rngColumn

ExitPoint:
    Set rngColumn = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "padding", dblPadding)
    GoTo ExitPoint
End Sub

