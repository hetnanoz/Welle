Option Explicit

Private Const CLASS_NAME As String = "modReports"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strTeamList As String
' Returns:       Collection
' Description:   Parses and deduplicates a semicolon-separated team list.
'-------------------------------------------------------------------------------
Public Function ParseSupportedTeams(ByVal strTeamList As String) As Collection
    Const METHOD_NAME As String = "ParseSupportedTeams"
    Dim arrTeams As Variant
    Dim colResult As Collection
    Dim dictTeams As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngTeam As Long
    Dim strTeam As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set colResult = New Collection
    Set dictTeams = CreateObject("Scripting.Dictionary")
    dictTeams.CompareMode = vbTextCompare
    arrTeams = Split(strTeamList, ";")

    For lngTeam = LBound(arrTeams) To UBound(arrTeams)
        strTeam = Trim$(CStr(arrTeams(lngTeam)))

        If Len(strTeam) > 0 Then
            If Not dictTeams.Exists(strTeam) Then
                dictTeams.Add strTeam, True
                colResult.Add strTeam
            End If
        End If
    Next lngTeam

    If colResult.Count = 0 Then Call VBA.Err.Raise(ERROR_CONFIGURATION, METHOD_NAME, "SUPPORTED_TEAMS does not contain any team values.")

ExitPoint:
    Set dictTeams = Nothing
    If errNumber = 0 Then Set ParseSupportedTeams = colResult
    Set colResult = Nothing
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
' Parameters:    appConfig As TAppConfig
' Returns:       Collection
' Description:   Resolves and deduplicates enabled output root folders.
'-------------------------------------------------------------------------------
Public Function BuildDestinationRoots(ByRef appConfig As TAppConfig) As Collection
    Const METHOD_NAME As String = "BuildDestinationRoots"
    Dim colResult As Collection
    Dim dictRoots As Object
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set colResult = New Collection
    Set dictRoots = CreateObject("Scripting.Dictionary")
    dictRoots.CompareMode = vbTextCompare

    If appConfig.SaveLocal Then Call AddDestinationRoot(appConfig.OutputLocalBase, colResult, dictRoots)
    If appConfig.SaveSharePoint Then Call AddDestinationRoot(appConfig.OutputSharePointBase, colResult, dictRoots)

    If colResult.Count = 0 Then Call VBA.Err.Raise(ERROR_OUTPUT, METHOD_NAME, "No enabled output destination could be resolved.")

ExitPoint:
    Set dictRoots = Nothing
    If errNumber = 0 Then Set BuildDestinationRoots = colResult
    Set colResult = Nothing
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
' Parameters:    strConfiguredPath As String; colRoots As Collection; dictRoots As Object
' Returns:       ---
' Description:   Adds one normalized and available output root.
'-------------------------------------------------------------------------------
Private Sub AddDestinationRoot(ByVal strConfiguredPath As String, ByVal colRoots As Collection, ByVal dictRoots As Object)
    Const METHOD_NAME As String = "AddDestinationRoot"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strKey As String
    Dim strResolvedPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResolvedPath = ResolveConfiguredPath(strConfiguredPath)
    Call EnsureFolderExists(strResolvedPath)
    strKey = NormalizeWorkbookPath(strResolvedPath)

    If Not dictRoots.Exists(strKey) Then
        dictRoots.Add strKey, True
        colRoots.Add strResolvedPath
    End If

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "configuredPath", strConfiguredPath)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    arrCombined As Variant; colTeams As Collection; colDestinationRoots As Collection; dtRunTimestamp As Date; strWorkspace As String
' Returns:       ---
' Description:   Creates and publishes the combined and per-team daily reports.
'-------------------------------------------------------------------------------
Public Sub CreateDailyReports(ByRef arrCombined As Variant, ByVal colTeams As Collection, ByVal colDestinationRoots As Collection, ByVal dtRunTimestamp As Date, ByVal strWorkspace As String)
    Const METHOD_NAME As String = "CreateDailyReports"
    Dim arrTeamData As Variant
    Dim blnEmpty As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngTeam As Long
    Dim strDateToken As String
    Dim strFileName As String
    Dim strTeam As String
    Dim strTeamFilePart As String
    Dim strTimeToken As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strDateToken = Format$(dtRunTimestamp, "yyyymmdd")
    strTimeToken = Format$(dtRunTimestamp, "hhnnss")

    strFileName = REPORT_COMBINED_PREFIX & strDateToken & "_" & strTimeToken & ".xlsx"
    Call PublishReport(arrCombined, strFileName, colDestinationRoots, DateValue(dtRunTimestamp), strWorkspace)

    For lngTeam = 1 To colTeams.Count
        strTeam = CStr(colTeams(lngTeam))
        strTeamFilePart = SanitizeFileNamePart(strTeam)
        arrTeamData = FilterTransactionsByTeam(arrCombined, strTeam)
        blnEmpty = UBound(arrTeamData, 1) = 1

        strFileName = REPORT_TEAM_PREFIX & strTeamFilePart & "_" & strDateToken & "_" & strTimeToken
        If blnEmpty Then strFileName = strFileName & "_EMPTY"
        strFileName = strFileName & ".xlsx"

        Call PublishReport(arrTeamData, strFileName, colDestinationRoots, DateValue(dtRunTimestamp), strWorkspace)
        arrTeamData = Empty
    Next lngTeam

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "team;workspace", strTeam, strWorkspace)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    arrReport As Variant; strFileName As String; colDestinationRoots As Collection; dtReportDate As Date; strWorkspace As String
' Returns:       ---
' Description:   Creates one staged workbook and copies it to all destinations.
'-------------------------------------------------------------------------------
Private Sub PublishReport(ByRef arrReport As Variant, ByVal strFileName As String, ByVal colDestinationRoots As Collection, ByVal dtReportDate As Date, ByVal strWorkspace As String)
    Const METHOD_NAME As String = "PublishReport"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngRoot As Long
    Dim strDailyFolder As String
    Dim strStagePath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strStagePath = CombinePath(strWorkspace, strFileName)
    Call DeleteFileIfExists(strStagePath)
    Call CreateReportWorkbook(arrReport, strStagePath)

    For lngRoot = 1 To colDestinationRoots.Count
        strDailyFolder = BuildDatedFolderPath(CStr(colDestinationRoots(lngRoot)), dtReportDate)
        Call CopyFileToFolder(strStagePath, strDailyFolder)
    Next lngRoot

    Call DeleteFileIfExists(strStagePath)

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "fileName;stagePath;destinationFolder", strFileName, strStagePath, strDailyFolder)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    arrReport As Variant; strStagePath As String
' Returns:       ---
' Description:   Creates an XLSX report with a hidden Static_Data worksheet.
'-------------------------------------------------------------------------------
Private Sub CreateReportWorkbook(ByRef arrReport As Variant, ByVal strStagePath As String)
    Const METHOD_NAME As String = "CreateReportWorkbook"
    Dim blnCloseRequired As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim handlerErrDescription As String
    Dim handlerErrNumber As Long
    Dim lngLastRow As Long
    Dim wkbOutput As Excel.Workbook
    Dim wksTransactions As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set wkbOutput = Application.Workbooks.Add(xlWBATWorksheet)
    blnCloseRequired = True
    Set wksTransactions = wkbOutput.Worksheets(1)
    wksTransactions.Name = SHEET_TRANSACTIONS

    lngLastRow = UBound(arrReport, 1)
    wksTransactions.Range(wksTransactions.Cells(1, 1), wksTransactions.Cells(lngLastRow, OUTPUT_COLUMN_COUNT)).Value2 = arrReport

    Call FormatTransactionWorksheet(wksTransactions, lngLastRow)
    Call CopyStaticDataValues(wkbOutput)

    Call wkbOutput.SaveAs(Filename:=strStagePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False, AddToMru:=False, Local:=True)

ExitPoint:
    Set wksTransactions = Nothing

    If blnCloseRequired Then
        blnCloseRequired = False
        Call wkbOutput.Close(SaveChanges:=False)
    End If

    Set wkbOutput = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    handlerErrNumber = VBA.Err.Number
    handlerErrDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, handlerErrNumber, handlerErrDescription, "stagePath", strStagePath)

    If errNumber = 0 Then
        errNumber = handlerErrNumber
        errDescription = handlerErrDescription
    End If

    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    wkbOutput As Excel.Workbook
' Returns:       ---
' Description:   Copies Static_Data values into a hidden output worksheet.
'-------------------------------------------------------------------------------
Private Sub CopyStaticDataValues(ByVal wkbOutput As Excel.Workbook)
    Const METHOD_NAME As String = "CopyStaticDataValues"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngFirstColumn As Long
    Dim lngLastColumn As Long
    Dim rngSource As Excel.Range
    Dim rngTarget As Excel.Range
    Dim wksSource As Excel.Worksheet
    Dim wksTarget As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set wksSource = ThisWorkbook.Worksheets(SHEET_STATIC_DATA)
    Set wksTarget = wkbOutput.Worksheets.Add(After:=wkbOutput.Worksheets(wkbOutput.Worksheets.Count))
    wksTarget.Name = SHEET_STATIC_DATA

    Set rngSource = wksSource.UsedRange
    Set rngTarget = wksTarget.Range(rngSource.Address)
    rngTarget.Value2 = rngSource.Value2

    lngFirstColumn = rngSource.Column
    lngLastColumn = rngSource.Column + rngSource.Columns.Count - 1

    For lngColumn = lngFirstColumn To lngLastColumn
        wksTarget.Columns(lngColumn).ColumnWidth = wksSource.Columns(lngColumn).ColumnWidth
    Next lngColumn

    wksTarget.Visible = xlSheetHidden

ExitPoint:
    Set rngTarget = Nothing
    Set rngSource = Nothing
    Set wksTarget = Nothing
    Set wksSource = Nothing
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
' Parameters:    wksTransactions As Excel.Worksheet; lngLastRow As Long
' Returns:       ---
' Description:   Applies stable report formatting without Select or Activate.
'-------------------------------------------------------------------------------
Private Sub FormatTransactionWorksheet(ByVal wksTransactions As Excel.Worksheet, ByVal lngLastRow As Long)
    Const METHOD_NAME As String = "FormatTransactionWorksheet"
    Dim errDescription As String
    Dim errNumber As Long
    Dim rngTable As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set rngTable = wksTransactions.Range(wksTransactions.Cells(1, 1), wksTransactions.Cells(lngLastRow, OUTPUT_COLUMN_COUNT))

    wksTransactions.Rows(1).Font.Bold = True
    wksTransactions.Rows(1).WrapText = True
    Call rngTable.AutoFilter

    wksTransactions.Columns("A:B").ColumnWidth = 18
    wksTransactions.Columns("C:D").ColumnWidth = 12
    wksTransactions.Columns("E:E").ColumnWidth = 10
    wksTransactions.Columns("F:G").ColumnWidth = 18
    wksTransactions.Columns("H:I").ColumnWidth = 12
    wksTransactions.Columns("J:Q").ColumnWidth = 14
    wksTransactions.Columns("R:S").ColumnWidth = 15
    wksTransactions.Columns("T:U").ColumnWidth = 10
    wksTransactions.Columns("C:D").NumberFormat = "dd.mm.yyyy"
    wksTransactions.Columns("R:S").NumberFormat = "#,##0.00"

ExitPoint:
    Set rngTable = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "lastRow", lngLastRow)
    GoTo ExitPoint
End Sub
