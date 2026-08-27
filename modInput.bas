Option Explicit

Private Const CLASS_NAME As String = "modInput"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    colAttachmentPaths As Collection; strWorksheetName As String
' Returns:       Variant
' Description:   Merges data from the configured worksheet in all A:P input files.
'-------------------------------------------------------------------------------
Public Function MergeInputFiles(ByVal colAttachmentPaths As Collection, ByVal strWorksheetName As String) As Variant
    Const METHOD_NAME As String = "MergeInputFiles"
    Dim arrCombined As Variant
    Dim arrFileData As Variant
    Dim arrWorking As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngAttachmentCount As Long
    Dim lngCapacity As Long
    Dim lngColumn As Long
    Dim lngFile As Long
    Dim lngFileRows As Long
    Dim lngRecord As Long
    Dim lngRecordCount As Long
    Dim lngRequiredCapacity As Long
    Dim lngRow As Long
    Dim strFilePath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If colAttachmentPaths Is Nothing Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Attachment collection is not available.")

    lngAttachmentCount = colAttachmentPaths.Count
    If lngAttachmentCount = 0 Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "No input files were supplied for merging.")

    lngCapacity = INITIAL_INPUT_CAPACITY
    ReDim arrWorking(1 To OUTPUT_COLUMN_COUNT, 1 To lngCapacity)

    For lngFile = 1 To lngAttachmentCount
        strFilePath = CStr(colAttachmentPaths(lngFile))
        Call ReadInputFileData(strFilePath, strWorksheetName, arrFileData, lngFileRows)

        If lngFileRows > 0 Then
            lngRequiredCapacity = lngRecordCount + lngFileRows
            Call EnsureWorkingCapacity(arrWorking, lngCapacity, lngRequiredCapacity)

            For lngRow = 1 To lngFileRows
                lngRecordCount = lngRecordCount + 1

                For lngColumn = 1 To INPUT_COLUMN_COUNT
                    arrWorking(lngColumn + 5, lngRecordCount) = arrFileData(lngRow, lngColumn)
                Next lngColumn
            Next lngRow
        End If

        arrFileData = Empty
    Next lngFile

    ReDim arrCombined(1 To lngRecordCount + 1, 1 To OUTPUT_COLUMN_COUNT)
    Call SetOutputHeaders(arrCombined)

    For lngRecord = 1 To lngRecordCount
        For lngColumn = 6 To OUTPUT_COLUMN_COUNT
            arrCombined(lngRecord + 1, lngColumn) = arrWorking(lngColumn, lngRecord)
        Next lngColumn
    Next lngRecord

ExitPoint:
    If errNumber = 0 Then MergeInputFiles = arrCombined
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "attachmentCount;currentFile;worksheetName", lngAttachmentCount, strFilePath, strWorksheetName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strFilePath As String; strWorksheetName As String; arrData As Variant; lngDataRows As Long
' Returns:       ---
' Description:   Reads the configured worksheet from one validated input workbook.
'-------------------------------------------------------------------------------
Private Sub ReadInputFileData(ByVal strFilePath As String, ByVal strWorksheetName As String, ByRef arrData As Variant, ByRef lngDataRows As Long)
    Const METHOD_NAME As String = "ReadInputFileData"
    Dim blnCloseRequired As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim handlerErrDescription As String
    Dim handlerErrNumber As Long
    Dim lngLastRow As Long
    Dim wkbInput As Excel.Workbook
    Dim wksCandidate As Excel.Worksheet
    Dim wksInput As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrData = Empty
    lngDataRows = 0

    Set wkbInput = Application.Workbooks.Open(Filename:=strFilePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
    blnCloseRequired = True

    If wkbInput.Worksheets.Count = 0 Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Input workbook does not contain a worksheet.")

    For Each wksCandidate In wkbInput.Worksheets
        If StrComp(CStr(wksCandidate.Name), strWorksheetName, vbTextCompare) = 0 Then
            Set wksInput = wksCandidate
            Exit For
        End If
    Next wksCandidate

    If wksInput Is Nothing Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Worksheet '" & strWorksheetName & "' was not found in input file '" & strFilePath & "'.")
    Call ValidateInputHeader(wksInput, strFilePath)

    lngLastRow = GetInputLastRow(wksInput)

    If lngLastRow >= INPUT_FIRST_DATA_ROW Then
        lngDataRows = lngLastRow - INPUT_HEADER_ROW
        arrData = wksInput.Range(wksInput.Cells(INPUT_FIRST_DATA_ROW, 1), wksInput.Cells(lngLastRow, INPUT_COLUMN_COUNT)).Value2
    End If

ExitPoint:
    Set wksCandidate = Nothing
    Set wksInput = Nothing

    If blnCloseRequired Then
        blnCloseRequired = False
        Call wkbInput.Close(SaveChanges:=False)
    End If

    Set wkbInput = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    handlerErrNumber = VBA.Err.Number
    handlerErrDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, handlerErrNumber, handlerErrDescription, "filePath;worksheetName", strFilePath, strWorksheetName)

    If errNumber = 0 Then
        errNumber = handlerErrNumber
        errDescription = handlerErrDescription
    End If

    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    wksInput As Excel.Worksheet; strFilePath As String
' Returns:       ---
' Description:   Validates the required sixteen-column input header.
'-------------------------------------------------------------------------------
Private Sub ValidateInputHeader(ByVal wksInput As Excel.Worksheet, ByVal strFilePath As String)
    Const METHOD_NAME As String = "ValidateInputHeader"
    Dim arrHeaders As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim strActual As String
    Dim strExpected As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrHeaders = GetInputHeaders()

    For lngColumn = 1 To INPUT_COLUMN_COUNT
        strActual = Trim$(CStr(wksInput.Cells(INPUT_HEADER_ROW, lngColumn).Value2))
        strExpected = CStr(arrHeaders(lngColumn - 1))

        If StrComp(strActual, strExpected, vbTextCompare) <> 0 Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Unexpected header in column " & CStr(lngColumn) & ". Expected '" & strExpected & "' and found '" & strActual & "'.")
    Next lngColumn

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "filePath;column", strFilePath, lngColumn)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    arrWorking As Variant; lngCapacity As Long; lngRequired As Long
' Returns:       ---
' Description:   Expands the column-major working array using capacity doubling.
'-------------------------------------------------------------------------------
Private Sub EnsureWorkingCapacity(ByRef arrWorking As Variant, ByRef lngCapacity As Long, ByVal lngRequired As Long)
    Const METHOD_NAME As String = "EnsureWorkingCapacity"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngNewCapacity As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If lngRequired <= lngCapacity Then GoTo ExitPoint

    lngNewCapacity = lngCapacity

    Do While lngNewCapacity < lngRequired
        lngNewCapacity = lngNewCapacity * 2
    Loop

    ReDim Preserve arrWorking(1 To OUTPUT_COLUMN_COUNT, 1 To lngNewCapacity)
    lngCapacity = lngNewCapacity

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "currentCapacity;requiredCapacity", lngCapacity, lngRequired)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    wksInput As Excel.Worksheet
' Returns:       Long
' Description:   Finds the last populated row across all sixteen input columns.
'-------------------------------------------------------------------------------
Private Function GetInputLastRow(ByVal wksInput As Excel.Worksheet) As Long
    Const METHOD_NAME As String = "GetInputLastRow"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngColumnLastRow As Long
    Dim lngResult As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngColumn = 1 To INPUT_COLUMN_COUNT
        lngColumnLastRow = wksInput.Cells(wksInput.Rows.Count, lngColumn).End(xlUp).Row
        If lngColumnLastRow > lngResult Then lngResult = lngColumnLastRow
    Next lngColumn

ExitPoint:
    If errNumber = 0 Then GetInputLastRow = lngResult
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
' Parameters:    ---
' Returns:       Variant
' Description:   Returns the canonical A:P input headers.
'-------------------------------------------------------------------------------
Private Function GetInputHeaders() As Variant
    Const METHOD_NAME As String = "GetInputHeaders"
    Dim arrResult As Variant
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrResult = Array("Account", "Account Name", "v/d", "t/d", "Journal", "BLZ", "Ref. 1", "Ref. 2", "Ref. 3", "Ref. 4", "Ref. 5", "Ref. 6", "EUR amount", "Amount", "CCY", "Reversal")

ExitPoint:
    If errNumber = 0 Then GetInputHeaders = arrResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function
