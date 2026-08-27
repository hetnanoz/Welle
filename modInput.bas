Option Explicit

Private Const CLASS_NAME As String = "modInput"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    colAttachmentPaths As Collection; strWorksheetName As String; dictAttachmentSubjects As Object
' Returns:       Variant
' Description:   Merges source files after mapping input columns by header name.
'-------------------------------------------------------------------------------
Public Function MergeInputFiles(ByVal colAttachmentPaths As Collection, ByVal strWorksheetName As String, ByVal dictAttachmentSubjects As Object) As Variant
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
    Dim strMailSubject As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If colAttachmentPaths Is Nothing Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Attachment collection is not available.")

    lngAttachmentCount = colAttachmentPaths.Count
    If lngAttachmentCount = 0 Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "No input files were supplied for merging.")

    lngCapacity = INITIAL_INPUT_CAPACITY
    ReDim arrWorking(1 To OUTPUT_COLUMN_COUNT, 1 To lngCapacity)

    For lngFile = 1 To lngAttachmentCount
        strFilePath = CStr(colAttachmentPaths(lngFile))
        strMailSubject = vbNullString

        If Not dictAttachmentSubjects Is Nothing Then
            If dictAttachmentSubjects.Exists(strFilePath) Then strMailSubject = CStr(dictAttachmentSubjects(strFilePath))
        End If

        Application.StatusBar = "Dallas Cash Transactions: merging input file " & CStr(lngFile) & "/" & CStr(lngAttachmentCount) & " - " & GetFileNameFromPath(strFilePath)
        DoEvents
        Call ReadInputFileData(strFilePath, strWorksheetName, strMailSubject, arrFileData, lngFileRows)

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
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "attachmentCount;currentFile;worksheetName;mailSubject", lngAttachmentCount, strFilePath, strWorksheetName, strMailSubject)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strFilePath As String; strWorksheetName As String; strMailSubject As String; arrData As Variant; lngDataRows As Long
' Returns:       ---
' Description:   Reads one input workbook and normalizes its columns by header.
'-------------------------------------------------------------------------------
Private Sub ReadInputFileData(ByVal strFilePath As String, ByVal strWorksheetName As String, ByVal strMailSubject As String, ByRef arrData As Variant, ByRef lngDataRows As Long)
    Const METHOD_NAME As String = "ReadInputFileData"
    Dim arrSource As Variant
    Dim blnAccountFormatMixed As Boolean
    Dim blnCloseRequired As Boolean
    Dim dictColumnMap As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim handlerErrDescription As String
    Dim handlerErrNumber As Long
    Dim lngAccountSourceColumn As Long
    Dim lngLastColumn As Long
    Dim lngLastRow As Long
    Dim lngOutputRow As Long
    Dim lngSourceColumn As Long
    Dim lngSourceRow As Long
    Dim lngTargetColumn As Long
    Dim strAccountContext As String
    Dim strAccountNumberFormat As String
    Dim strDisplayedAccount As String
    Dim varAccountNumberFormat As Variant
    Dim wkbAlreadyOpen As Excel.Workbook
    Dim wkbInput As Excel.Workbook
    Dim wksCandidate As Excel.Worksheet
    Dim wksInput As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrData = Empty
    lngDataRows = 0

    Set wkbAlreadyOpen = GetOpenWorkbookByFullName(strFilePath, False)

    If wkbAlreadyOpen Is Nothing Then
        Set wkbInput = Application.Workbooks.Open(Filename:=strFilePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
        blnCloseRequired = True
    Else
        Set wkbInput = wkbAlreadyOpen
    End If

    If wkbInput Is Nothing Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Input workbook could not be opened or reused: '" & strFilePath & "'. Outlook mail subject: '" & strMailSubject & "'.")
    If wkbInput.Worksheets.Count = 0 Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Input workbook does not contain a worksheet. Outlook mail subject: '" & strMailSubject & "'.")

    For Each wksCandidate In wkbInput.Worksheets
        If StrComp(CStr(wksCandidate.Name), strWorksheetName, vbTextCompare) = 0 Then
            Set wksInput = wksCandidate
            Exit For
        End If
    Next wksCandidate

    If wksInput Is Nothing Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Worksheet '" & strWorksheetName & "' was not found in input file '" & strFilePath & "'. Outlook mail subject: '" & strMailSubject & "'.")

    lngLastColumn = GetInputLastHeaderColumn(wksInput)
    lngLastRow = GetInputLastRow(wksInput, lngLastColumn)

    If lngLastRow < INPUT_FIRST_DATA_ROW Then GoTo ExitPoint

    arrSource = wksInput.Range(wksInput.Cells(INPUT_FIRST_DATA_ROW, 1), wksInput.Cells(lngLastRow, lngLastColumn)).Value2

    If Not ArrayContainsInputData(arrSource) Then GoTo ExitPoint

    Set dictColumnMap = BuildInputColumnMap(wksInput, lngLastColumn, strFilePath, strMailSubject)

    If dictColumnMap.Exists(1) Then
        lngAccountSourceColumn = CLng(dictColumnMap(1))
        varAccountNumberFormat = wksInput.Range(wksInput.Cells(INPUT_FIRST_DATA_ROW, lngAccountSourceColumn), wksInput.Cells(lngLastRow, lngAccountSourceColumn)).NumberFormat

        If IsNull(varAccountNumberFormat) Then
            blnAccountFormatMixed = True
        Else
            strAccountNumberFormat = CStr(varAccountNumberFormat)
        End If
    End If

    lngDataRows = CountInputDataRows(arrSource, dictColumnMap)

    If lngDataRows = 0 Then GoTo ExitPoint

    ReDim arrData(1 To lngDataRows, 1 To INPUT_COLUMN_COUNT)

    For lngSourceRow = 1 To UBound(arrSource, 1)
        If InputRowHasData(arrSource, lngSourceRow, dictColumnMap) Then
            lngOutputRow = lngOutputRow + 1

            For lngTargetColumn = 1 To INPUT_COLUMN_COUNT
                If dictColumnMap.Exists(lngTargetColumn) Then
                    lngSourceColumn = CLng(dictColumnMap(lngTargetColumn))

                    If lngTargetColumn = 1 Then
                        strAccountContext = "Input file: '" & strFilePath & "'. Outlook mail subject: '" & strMailSubject & "'."

                        If IsNumeric(arrSource(lngSourceRow, lngSourceColumn)) Then
                            strDisplayedAccount = GetFastDisplayedAccountText(arrSource(lngSourceRow, lngSourceColumn), strAccountNumberFormat)

                            If blnAccountFormatMixed Then
                                strDisplayedAccount = CStr(wksInput.Cells(INPUT_FIRST_DATA_ROW + lngSourceRow - 1, lngSourceColumn).Text)
                            End If
                        ElseIf IsError(arrSource(lngSourceRow, lngSourceColumn)) Then
                            strDisplayedAccount = vbNullString
                        ElseIf IsNull(arrSource(lngSourceRow, lngSourceColumn)) Or IsEmpty(arrSource(lngSourceRow, lngSourceColumn)) Then
                            strDisplayedAccount = vbNullString
                        Else
                            strDisplayedAccount = Trim$(CStr(arrSource(lngSourceRow, lngSourceColumn)))
                        End If

                        arrData(lngOutputRow, lngTargetColumn) = NormalizeAccountIdentifier(arrSource(lngSourceRow, lngSourceColumn), strDisplayedAccount, ERROR_INPUT_DATA, strAccountContext)
                    ElseIf lngTargetColumn = 3 Or lngTargetColumn = 4 Then
                        strAccountContext = "Input file: '" & strFilePath & "'. Outlook mail subject: '" & strMailSubject & "'."
                        arrData(lngOutputRow, lngTargetColumn) = NormalizeInputDateValue(arrSource(lngSourceRow, lngSourceColumn), strAccountContext)
                    Else
                        arrData(lngOutputRow, lngTargetColumn) = arrSource(lngSourceRow, lngSourceColumn)
                    End If
                Else
                    arrData(lngOutputRow, lngTargetColumn) = vbNullString
                End If
            Next lngTargetColumn
        End If
    Next lngSourceRow

ExitPoint:
    Set dictColumnMap = Nothing
    Set wksCandidate = Nothing
    Set wksInput = Nothing
    Set wkbAlreadyOpen = Nothing

    If blnCloseRequired Then
        blnCloseRequired = False
        If Not wkbInput Is Nothing Then Call wkbInput.Close(SaveChanges:=False)
    End If

    Set wkbInput = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    handlerErrNumber = VBA.Err.Number
    handlerErrDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, handlerErrNumber, handlerErrDescription, "filePath;worksheetName;mailSubject", strFilePath, strWorksheetName, strMailSubject)

    If errNumber = 0 Then
        errNumber = handlerErrNumber
        errDescription = handlerErrDescription
    End If

    GoTo ExitPoint
End Sub



'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    varValue As Variant; strNumberFormat As String
' Returns:       String
' Description:   Reconstructs leading-zero Account display text without a per-row worksheet COM call.
'-------------------------------------------------------------------------------
Private Function GetFastDisplayedAccountText(ByVal varValue As Variant, ByVal strNumberFormat As String) As String
    Const METHOD_NAME As String = "GetFastDisplayedAccountText"
    Dim blnZeroMask As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPosition As Long
    Dim strCharacter As String
    Dim strFormat As String
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not IsNumeric(varValue) Then GoTo ExitPoint

    strFormat = Trim$(strNumberFormat)
    If Len(strFormat) = 0 Then GoTo ExitPoint

    blnZeroMask = True

    For lngPosition = 1 To Len(strFormat)
        strCharacter = Mid$(strFormat, lngPosition, 1)

        If strCharacter <> "0" Then
            blnZeroMask = False
            Exit For
        End If
    Next lngPosition

    If blnZeroMask Then strResult = Format$(CDbl(varValue), strFormat)

ExitPoint:
    If errNumber = 0 Then GetFastDisplayedAccountText = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "numberFormat", strNumberFormat)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    varValue As Variant; strContext As String
' Returns:       Variant
' Description:   Normalizes an input date to an Excel date serial without time.
'-------------------------------------------------------------------------------
Private Function NormalizeInputDateValue(ByVal varValue As Variant, ByVal strContext As String) As Variant
    Const METHOD_NAME As String = "NormalizeInputDateValue"
    Dim dblResult As Double
    Dim dtValue As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngDay As Long
    Dim lngMonth As Long
    Dim lngYear As Long
    Dim strValue As String
    Dim varResult As Variant

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(varValue) Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "An Excel error value cannot be converted to a date. " & strContext)

    If IsNull(varValue) Or IsEmpty(varValue) Then
        varResult = vbNullString
        GoTo ExitPoint
    End If

    strValue = Trim$(CStr(varValue))

    If Len(strValue) = 0 Then
        varResult = vbNullString
        GoTo ExitPoint
    End If

    If IsNumeric(varValue) Then
        dblResult = Fix(CDbl(varValue))

        If dblResult <= 0 Or dblResult > 2958465# Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Invalid Excel serial date value: " & strValue & ". " & strContext)

        varResult = dblResult
    ElseIf Len(strValue) = 10 And Mid$(strValue, 5, 1) = "-" And Mid$(strValue, 8, 1) = "-" Then
        If Not IsNumeric(Left$(strValue, 4)) Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Invalid ISO date value: " & strValue & ". " & strContext)
        If Not IsNumeric(Mid$(strValue, 6, 2)) Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Invalid ISO date value: " & strValue & ". " & strContext)
        If Not IsNumeric(Right$(strValue, 2)) Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Invalid ISO date value: " & strValue & ". " & strContext)

        lngYear = CLng(Left$(strValue, 4))
        lngMonth = CLng(Mid$(strValue, 6, 2))
        lngDay = CLng(Right$(strValue, 2))
        dtValue = DateSerial(lngYear, lngMonth, lngDay)

        If Year(dtValue) <> lngYear Or Month(dtValue) <> lngMonth Or Day(dtValue) <> lngDay Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Invalid ISO date value: " & strValue & ". " & strContext)

        varResult = CDbl(dtValue)
    ElseIf IsDate(varValue) Then
        varResult = CDbl(DateValue(CDate(varValue)))
    Else
        Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Unsupported input date value: " & strValue & ". " & strContext)
    End If

ExitPoint:
    If errNumber = 0 Then NormalizeInputDateValue = varResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "dateValue;context", strValue, strContext)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    wksInput As Excel.Worksheet; lngLastColumn As Long; strFilePath As String; strMailSubject As String
' Returns:       Object
' Description:   Maps source header positions to the canonical sixteen columns.
'-------------------------------------------------------------------------------
Private Function BuildInputColumnMap(ByVal wksInput As Excel.Worksheet, ByVal lngLastColumn As Long, ByVal strFilePath As String, ByVal strMailSubject As String) As Object
    Const METHOD_NAME As String = "BuildInputColumnMap"
    Dim arrHeaders As Variant
    Dim dictColumnMap As Object
    Dim dictExpected As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngTargetColumn As Long
    Dim strActualHeader As String
    Dim strExpectedHeader As String
    Dim strHeaderKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrHeaders = GetInputHeaders()

    Set dictExpected = CreateObject("Scripting.Dictionary")
    dictExpected.CompareMode = vbTextCompare

    Set dictColumnMap = CreateObject("Scripting.Dictionary")
    dictColumnMap.CompareMode = vbTextCompare

    For lngTargetColumn = 1 To INPUT_COLUMN_COUNT
        strExpectedHeader = CStr(arrHeaders(lngTargetColumn - 1))
        strHeaderKey = NormalizeHeaderName(strExpectedHeader)
        dictExpected(strHeaderKey) = lngTargetColumn
    Next lngTargetColumn

    For lngColumn = 1 To lngLastColumn
        strActualHeader = GetHeaderText(wksInput.Cells(INPUT_HEADER_ROW, lngColumn).Value2)
        strHeaderKey = NormalizeHeaderName(strActualHeader)

        If Len(strHeaderKey) > 0 Then
            If dictExpected.Exists(strHeaderKey) Then
                lngTargetColumn = CLng(dictExpected(strHeaderKey))

                If dictColumnMap.Exists(lngTargetColumn) Then
                    strExpectedHeader = CStr(arrHeaders(lngTargetColumn - 1))
                    Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Duplicate input header matching '" & strExpectedHeader & "' was found. Input file: '" & strFilePath & "'. Outlook mail subject: '" & strMailSubject & "'.")
                End If

                dictColumnMap.Add lngTargetColumn, lngColumn
            End If
        End If
    Next lngColumn

    For lngTargetColumn = 1 To INPUT_COLUMN_COUNT
        If IsRequiredInputHeader(lngTargetColumn) Then
            If Not dictColumnMap.Exists(lngTargetColumn) Then
                strExpectedHeader = CStr(arrHeaders(lngTargetColumn - 1))
                Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "Required input header '" & strExpectedHeader & "' was not found in a non-empty input file. Input file: '" & strFilePath & "'. Outlook mail subject: '" & strMailSubject & "'.")
            End If
        End If
    Next lngTargetColumn

ExitPoint:
    Set dictExpected = Nothing
    If errNumber = 0 Then Set BuildInputColumnMap = dictColumnMap
    Set dictColumnMap = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "filePath;column;mailSubject", strFilePath, lngColumn, strMailSubject)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    varHeader As Variant
' Returns:       String
' Description:   Converts one header cell to trimmed text.
'-------------------------------------------------------------------------------
Private Function GetHeaderText(ByVal varHeader As Variant) As String
    Const METHOD_NAME As String = "GetHeaderText"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(varHeader) Then Call VBA.Err.Raise(ERROR_INPUT_DATA, METHOD_NAME, "An Excel error value cannot be used as an input header.")

    If IsNull(varHeader) Or IsEmpty(varHeader) Then
        strResult = vbNullString
    Else
        strResult = Trim$(CStr(varHeader))
    End If

ExitPoint:
    If errNumber = 0 Then GetHeaderText = strResult
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
' Creation date: 2026-08-27
' Parameters:    strHeader As String
' Returns:       String
' Description:   Normalizes source headers and supported known aliases.
'-------------------------------------------------------------------------------
Private Function NormalizeHeaderName(ByVal strHeader As String) As String
    Const METHOD_NAME As String = "NormalizeHeaderName"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = Trim$(strHeader)
    strResult = Replace$(strResult, Chr$(160), " ")
    strResult = Replace$(strResult, vbCr, " ")
    strResult = Replace$(strResult, vbLf, " ")
    strResult = Replace$(strResult, vbTab, " ")

    Do While InStr(1, strResult, "  ", vbBinaryCompare) > 0
        strResult = Replace$(strResult, "  ", " ")
    Loop

    strResult = LCase$(strResult)
    strResult = Replace$(strResult, ".", vbNullString)

    If StrComp(strResult, "accout name", vbTextCompare) = 0 Then strResult = "account name"

ExitPoint:
    If errNumber = 0 Then NormalizeHeaderName = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "header", strHeader)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    lngTargetColumn As Long
' Returns:       Boolean
' Description:   Treats Account Name as optional and all other source headers as required.
'-------------------------------------------------------------------------------
Private Function IsRequiredInputHeader(ByVal lngTargetColumn As Long) As Boolean
    Const METHOD_NAME As String = "IsRequiredInputHeader"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnResult = lngTargetColumn <> 2

ExitPoint:
    If errNumber = 0 Then IsRequiredInputHeader = blnResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "targetColumn", lngTargetColumn)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    arrSource As Variant
' Returns:       Boolean
' Description:   Checks whether the data area contains any non-empty values.
'-------------------------------------------------------------------------------
Private Function ArrayContainsInputData(ByRef arrSource As Variant) As Boolean
    Const METHOD_NAME As String = "ArrayContainsInputData"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngRow As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngRow = 1 To UBound(arrSource, 1)
        For lngColumn = 1 To UBound(arrSource, 2)
            If IsInputValuePopulated(arrSource(lngRow, lngColumn)) Then
                blnResult = True
                GoTo ExitPoint
            End If
        Next lngColumn
    Next lngRow

ExitPoint:
    If errNumber = 0 Then ArrayContainsInputData = blnResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "row;column", lngRow, lngColumn)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    arrSource As Variant; dictColumnMap As Object
' Returns:       Long
' Description:   Counts non-empty transaction rows after header mapping.
'-------------------------------------------------------------------------------
Private Function CountInputDataRows(ByRef arrSource As Variant, ByVal dictColumnMap As Object) As Long
    Const METHOD_NAME As String = "CountInputDataRows"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngResult As Long
    Dim lngRow As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngRow = 1 To UBound(arrSource, 1)
        If InputRowHasData(arrSource, lngRow, dictColumnMap) Then lngResult = lngResult + 1
    Next lngRow

ExitPoint:
    If errNumber = 0 Then CountInputDataRows = lngResult
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
' Creation date: 2026-08-27
' Parameters:    arrSource As Variant; lngRow As Long; dictColumnMap As Object
' Returns:       Boolean
' Description:   Checks whether a source row contains data in any recognized column.
'-------------------------------------------------------------------------------
Private Function InputRowHasData(ByRef arrSource As Variant, ByVal lngRow As Long, ByVal dictColumnMap As Object) As Boolean
    Const METHOD_NAME As String = "InputRowHasData"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngSourceColumn As Long
    Dim lngTargetColumn As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngTargetColumn = 1 To INPUT_COLUMN_COUNT
        If dictColumnMap.Exists(lngTargetColumn) Then
            lngSourceColumn = CLng(dictColumnMap(lngTargetColumn))

            If IsInputValuePopulated(arrSource(lngRow, lngSourceColumn)) Then
                blnResult = True
                GoTo ExitPoint
            End If
        End If
    Next lngTargetColumn

ExitPoint:
    If errNumber = 0 Then InputRowHasData = blnResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "row;targetColumn", lngRow, lngTargetColumn)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    varValue As Variant
' Returns:       Boolean
' Description:   Checks whether one input value should count as populated.
'-------------------------------------------------------------------------------
Private Function IsInputValuePopulated(ByVal varValue As Variant) As Boolean
    Const METHOD_NAME As String = "IsInputValuePopulated"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(varValue) Then
        blnResult = True
    ElseIf IsNull(varValue) Or IsEmpty(varValue) Then
        blnResult = False
    Else
        blnResult = Len(Trim$(CStr(varValue))) > 0
    End If

ExitPoint:
    If errNumber = 0 Then IsInputValuePopulated = blnResult
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
' Creation date: 2026-08-27
' Parameters:    wksInput As Excel.Worksheet
' Returns:       Long
' Description:   Finds the last non-empty header column on the configured input sheet.
'-------------------------------------------------------------------------------
Private Function GetInputLastHeaderColumn(ByVal wksInput As Excel.Worksheet) As Long
    Const METHOD_NAME As String = "GetInputLastHeaderColumn"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngResult As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngResult = wksInput.Cells(INPUT_HEADER_ROW, wksInput.Columns.Count).End(xlToLeft).Column

ExitPoint:
    If errNumber = 0 Then GetInputLastHeaderColumn = lngResult
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
' Creation date: 2026-08-27
' Parameters:    wksInput As Excel.Worksheet; lngLastColumn As Long
' Returns:       Long
' Description:   Finds the last populated row across all source columns under the header.
'-------------------------------------------------------------------------------
Private Function GetInputLastRow(ByVal wksInput As Excel.Worksheet, ByVal lngLastColumn As Long) As Long
    Const METHOD_NAME As String = "GetInputLastRow"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngColumnLastRow As Long
    Dim lngResult As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngColumn = 1 To lngLastColumn
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
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "lastColumn", lngLastColumn)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    ---
' Returns:       Variant
' Description:   Returns the canonical A:P input headers used in final reports.
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
