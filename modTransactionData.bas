Option Explicit

Private Const CLASS_NAME As String = "modTransactionData"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    ---
' Returns:       Variant
' Description:   Returns the canonical A:U output headers.
'-------------------------------------------------------------------------------
Public Function GetOutputHeaders() As Variant
    Const METHOD_NAME As String = "GetOutputHeaders"
    Dim arrResult As Variant
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrResult = Array("MF", "Master", "VD", "TD", "Team", "Account", "Account Name", "v/d", "t/d", "Journal", "BLZ", "Ref. 1", "Ref. 2", "Ref. 3", "Ref. 4", "Ref. 5", "Ref. 6", "EUR amount", "Amount", "CCY", "Reversal")

ExitPoint:
    If errNumber = 0 Then GetOutputHeaders = arrResult
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
' Parameters:    arrData As Variant
' Returns:       ---
' Description:   Writes the canonical output header into row one of an array.
'-------------------------------------------------------------------------------
Public Sub SetOutputHeaders(ByRef arrData As Variant)
    Const METHOD_NAME As String = "SetOutputHeaders"
    Dim arrHeaders As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrHeaders = GetOutputHeaders()

    For lngColumn = 0 To UBound(arrHeaders)
        arrData(1, lngColumn + 1) = arrHeaders(lngColumn)
    Next lngColumn

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
' Parameters:    arrData As Variant; strTeam As String
' Returns:       Variant
' Description:   Returns the header and only rows assigned to one team.
'-------------------------------------------------------------------------------
Public Function FilterTransactionsByTeam(ByRef arrData As Variant, ByVal strTeam As String) As Variant
    Const METHOD_NAME As String = "FilterTransactionsByTeam"
    Dim arrResult As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngMatchCount As Long
    Dim lngOutputRow As Long
    Dim lngRow As Long
    Dim strRowTeam As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngRow = 2 To UBound(arrData, 1)
        strRowTeam = Trim$(CStr(arrData(lngRow, 5)))
        If StrComp(strRowTeam, strTeam, vbTextCompare) = 0 Then lngMatchCount = lngMatchCount + 1
    Next lngRow

    ReDim arrResult(1 To lngMatchCount + 1, 1 To OUTPUT_COLUMN_COUNT)

    For lngColumn = 1 To OUTPUT_COLUMN_COUNT
        arrResult(1, lngColumn) = arrData(1, lngColumn)
    Next lngColumn

    lngOutputRow = 1

    For lngRow = 2 To UBound(arrData, 1)
        strRowTeam = Trim$(CStr(arrData(lngRow, 5)))

        If StrComp(strRowTeam, strTeam, vbTextCompare) = 0 Then
            lngOutputRow = lngOutputRow + 1

            For lngColumn = 1 To OUTPUT_COLUMN_COUNT
                arrResult(lngOutputRow, lngColumn) = arrData(lngRow, lngColumn)
            Next lngColumn
        End If
    Next lngRow

ExitPoint:
    If errNumber = 0 Then FilterTransactionsByTeam = arrResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "team", strTeam)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    arrData As Variant
' Returns:       Variant
' Description:   Removes duplicate current-run transactions using the historical key.
'-------------------------------------------------------------------------------
Public Function DeduplicateCurrentTransactions(ByRef arrData As Variant) As Variant
    Const METHOD_NAME As String = "DeduplicateCurrentTransactions"
    Dim arrInclude() As Boolean
    Dim arrResult As Variant
    Dim dictKeys As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngLastRow As Long
    Dim lngOutputRow As Long
    Dim lngRow As Long
    Dim lngUniqueCount As Long
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = UBound(arrData, 1)

    If lngLastRow <= 1 Then
        arrResult = arrData
        GoTo ExitPoint
    End If

    Set dictKeys = CreateObject("Scripting.Dictionary")
    dictKeys.CompareMode = vbBinaryCompare
    ReDim arrInclude(2 To lngLastRow)

    For lngRow = 2 To lngLastRow
        strKey = BuildTransactionKey(arrData, lngRow)

        If Not dictKeys.Exists(strKey) Then
            dictKeys.Add strKey, True
            arrInclude(lngRow) = True
            lngUniqueCount = lngUniqueCount + 1
        End If
    Next lngRow

    ReDim arrResult(1 To lngUniqueCount + 1, 1 To OUTPUT_COLUMN_COUNT)

    For lngColumn = 1 To OUTPUT_COLUMN_COUNT
        arrResult(1, lngColumn) = arrData(1, lngColumn)
    Next lngColumn

    lngOutputRow = 1

    For lngRow = 2 To lngLastRow
        If arrInclude(lngRow) Then
            lngOutputRow = lngOutputRow + 1

            For lngColumn = 1 To OUTPUT_COLUMN_COUNT
                arrResult(lngOutputRow, lngColumn) = arrData(lngRow, lngColumn)
            Next lngColumn
        End If
    Next lngRow

ExitPoint:
    Set dictKeys = Nothing
    If errNumber = 0 Then DeduplicateCurrentTransactions = arrResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "row;uniqueCount", lngRow, lngUniqueCount)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    varValue As Variant; strDisplayedText As String; lngBusinessError As Long; strContext As String
' Returns:       String
' Description:   Normalizes Account identifiers and protects against Excel numeric precision loss.
'-------------------------------------------------------------------------------
Public Function NormalizeAccountIdentifier(ByVal varValue As Variant, ByVal strDisplayedText As String, ByVal lngBusinessError As Long, ByVal strContext As String) As String
    Const METHOD_NAME As String = "NormalizeAccountIdentifier"
    Dim dblValue As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim strDisplayed As String
    Dim strNumeric As String
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(varValue) Then Call VBA.Err.Raise(lngBusinessError, METHOD_NAME, "An Excel error value cannot be used as an Account identifier. " & strContext)

    If IsNull(varValue) Or IsEmpty(varValue) Then GoTo ExitPoint

    If IsNumeric(varValue) Then
        dblValue = CDbl(varValue)

        If dblValue < 0 Or dblValue <> Fix(dblValue) Then Call VBA.Err.Raise(lngBusinessError, METHOD_NAME, "A numeric Account identifier must be a non-negative whole number. Store the Account as text if it contains special characters. " & strContext)

        strNumeric = Format$(dblValue, "0")

        If Len(strNumeric) > 15 Then Call VBA.Err.Raise(lngBusinessError, METHOD_NAME, "A numeric Account identifier contains more than 15 digits. Excel may already have lost precision. Store the Account as text in the source workbook. " & strContext)

        strDisplayed = Trim$(strDisplayedText)

        If IsDigitsOnly(strDisplayed) Then
            If Len(strDisplayed) >= Len(strNumeric) Then
                If Right$(strDisplayed, Len(strNumeric)) = strNumeric Then
                    strResult = strDisplayed
                End If
            End If
        End If

        If Len(strResult) = 0 Then strResult = strNumeric
    Else
        strResult = Trim$(CStr(varValue))
    End If

ExitPoint:
    If errNumber = 0 Then NormalizeAccountIdentifier = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "context", strContext)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strValue As String
' Returns:       Boolean
' Description:   Checks whether a string contains digits only.
'-------------------------------------------------------------------------------
Private Function IsDigitsOnly(ByVal strValue As String) As Boolean
    Const METHOD_NAME As String = "IsDigitsOnly"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPosition As Long
    Dim strCharacter As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Len(strValue) = 0 Then GoTo ExitPoint

    blnResult = True

    For lngPosition = 1 To Len(strValue)
        strCharacter = Mid$(strValue, lngPosition, 1)

        If strCharacter < "0" Or strCharacter > "9" Then
            blnResult = False
            Exit For
        End If
    Next lngPosition

ExitPoint:
    If errNumber = 0 Then IsDigitsOnly = blnResult
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
' Parameters:    arrData As Variant; lngRow As Long
' Returns:       String
' Description:   Builds the required historical deduplication key.
'-------------------------------------------------------------------------------
Public Function BuildTransactionKey(ByRef arrData As Variant, ByVal lngRow As Long) As String
    Const METHOD_NAME As String = "BuildTransactionKey"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = GetKeyPart(arrData(lngRow, 6))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 8))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 9))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 12))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 13))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 14))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 15))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 16))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 18))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 19))
    strResult = strResult & "|" & GetKeyPart(arrData(lngRow, 20))

ExitPoint:
    If errNumber = 0 Then BuildTransactionKey = strResult
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
' Parameters:    varValue As Variant
' Returns:       String
' Description:   Converts one historical key component using CStr semantics.
'-------------------------------------------------------------------------------
Private Function GetKeyPart(ByVal varValue As Variant) As String
    Const METHOD_NAME As String = "GetKeyPart"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(varValue) Then Call VBA.Err.Raise(ERROR_HISTORICAL, METHOD_NAME, "An Excel error value cannot be used in a historical key.")

    If IsNull(varValue) Or IsEmpty(varValue) Then
        strResult = vbNullString
    Else
        strResult = CStr(varValue)
    End If

ExitPoint:
    If errNumber = 0 Then GetKeyPart = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function
