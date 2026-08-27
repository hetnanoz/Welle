Option Explicit

Private Const CLASS_NAME As String = "modFileSystem"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strPath As String
' Returns:       String
' Description:   Resolves local, UNC, WebDAV, or HTTPS source paths.
'-------------------------------------------------------------------------------
Public Function ResolveConfiguredPath(ByVal strPath As String) As String
    Const METHOD_NAME As String = "ResolveConfiguredPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = Trim$(strPath)
    If Len(strResult) = 0 Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "A configured path is blank.")

    If IsHttpPath(strResult) Then
        strResult = ConvertHttpUrlToWebDav(strResult)
    Else
        strResult = NormalizeBackslashes(strResult)
    End If

    Do While Len(strResult) > 3 And Right$(strResult, 1) = "\"
        strResult = Left$(strResult, Len(strResult) - 1)
    Loop

ExitPoint:
    If errNumber = 0 Then ResolveConfiguredPath = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "path", strPath)
    GoTo ExitPoint
End Function
'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strPath As String
' Returns:       Boolean
' Description:   Returns True when the value is an HTTP or HTTPS URL.
'-------------------------------------------------------------------------------
Public Function IsHttpPath(ByVal strPath As String) As Boolean
    Const METHOD_NAME As String = "IsHttpPath"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strValue = LCase$(Trim$(strPath))
    blnResult = Left$(strValue, 8) = "https://" Or Left$(strValue, 7) = "http://"

ExitPoint:
    If errNumber = 0 Then IsHttpPath = blnResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "path", strPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strPath As String
' Returns:       String
' Description:   Resolves output paths while preserving clean SharePoint HTTPS URLs.
'-------------------------------------------------------------------------------
Public Function ResolveOutputPath(ByVal strPath As String) As String
    Const METHOD_NAME As String = "ResolveOutputPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsHttpPath(strPath) Then
        strResult = NormalizeHttpUrl(strPath)
    Else
        strResult = ResolveConfiguredPath(strPath)
    End If

ExitPoint:
    If errNumber = 0 Then ResolveOutputPath = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "path", strPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strUrl As String
' Returns:       String
' Description:   Normalizes a clean HTTP/HTTPS SharePoint URL without converting it to WebDAV.
'-------------------------------------------------------------------------------
Private Function NormalizeHttpUrl(ByVal strUrl As String) As String
    Const METHOD_NAME As String = "NormalizeHttpUrl"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngFragmentPosition As Long
    Dim lngQueryPosition As Long
    Dim lngSchemePosition As Long
    Dim strPrefix As String
    Dim strRemainder As String
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = Trim$(strUrl)
    strResult = Replace$(strResult, "\", "/")

    If Not IsHttpPath(strResult) Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "The output URL must start with http:// or https://.")
    If InStr(1, LCase$(strResult), "/forms/allitems.aspx", vbBinaryCompare) > 0 Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "Use a clean SharePoint folder URL, not a Forms/AllItems.aspx browser-view URL.")
    If InStr(1, LCase$(strResult), "/shared?", vbBinaryCompare) > 0 Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "Use a clean SharePoint folder URL, not a /shared?id= sharing URL.")

    lngQueryPosition = InStr(1, strResult, "?", vbBinaryCompare)
    If lngQueryPosition > 0 Then strResult = Left$(strResult, lngQueryPosition - 1)

    lngFragmentPosition = InStr(1, strResult, "#", vbBinaryCompare)
    If lngFragmentPosition > 0 Then strResult = Left$(strResult, lngFragmentPosition - 1)

    lngSchemePosition = InStr(1, strResult, "://", vbBinaryCompare)
    strPrefix = Left$(strResult, lngSchemePosition + 2)
    strRemainder = Mid$(strResult, lngSchemePosition + 3)

    Do While InStr(1, strRemainder, "//", vbBinaryCompare) > 0
        strRemainder = Replace$(strRemainder, "//", "/")
    Loop

    strResult = strPrefix & strRemainder
    strResult = Replace$(strResult, " ", "%20")

    Do While Len(strResult) > Len(strPrefix) And Right$(strResult, 1) = "/"
        strResult = Left$(strResult, Len(strResult) - 1)
    Loop

ExitPoint:
    If errNumber = 0 Then NormalizeHttpUrl = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "url", strUrl)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strUrl As String
' Returns:       String
' Description:   Converts an HTTP or HTTPS SharePoint URL to a WebDAV UNC path.
'-------------------------------------------------------------------------------
Private Function ConvertHttpUrlToWebDav(ByVal strUrl As String) As String
    Const METHOD_NAME As String = "ConvertHttpUrlToWebDav"
    Dim blnUseSsl As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngFragmentPosition As Long
    Dim lngPathPosition As Long
    Dim lngQueryPosition As Long
    Dim lngSchemePosition As Long
    Dim strHost As String
    Dim strHostAndPath As String
    Dim strPathPart As String
    Dim strResult As String
    Dim strWorkingUrl As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strWorkingUrl = Trim$(strUrl)
    blnUseSsl = LCase$(Left$(strWorkingUrl, 8)) = "https://"

    lngQueryPosition = InStr(1, strWorkingUrl, "?", vbBinaryCompare)
    If lngQueryPosition > 0 Then strWorkingUrl = Left$(strWorkingUrl, lngQueryPosition - 1)

    lngFragmentPosition = InStr(1, strWorkingUrl, "#", vbBinaryCompare)
    If lngFragmentPosition > 0 Then strWorkingUrl = Left$(strWorkingUrl, lngFragmentPosition - 1)

    lngSchemePosition = InStr(1, strWorkingUrl, "://", vbBinaryCompare)
    If lngSchemePosition = 0 Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "Invalid SharePoint URL.")

    strHostAndPath = Mid$(strWorkingUrl, lngSchemePosition + 3)
    lngPathPosition = InStr(1, strHostAndPath, "/", vbBinaryCompare)

    If lngPathPosition > 0 Then
        strHost = Left$(strHostAndPath, lngPathPosition - 1)
        strPathPart = Mid$(strHostAndPath, lngPathPosition + 1)
    Else
        strHost = strHostAndPath
        strPathPart = vbNullString
    End If

    If Len(strHost) = 0 Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "SharePoint URL does not contain a host name.")

    strPathPart = DecodeCommonUrlEscapes(strPathPart)
    strPathPart = NormalizeBackslashes(strPathPart)

    If blnUseSsl Then
        strResult = "\\" & strHost & "@SSL\DavWWWRoot"
    Else
        strResult = "\\" & strHost & "\DavWWWRoot"
    End If

    If Len(strPathPart) > 0 Then strResult = CombinePath(strResult, strPathPart)

ExitPoint:
    If errNumber = 0 Then ConvertHttpUrlToWebDav = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "url", strUrl)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strValue As String
' Returns:       String
' Description:   Decodes common URL characters used in SharePoint paths.
'-------------------------------------------------------------------------------
Private Function DecodeCommonUrlEscapes(ByVal strValue As String) As String
    Const METHOD_NAME As String = "DecodeCommonUrlEscapes"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = strValue
    strResult = Replace$(strResult, "%20", " ", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%23", "#", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%25", "%", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%26", "&", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%27", "'", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%28", "(", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%29", ")", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%2D", "-", 1, -1, vbTextCompare)
    strResult = Replace$(strResult, "%5F", "_", 1, -1, vbTextCompare)

ExitPoint:
    If errNumber = 0 Then DecodeCommonUrlEscapes = strResult
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
' Parameters:    strPath As String
' Returns:       String
' Description:   Normalizes slash direction while preserving a UNC prefix.
'-------------------------------------------------------------------------------
Private Function NormalizeBackslashes(ByVal strPath As String) As String
    Const METHOD_NAME As String = "NormalizeBackslashes"
    Dim blnUncPath As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim strRemainder As String
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = Replace$(Trim$(strPath), "/", "\")
    blnUncPath = Left$(strResult, 2) = "\\"

    If blnUncPath Then
        strRemainder = Mid$(strResult, 3)

        Do While InStr(1, strRemainder, "\\", vbBinaryCompare) > 0
            strRemainder = Replace$(strRemainder, "\\", "\")
        Loop

        strResult = "\\" & strRemainder
    Else
        Do While InStr(1, strResult, "\\", vbBinaryCompare) > 0
            strResult = Replace$(strResult, "\\", "\")
        Loop
    End If

ExitPoint:
    If errNumber = 0 Then NormalizeBackslashes = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "path", strPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strLeft As String; strRight As String
' Returns:       String
' Description:   Combines Windows paths or clean HTTP/HTTPS SharePoint URL components.
'-------------------------------------------------------------------------------
Public Function CombinePath(ByVal strLeft As String, ByVal strRight As String) As String
    Const METHOD_NAME As String = "CombinePath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strLeftPart As String
    Dim strResult As String
    Dim strRightPart As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsHttpPath(strLeft) Then
        strLeftPart = NormalizeHttpUrl(strLeft)
        strRightPart = Replace$(Trim$(strRight), "\", "/")

        Do While Len(strRightPart) > 0 And Left$(strRightPart, 1) = "/"
            strRightPart = Mid$(strRightPart, 2)
        Loop

        strRightPart = Replace$(strRightPart, " ", "%20")

        If Len(strRightPart) = 0 Then
            strResult = strLeftPart
        Else
            strResult = strLeftPart & "/" & strRightPart
        End If
    Else
        strLeftPart = NormalizeBackslashes(strLeft)
        strRightPart = NormalizeBackslashes(strRight)

        Do While Len(strLeftPart) > 3 And Right$(strLeftPart, 1) = "\"
            strLeftPart = Left$(strLeftPart, Len(strLeftPart) - 1)
        Loop

        Do While Len(strRightPart) > 0 And Left$(strRightPart, 1) = "\"
            strRightPart = Mid$(strRightPart, 2)
        Loop

        If Len(strRightPart) = 0 Then
            strResult = strLeftPart
        ElseIf Right$(strLeftPart, 1) = "\" Then
            strResult = strLeftPart & strRightPart
        Else
            strResult = strLeftPart & "\" & strRightPart
        End If
    End If

ExitPoint:
    If errNumber = 0 Then CombinePath = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "leftPath;rightPath", strLeft, strRight)
    GoTo ExitPoint
End Function
'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strFolderPath As String
' Returns:       ---
' Description:   Creates local/UNC/WebDAV folders; clean HTTPS SharePoint folders must already exist.
'-------------------------------------------------------------------------------
Public Sub EnsureFolderExists(ByVal strFolderPath As String)
    Const METHOD_NAME As String = "EnsureFolderExists"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object
    Dim strParentPath As String
    Dim strResolvedPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsHttpPath(strFolderPath) Then GoTo ExitPoint

    strResolvedPath = ResolveConfiguredPath(strFolderPath)
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")

    If objFileSystem.FolderExists(strResolvedPath) Then GoTo ExitPoint

    strParentPath = objFileSystem.GetParentFolderName(strResolvedPath)

    If Len(strParentPath) = 0 Or StrComp(strParentPath, strResolvedPath, vbTextCompare) = 0 Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "The folder cannot be created because its parent path is unavailable: " & strResolvedPath)
    If Not objFileSystem.FolderExists(strParentPath) Then Call EnsureFolderExists(strParentPath)

    Call objFileSystem.CreateFolder(strResolvedPath)

ExitPoint:
    Set objFileSystem = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "folderPath", strFolderPath)
    GoTo ExitPoint
End Sub
'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    dtRunTimestamp As Date
' Returns:       String
' Description:   Creates a unique temporary workspace for one macro run.
'-------------------------------------------------------------------------------
Public Function CreateRunWorkspace(ByVal dtRunTimestamp As Date) As String
    Const METHOD_NAME As String = "CreateRunWorkspace"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngTimerToken As Long
    Dim strBasePath As String
    Dim strResult As String
    Dim strRootPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strBasePath = Trim$(Environ$("TEMP"))
    If Len(strBasePath) = 0 Then strBasePath = ThisWorkbook.Path
    If Len(strBasePath) = 0 Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "A temporary base folder could not be determined.")

    strRootPath = CombinePath(strBasePath, "DallasCashTransactions")
    Call EnsureFolderExists(strRootPath)

    lngTimerToken = CLng(Timer * 1000) Mod 1000000
    strResult = CombinePath(strRootPath, Format$(dtRunTimestamp, "yyyymmdd_hhnnss") & "_" & Format$(lngTimerToken, "000000"))
    Call EnsureFolderExists(strResult)

ExitPoint:
    If errNumber = 0 Then CreateRunWorkspace = strResult
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
' Parameters:    strRootPath As String; dtReportDate As Date
' Returns:       String
' Description:   Builds the YYYY\MM\YYYY-MM-DD daily output path.
'-------------------------------------------------------------------------------
Public Function BuildDatedFolderPath(ByVal strRootPath As String, ByVal dtReportDate As Date) As String
    Const METHOD_NAME As String = "BuildDatedFolderPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = CombinePath(strRootPath, Format$(dtReportDate, "yyyy"))
    strResult = CombinePath(strResult, Format$(dtReportDate, "mm"))
    strResult = CombinePath(strResult, Format$(dtReportDate, "yyyy-mm-dd"))

ExitPoint:
    If errNumber = 0 Then BuildDatedFolderPath = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "rootPath;reportDate", strRootPath, dtReportDate)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strSourcePath As String; strDestinationFolder As String
' Returns:       ---
' Description:   Copies an Excel file to a Windows folder or directly to a clean SharePoint HTTPS URL.
'-------------------------------------------------------------------------------
Public Sub CopyFileToFolder(ByVal strSourcePath As String, ByVal strDestinationFolder As String)
    Const METHOD_NAME As String = "CopyFileToFolder"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object
    Dim strDestinationPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Call EnsureFolderExists(strDestinationFolder)
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")

    If Not objFileSystem.FileExists(strSourcePath) Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "Source file does not exist: " & strSourcePath)

    If IsHttpPath(strDestinationFolder) Then
        Call CopyExcelWorkbookToHttpFolder(strSourcePath, strDestinationFolder)
    Else
        strDestinationPath = CombinePath(strDestinationFolder, objFileSystem.GetFileName(strSourcePath))
        Call objFileSystem.CopyFile(strSourcePath, strDestinationPath, True)
    End If

ExitPoint:
    Set objFileSystem = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "sourcePath;destinationFolder", strSourcePath, strDestinationFolder)
    GoTo ExitPoint
End Sub
'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strSourcePath As String; strDestinationFolder As String
' Returns:       ---
' Description:   Publishes a local Excel workbook to an existing SharePoint HTTPS folder using Excel SaveCopyAs.
'-------------------------------------------------------------------------------
Private Sub CopyExcelWorkbookToHttpFolder(ByVal strSourcePath As String, ByVal strDestinationFolder As String)
    Const METHOD_NAME As String = "CopyExcelWorkbookToHttpFolder"
    Dim blnCloseRequired As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object
    Dim strDestinationPath As String
    Dim strExtension As String
    Dim wkbAlreadyOpen As Excel.Workbook
    Dim wkbSource As Excel.Workbook

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    strExtension = GetFileExtension(strSourcePath)

    If strExtension <> ".xls" And strExtension <> ".xlsx" Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "HTTPS publishing currently supports Excel .xls and .xlsx files only: " & strSourcePath)

    strDestinationPath = CombinePath(strDestinationFolder, objFileSystem.GetFileName(strSourcePath))
    Set wkbAlreadyOpen = GetOpenWorkbookByFullName(strSourcePath, False)

    If wkbAlreadyOpen Is Nothing Then
        Set wkbSource = Application.Workbooks.Open(Filename:=strSourcePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
        blnCloseRequired = True
    Else
        Set wkbSource = wkbAlreadyOpen
    End If

    If wkbSource Is Nothing Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "The staged Excel workbook could not be opened for HTTPS publishing: " & strSourcePath)

    Call wkbSource.SaveCopyAs(strDestinationPath)

ExitPoint:
    Set wkbAlreadyOpen = Nothing
    Set objFileSystem = Nothing

    If blnCloseRequired Then
        blnCloseRequired = False
        Call wkbSource.Close(SaveChanges:=False)
    End If

    Set wkbSource = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = "HTTPS SharePoint upload failed. The target SharePoint folder must already exist and the current Office user must have write access. Target: " & strDestinationPath & ". Original error: " & VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "sourcePath;destinationFolder;destinationPath", strSourcePath, strDestinationFolder, strDestinationPath)
    GoTo ExitPoint
End Sub
'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFilePath As String
' Returns:       ---
' Description:   Deletes a file when it exists.
'-------------------------------------------------------------------------------
Public Sub DeleteFileIfExists(ByVal strFilePath As String)
    Const METHOD_NAME As String = "DeleteFileIfExists"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    If objFileSystem.FileExists(strFilePath) Then Call objFileSystem.DeleteFile(strFilePath, True)

ExitPoint:
    Set objFileSystem = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "filePath", strFilePath)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFolderPath As String
' Returns:       ---
' Description:   Deletes a folder recursively when it exists.
'-------------------------------------------------------------------------------
Public Sub DeleteFolderIfExists(ByVal strFolderPath As String)
    Const METHOD_NAME As String = "DeleteFolderIfExists"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    If objFileSystem.FolderExists(strFolderPath) Then Call objFileSystem.DeleteFolder(strFolderPath, True)

ExitPoint:
    Set objFileSystem = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "folderPath", strFolderPath)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFilePath As String
' Returns:       Boolean
' Description:   Checks whether a file exists.
'-------------------------------------------------------------------------------
Public Function FileExists(ByVal strFilePath As String) As Boolean
    Const METHOD_NAME As String = "FileExists"
    Dim blnResult As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    blnResult = objFileSystem.FileExists(strFilePath)

ExitPoint:
    Set objFileSystem = Nothing
    If errNumber = 0 Then FileExists = blnResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "filePath", strFilePath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFolderPath As String; strFilePattern As String
' Returns:       String
' Description:   Returns the lexicographically latest matching file name.
'-------------------------------------------------------------------------------
Public Function GetLatestFile(ByVal strFolderPath As String, ByVal strFilePattern As String) As String
    Const METHOD_NAME As String = "GetLatestFile"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objFileSystem As Object
    Dim strFileName As String
    Dim strLatestFile As String
    Dim strMaximumName As String
    Dim strResolvedFolder As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResolvedFolder = ResolveConfiguredPath(strFolderPath)
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")

    If Not objFileSystem.FolderExists(strResolvedFolder) Then Call VBA.Err.Raise(ERROR_FILE_SYSTEM, METHOD_NAME, "Configured folder does not exist: " & strResolvedFolder)

    strFileName = Dir$(CombinePath(strResolvedFolder, strFilePattern))

    Do While Len(strFileName) > 0
        If UCase$(strFileName) > UCase$(strMaximumName) Then
            strMaximumName = strFileName
            strLatestFile = CombinePath(strResolvedFolder, strFileName)
        End If

        strFileName = Dir$()
    Loop

ExitPoint:
    Set objFileSystem = Nothing
    If errNumber = 0 Then GetLatestFile = strLatestFile
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "folderPath;filePattern", strFolderPath, strFilePattern)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    strPath As String
' Returns:       String
' Description:   Normalizes a workbook path or SharePoint HTTPS URL for comparison.
'-------------------------------------------------------------------------------
Public Function NormalizeWorkbookPath(ByVal strPath As String) As String
    Const METHOD_NAME As String = "NormalizeWorkbookPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsHttpPath(strPath) Then
        strResult = NormalizeHttpUrl(strPath)
    Else
        strResult = ResolveConfiguredPath(strPath)
        strResult = NormalizeBackslashes(strResult)

        Do While Len(strResult) > 3 And Right$(strResult, 1) = "\"
            strResult = Left$(strResult, Len(strResult) - 1)
        Loop
    End If

    strResult = LCase$(strResult)

ExitPoint:
    If errNumber = 0 Then NormalizeWorkbookPath = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "path", strPath)
    GoTo ExitPoint
End Function
'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFullPath As String
' Returns:       String
' Description:   Extracts a file name from a full path or URL.
'-------------------------------------------------------------------------------
Public Function GetFileNameFromPath(ByVal strFullPath As String) As String
    Const METHOD_NAME As String = "GetFileNameFromPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPosition As Long
    Dim strResult As String
    Dim strWorkingPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strWorkingPath = Replace$(strFullPath, "/", "\")
    lngPosition = InStrRev(strWorkingPath, "\")

    If lngPosition > 0 Then
        strResult = Mid$(strWorkingPath, lngPosition + 1)
    Else
        strResult = strWorkingPath
    End If

ExitPoint:
    If errNumber = 0 Then GetFileNameFromPath = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "fullPath", strFullPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFullName As String; blnAllowNameFallback As Boolean
' Returns:       Excel.Workbook
' Description:   Finds an open workbook by normalized path and optional name fallback.
'-------------------------------------------------------------------------------
Public Function GetOpenWorkbookByFullName(ByVal strFullName As String, Optional ByVal blnAllowNameFallback As Boolean = True) As Excel.Workbook
    Const METHOD_NAME As String = "GetOpenWorkbookByFullName"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strCandidateFullName As String
    Dim strTargetFileName As String
    Dim strTargetFullName As String
    Dim wkbCandidate As Excel.Workbook
    Dim wkbNameMatch As Excel.Workbook
    Dim wkbResult As Excel.Workbook

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strTargetFullName = NormalizeWorkbookPath(strFullName)
    strTargetFileName = GetFileNameFromPath(strFullName)

    For Each wkbCandidate In Application.Workbooks
        If StrComp(wkbCandidate.Name, strTargetFileName, vbTextCompare) = 0 Then
            If wkbNameMatch Is Nothing Then Set wkbNameMatch = wkbCandidate

            strCandidateFullName = NormalizeWorkbookPath(wkbCandidate.FullName)

            If StrComp(strCandidateFullName, strTargetFullName, vbTextCompare) = 0 Then
                Set wkbResult = wkbCandidate
                Exit For
            End If
        End If
    Next wkbCandidate

    If wkbResult Is Nothing And blnAllowNameFallback Then Set wkbResult = wkbNameMatch

ExitPoint:
    Set wkbCandidate = Nothing
    Set wkbNameMatch = Nothing
    If errNumber = 0 Then Set GetOpenWorkbookByFullName = wkbResult
    Set wkbResult = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "fullName;allowNameFallback", strFullName, blnAllowNameFallback)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strFileName As String
' Returns:       String
' Description:   Returns the lower-case extension including the leading dot.
'-------------------------------------------------------------------------------
Public Function GetFileExtension(ByVal strFileName As String) As String
    Const METHOD_NAME As String = "GetFileExtension"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPosition As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngPosition = InStrRev(strFileName, ".")
    If lngPosition > 0 Then strResult = LCase$(Mid$(strFileName, lngPosition))

ExitPoint:
    If errNumber = 0 Then GetFileExtension = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "fileName", strFileName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strValue As String
' Returns:       String
' Description:   Removes characters that are invalid in Windows file names.
'-------------------------------------------------------------------------------
Public Function SanitizeFileNamePart(ByVal strValue As String) As String
    Const METHOD_NAME As String = "SanitizeFileNamePart"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = Trim$(strValue)
    strResult = Replace$(strResult, "\", "_")
    strResult = Replace$(strResult, "/", "_")
    strResult = Replace$(strResult, ":", "_")
    strResult = Replace$(strResult, "*", "_")
    strResult = Replace$(strResult, "?", "_")
    strResult = Replace$(strResult, Chr$(34), "_")
    strResult = Replace$(strResult, "<", "_")
    strResult = Replace$(strResult, ">", "_")
    strResult = Replace$(strResult, "|", "_")

    Do While Len(strResult) > 0 And (Right$(strResult, 1) = "." Or Right$(strResult, 1) = " ")
        strResult = Left$(strResult, Len(strResult) - 1)
    Loop

    If Len(strResult) = 0 Then strResult = "UNNAMED"

ExitPoint:
    If errNumber = 0 Then SanitizeFileNamePart = strResult
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
' Parameters:    appConfig As TAppConfig; dtRunTimestamp As Date; strFallbackWorkspace As String
' Returns:       String
' Description:   Creates the working Input folder, preferring the enabled local daily output folder.
'-------------------------------------------------------------------------------
Public Function CreateInputWorkingFolder(ByRef appConfig As TAppConfig, ByVal dtRunTimestamp As Date, ByVal strFallbackWorkspace As String) As String
    Const METHOD_NAME As String = "CreateInputWorkingFolder"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strDailyFolder As String
    Dim strResolvedLocalRoot As String
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If appConfig.SaveLocal Then
        strResolvedLocalRoot = ResolveConfiguredPath(appConfig.OutputLocalBase)
        strDailyFolder = BuildDatedFolderPath(strResolvedLocalRoot, DateValue(dtRunTimestamp))
        strResult = CombinePath(strDailyFolder, "Input")
    Else
        strResult = CombinePath(strFallbackWorkspace, "Input")
    End If

    Call EnsureFolderExists(strResult)

ExitPoint:
    If errNumber = 0 Then CreateInputWorkingFolder = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "saveLocal;localBase;fallbackWorkspace", appConfig.SaveLocal, appConfig.OutputLocalBase, strFallbackWorkspace)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    colAttachmentPaths As Collection; colDestinationRoots As Collection; dtReportDate As Date; strSourceInputFolder As String
' Returns:       ---
' Description:   Publishes downloaded source files to an Input subfolder in every enabled daily output location.
'-------------------------------------------------------------------------------
Public Sub PublishInputFilesToDailyFolders(ByVal colAttachmentPaths As Collection, ByVal colDestinationRoots As Collection, ByVal dtReportDate As Date, ByVal strSourceInputFolder As String)
    Const METHOD_NAME As String = "PublishInputFilesToDailyFolders"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngFile As Long
    Dim lngRoot As Long
    Dim strDestinationInputFolder As String
    Dim strSourceFile As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If colAttachmentPaths Is Nothing Then GoTo ExitPoint
    If colDestinationRoots Is Nothing Then GoTo ExitPoint

    For lngRoot = 1 To colDestinationRoots.Count
        strDestinationInputFolder = BuildDatedFolderPath(CStr(colDestinationRoots(lngRoot)), dtReportDate)
        strDestinationInputFolder = CombinePath(strDestinationInputFolder, "Input")

        If StrComp(NormalizeWorkbookPath(strDestinationInputFolder), NormalizeWorkbookPath(strSourceInputFolder), vbTextCompare) <> 0 Then
            Call EnsureFolderExists(strDestinationInputFolder)

            For lngFile = 1 To colAttachmentPaths.Count
                strSourceFile = CStr(colAttachmentPaths(lngFile))
                Call CopyFileToFolder(strSourceFile, strDestinationInputFolder)
            Next lngFile
        End If
    Next lngRoot

ExitPoint:
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "sourceInputFolder;destinationInputFolder;sourceFile", strSourceInputFolder, strDestinationInputFolder, strSourceFile)
    GoTo ExitPoint
End Sub

