Option Explicit

Private Const CLASS_NAME As String = "modOutlook"
Private Const OL_FOLDER_INBOX As Long = 6
Private Const OL_MAIL_ITEM_CLASS As Long = 43

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    appConfig As TAppConfig; strWorkspace As String
' Returns:       Collection
' Description:   Saves matching Excel attachments and archives processed mail.
'-------------------------------------------------------------------------------
Public Function DownloadCashTransactionAttachments(ByRef appConfig As TAppConfig, ByVal strWorkspace As String) As Collection
    Const METHOD_NAME As String = "DownloadCashTransactionAttachments"
    Dim colAttachmentPaths As Collection
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngAttachment As Long
    Dim lngItem As Long
    Dim lngSavedForMail As Long
    Dim objArchiveFolder As Object
    Dim objAttachment As Object
    Dim objItem As Object
    Dim objItems As Object
    Dim objMail As Object
    Dim objMovedMail As Object
    Dim objNamespace As Object
    Dim objOutlook As Object
    Dim objRestrictedItems As Object
    Dim objRootFolder As Object
    Dim objSourceFolder As Object
    Dim strExtension As String
    Dim strFilter As String
    Dim strSafeFileName As String
    Dim strTargetPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set colAttachmentPaths = New Collection
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objRootFolder = GetMailboxRootFolder(objNamespace, appConfig.OutlookMailbox)
    Set objSourceFolder = GetFolderByPath(objRootFolder, appConfig.OutlookSourceFolder)
    Set objArchiveFolder = GetFolderByPath(objRootFolder, appConfig.OutlookArchiveFolder)

    If StrComp(CStr(objSourceFolder.EntryID), CStr(objArchiveFolder.EntryID), vbBinaryCompare) = 0 Then Call VBA.Err.Raise(ERROR_OUTLOOK, METHOD_NAME, "Outlook source and archive folders resolve to the same folder.")

    Set objItems = objSourceFolder.Items
    strFilter = BuildSubjectFilter(OUTLOOK_SUBJECT_PHRASE)
    Set objRestrictedItems = objItems.Restrict(strFilter)
    Call objRestrictedItems.Sort("[ReceivedTime]", False)

    For lngItem = objRestrictedItems.Count To 1 Step -1
        Set objItem = objRestrictedItems.Item(lngItem)

        If CLng(objItem.Class) = OL_MAIL_ITEM_CLASS Then
            Set objMail = objItem

            If InStr(1, CStr(objMail.Subject), OUTLOOK_SUBJECT_PHRASE, vbTextCompare) > 0 Then
                lngSavedForMail = 0

                For lngAttachment = 1 To objMail.Attachments.Count
                    Set objAttachment = objMail.Attachments.Item(lngAttachment)
                    strExtension = GetFileExtension(CStr(objAttachment.FileName))

                    If strExtension = ".xlsx" Or strExtension = ".xls" Then
                        strSafeFileName = Format$(CDate(objMail.ReceivedTime), "yyyymmdd_hhnnss") & "_" & Format$(lngItem, "0000") & "_" & Format$(lngAttachment, "00") & "_" & SanitizeFileNamePart(CStr(objAttachment.FileName))
                        strTargetPath = CombinePath(strWorkspace, strSafeFileName)
                        Call objAttachment.SaveAsFile(strTargetPath)
                        colAttachmentPaths.Add strTargetPath
                        lngSavedForMail = lngSavedForMail + 1
                    End If

                    Set objAttachment = Nothing
                Next lngAttachment

                If lngSavedForMail > 0 Then
                    Set objMovedMail = objMail.Move(objArchiveFolder)
                    Set objMovedMail = Nothing
                End If
            End If
        End If

        Set objMail = Nothing
        Set objItem = Nothing
    Next lngItem

    If colAttachmentPaths.Count = 0 Then Call VBA.Err.Raise(ERROR_OUTLOOK, METHOD_NAME, "No .xlsx or .xls attachments were found in matching Outlook messages.")

ExitPoint:
    Set objMovedMail = Nothing
    Set objAttachment = Nothing
    Set objMail = Nothing
    Set objItem = Nothing
    Set objRestrictedItems = Nothing
    Set objItems = Nothing
    Set objArchiveFolder = Nothing
    Set objSourceFolder = Nothing
    Set objRootFolder = Nothing
    Set objNamespace = Nothing
    Set objOutlook = Nothing

    If errNumber = 0 Then Set DownloadCashTransactionAttachments = colAttachmentPaths
    Set colAttachmentPaths = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "mailbox;sourceFolder;archiveFolder;workspace", appConfig.OutlookMailbox, appConfig.OutlookSourceFolder, appConfig.OutlookArchiveFolder, strWorkspace)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    objNamespace As Object; strMailbox As String
' Returns:       Object
' Description:   Resolves a shared mailbox root by store name or recipient.
'-------------------------------------------------------------------------------
Private Function GetMailboxRootFolder(ByVal objNamespace As Object, ByVal strMailbox As String) As Object
    Const METHOD_NAME As String = "GetMailboxRootFolder"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objCandidateFolder As Object
    Dim objInbox As Object
    Dim objRecipient As Object
    Dim objResult As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each objCandidateFolder In objNamespace.Folders
        If StrComp(Trim$(CStr(objCandidateFolder.Name)), Trim$(strMailbox), vbTextCompare) = 0 Then
            Set objResult = objCandidateFolder
            Exit For
        End If
    Next objCandidateFolder

    If objResult Is Nothing Then
        Set objRecipient = objNamespace.CreateRecipient(strMailbox)

        If Not CBool(objRecipient.Resolve) Then Call VBA.Err.Raise(ERROR_OUTLOOK, METHOD_NAME, "The shared mailbox could not be resolved: " & strMailbox)

        Set objInbox = objNamespace.GetSharedDefaultFolder(objRecipient, OL_FOLDER_INBOX)
        Set objResult = objInbox.Parent
    End If

    If objResult Is Nothing Then Call VBA.Err.Raise(ERROR_OUTLOOK, METHOD_NAME, "The shared mailbox root could not be accessed: " & strMailbox)

ExitPoint:
    Set objRecipient = Nothing
    Set objInbox = Nothing
    Set objCandidateFolder = Nothing
    If errNumber = 0 Then Set GetMailboxRootFolder = objResult
    Set objResult = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "mailbox", strMailbox)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    objRootFolder As Object; strFolderPath As String
' Returns:       Object
' Description:   Navigates an Outlook folder path below a mailbox root.
'-------------------------------------------------------------------------------
Private Function GetFolderByPath(ByVal objRootFolder As Object, ByVal strFolderPath As String) As Object
    Const METHOD_NAME As String = "GetFolderByPath"
    Dim arrParts As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPart As Long
    Dim objCurrentFolder As Object
    Dim objNextFolder As Object
    Dim strPart As String
    Dim strWorkingPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strWorkingPath = Replace$(Trim$(strFolderPath), "/", "\")
    arrParts = Split(strWorkingPath, "\")
    Set objCurrentFolder = objRootFolder

    For lngPart = LBound(arrParts) To UBound(arrParts)
        strPart = Trim$(CStr(arrParts(lngPart)))

        If Len(strPart) > 0 Then
            If StrComp(strPart, CStr(objRootFolder.Name), vbTextCompare) <> 0 Then
                Set objNextFolder = objCurrentFolder.Folders.Item(strPart)
                Set objCurrentFolder = objNextFolder
                Set objNextFolder = Nothing
            End If
        End If
    Next lngPart

ExitPoint:
    Set objNextFolder = Nothing
    If errNumber = 0 Then Set GetFolderByPath = objCurrentFolder
    Set objCurrentFolder = Nothing
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "folderPath", strFolderPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-26
' Parameters:    strSubjectPhrase As String
' Returns:       String
' Description:   Builds a DASL case-insensitive phrase filter for Subject.
'-------------------------------------------------------------------------------
Private Function BuildSubjectFilter(ByVal strSubjectPhrase As String) As String
    Const METHOD_NAME As String = "BuildSubjectFilter"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strEscapedPhrase As String
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strEscapedPhrase = Replace$(strSubjectPhrase, "'", "''")
    strResult = "@SQL=" & Chr$(34) & "urn:schemas:httpmail:subject" & Chr$(34) & " LIKE '%" & strEscapedPhrase & "%'"

ExitPoint:
    If errNumber = 0 Then BuildSubjectFilter = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function
