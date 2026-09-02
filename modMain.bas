Option Explicit

Private Const CLASS_NAME As String = "modMain"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    ---
' Returns:       ---
' Description:   Runs the complete Dallas Cash Transactions process.
'-------------------------------------------------------------------------------
Public Sub RunDallasCashTransactions()
    Const METHOD_NAME As String = "RunDallasCashTransactions"
    Dim appConfig As TAppConfig
    Dim arrCombined As Variant
    Dim blnEmailDraftCreated As Boolean
    Dim blnEmailLabelApplied As Boolean
    Dim blnPreviousDisplayAlerts As Boolean
    Dim blnPreviousEnableEvents As Boolean
    Dim blnPreviousScreenUpdating As Boolean
    Dim blnStateCaptured As Boolean
    Dim colAttachmentPaths As Collection
    Dim colDestinationRoots As Collection
    Dim colTeams As Collection
    Dim dictAccounts As Object
    Dim dictAttachmentSubjects As Object
    Dim dictFundTeams As Object
    Dim dictProcessedMailSubjects As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPreviousAutomationSecurity As Long
    Dim strInputFolder As String
    Dim strLatestFondslistePath As String
    Dim strResolvedLocalOutputFolder As String
    Dim strResolvedSharePointOutputFolder As String
    Dim strStageWorkspace As String
    Dim strSuccessMessage As String
    Dim varPreviousStatusBar As Variant
    Dim xlPreviousCalculation As XlCalculation
    Dim dtRunTimestamp As Date

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousScreenUpdating = Application.ScreenUpdating
    blnPreviousEnableEvents = Application.EnableEvents
    blnPreviousDisplayAlerts = Application.DisplayAlerts
    xlPreviousCalculation = Application.Calculation
    varPreviousStatusBar = Application.StatusBar

    If VarType(varPreviousStatusBar) = vbString Then
        If InStr(1, CStr(varPreviousStatusBar), "Dallas Cash Transactions:", vbTextCompare) = 1 Then
            varPreviousStatusBar = False
        End If
    End If

    lngPreviousAutomationSecurity = Application.AutomationSecurity
    blnStateCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = AUTOMATION_SECURITY_FORCE_DISABLE
    Application.StatusBar = "Dallas Cash Transactions: reading configuration..."

    dtRunTimestamp = Now
    Call LoadConfiguration(appConfig)
    strStageWorkspace = CreateRunWorkspace(dtRunTimestamp)
    strInputFolder = CreateInputWorkingFolder(appConfig, dtRunTimestamp, strStageWorkspace)

    Application.StatusBar = "Dallas Cash Transactions: downloading Outlook attachments to daily Input folder..."
    Set colAttachmentPaths = DownloadCashTransactionAttachments(appConfig, strInputFolder, dictAttachmentSubjects, dictProcessedMailSubjects)

    Application.StatusBar = "Dallas Cash Transactions: merging input files from disk..."
    arrCombined = MergeInputFiles(colAttachmentPaths, appConfig.InputWorksheetName, dictAttachmentSubjects)

    Application.StatusBar = "Dallas Cash Transactions: removing duplicate transactions..."
    arrCombined = DeduplicateCurrentTransactions(arrCombined)

    Application.StatusBar = "Dallas Cash Transactions: loading the latest Fondsliste..."
    strLatestFondslistePath = GetLatestFondslistePath(appConfig)
    Set dictFundTeams = BuildOpenFundTeamDictionary(strLatestFondslistePath)

    Application.StatusBar = "Dallas Cash Transactions: updating Static_Data..."
    Call UpdateStaticDataFromFondsliste(dictFundTeams)
    Set dictAccounts = BuildAccountMappingDictionary()
    Call EnrichTransactions(arrCombined, dictAccounts)

    Set colTeams = ParseSupportedTeams(appConfig.SupportedTeams)
    Set colDestinationRoots = BuildDestinationRoots(appConfig)

    If appConfig.SaveLocal Then
        strResolvedLocalOutputFolder = BuildDatedFolderPath(ResolveOutputPath(appConfig.OutputLocalBase), DateValue(dtRunTimestamp))
    End If

    If appConfig.SaveSharePoint Then
        strResolvedSharePointOutputFolder = BuildDatedFolderPath(ResolveOutputPath(appConfig.OutputSharePointBase), DateValue(dtRunTimestamp))
    End If

    Application.StatusBar = "Dallas Cash Transactions: publishing source input files..."
    Call PublishInputFilesToDailyFolders(colAttachmentPaths, colDestinationRoots, DateValue(dtRunTimestamp), strInputFolder)

    Application.StatusBar = "Dallas Cash Transactions: creating daily reports..."
    Call CreateDailyReports(arrCombined, colTeams, colDestinationRoots, dtRunTimestamp, strStageWorkspace)

    If appConfig.UpdateHistorical Then
        Application.StatusBar = "Dallas Cash Transactions: updating historical master files..."
        Call UpdateHistoricalFiles(arrCombined, colTeams, colDestinationRoots)
    End If

    If appConfig.CreateEmailDraft Then
        Application.StatusBar = "Dallas Cash Transactions: creating Outlook email draft..."
        Call CreateCashTransactionsEmailDraft(appConfig, arrCombined, colTeams, dtRunTimestamp, blnEmailLabelApplied)
        blnEmailDraftCreated = True
    End If

    Application.StatusBar = "Dallas Cash Transactions: archiving processed Outlook messages..."
    Call ArchiveProcessedOutlookMessages(appConfig, dictProcessedMailSubjects)

    strSuccessMessage = BuildSuccessMessage(arrCombined, colTeams, colAttachmentPaths.Count, dictProcessedMailSubjects.Count, appConfig.CreateEmailDraft, blnEmailDraftCreated, blnEmailLabelApplied, strResolvedLocalOutputFolder, strResolvedSharePointOutputFolder)

ExitPoint:
    Set dictAccounts = Nothing
    Set dictAttachmentSubjects = Nothing
    Set dictFundTeams = Nothing
    Set dictProcessedMailSubjects = Nothing
    Set colDestinationRoots = Nothing
    Set colTeams = Nothing
    Set colAttachmentPaths = Nothing

    If errNumber = 0 Then
        If Len(strStageWorkspace) > 0 Then Call DeleteFolderIfExists(strStageWorkspace)
    End If

    If blnStateCaptured Then
        Application.StatusBar = varPreviousStatusBar
        Application.AutomationSecurity = lngPreviousAutomationSecurity
        Application.Calculation = xlPreviousCalculation
        Application.DisplayAlerts = blnPreviousDisplayAlerts
        Application.EnableEvents = blnPreviousEnableEvents
        Application.ScreenUpdating = blnPreviousScreenUpdating
    End If

    If errNumber <> 0 Then
        Call ErrorManager.display
    ElseIf Len(strSuccessMessage) > 0 Then
        Call MsgBox(strSuccessMessage, vbInformation, "Dallas Cash Transactions")
    End If

    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "inputFolder;stageWorkspace;fondslistePath", strInputFolder, strStageWorkspace, strLatestFondslistePath)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-27
' Parameters:    arrCombined As Variant; colTeams As Collection; lngAttachmentCount As Long; lngMailCount As Long; blnEmailRequested As Boolean; blnEmailDraftCreated As Boolean; blnEmailLabelApplied As Boolean; strLocalOutputFolder As String; strSharePointOutputFolder As String
' Returns:       String
' Description:   Builds the final success message with transaction counts per team.
'-------------------------------------------------------------------------------
Private Function BuildSuccessMessage(ByRef arrCombined As Variant, ByVal colTeams As Collection, ByVal lngAttachmentCount As Long, ByVal lngMailCount As Long, ByVal blnEmailRequested As Boolean, ByVal blnEmailDraftCreated As Boolean, ByVal blnEmailLabelApplied As Boolean, ByVal strLocalOutputFolder As String, ByVal strSharePointOutputFolder As String) As String
    Const METHOD_NAME As String = "BuildSuccessMessage"
    Dim arrKeys As Variant
    Dim dictConfigured As Object
    Dim dictTeamCounts As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCount As Long
    Dim lngKey As Long
    Dim lngOtherCount As Long
    Dim lngRow As Long
    Dim lngTeam As Long
    Dim lngTransactionCount As Long
    Dim lngUnknownCount As Long
    Dim strOtherTeams As String
    Dim strResult As String
    Dim strTeam As String
    Dim strTeamsWithTrades As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictConfigured = CreateObject("Scripting.Dictionary")
    dictConfigured.CompareMode = vbTextCompare
    Set dictTeamCounts = CreateObject("Scripting.Dictionary")
    dictTeamCounts.CompareMode = vbTextCompare

    lngTransactionCount = UBound(arrCombined, 1) - 1
    lngUnknownCount = CountUnknownFundTransactions(arrCombined)

    For lngRow = 2 To UBound(arrCombined, 1)
        strTeam = Trim$(CStr(arrCombined(lngRow, 5)))
        If Len(strTeam) = 0 Then strTeam = "UNKNOWN TEAM"

        If dictTeamCounts.Exists(strTeam) Then
            dictTeamCounts(strTeam) = CLng(dictTeamCounts(strTeam)) + 1
        Else
            dictTeamCounts.Add strTeam, 1
        End If
    Next lngRow

    For lngTeam = 1 To colTeams.Count
        strTeam = Trim$(CStr(colTeams(lngTeam)))
        If Len(strTeam) > 0 Then dictConfigured(strTeam) = True

        If dictTeamCounts.Exists(strTeam) Then
            lngCount = CLng(dictTeamCounts(strTeam))

            If Len(strTeamsWithTrades) > 0 Then strTeamsWithTrades = strTeamsWithTrades & ", "
            strTeamsWithTrades = strTeamsWithTrades & strTeam & " (" & CStr(lngCount) & ")"
        End If
    Next lngTeam

    If dictTeamCounts.Count > 0 Then
        arrKeys = dictTeamCounts.Keys

        For lngKey = LBound(arrKeys) To UBound(arrKeys)
            strTeam = CStr(arrKeys(lngKey))

            If Not dictConfigured.Exists(strTeam) Then
                lngCount = CLng(dictTeamCounts(strTeam))
                lngOtherCount = lngOtherCount + lngCount

                If Len(strOtherTeams) > 0 Then strOtherTeams = strOtherTeams & ", "
                strOtherTeams = strOtherTeams & strTeam & " (" & CStr(lngCount) & ")"
            End If
        Next lngKey
    End If

    If Len(strTeamsWithTrades) = 0 Then strTeamsWithTrades = "none"

    strResult = "Dallas Cash Transactions completed successfully."
    strResult = strResult & vbCrLf & vbCrLf
    strResult = strResult & "Unique transactions: " & CStr(lngTransactionCount) & vbCrLf
    strResult = strResult & "Excel attachments processed: " & CStr(lngAttachmentCount) & vbCrLf
    strResult = strResult & "Outlook messages archived: " & CStr(lngMailCount) & vbCrLf
    strResult = strResult & "Teams with trades: " & strTeamsWithTrades & vbCrLf
    strResult = strResult & "Unknown fund transactions: " & CStr(lngUnknownCount)

    If blnEmailRequested Then
        strResult = strResult & vbCrLf

        If blnEmailDraftCreated Then
            strResult = strResult & "Outlook email draft: created"
        Else
            strResult = strResult & "Outlook email draft: not created"
        End If

        strResult = strResult & vbCrLf

        If blnEmailLabelApplied Then
            strResult = strResult & "Outlook Internal label: applied"
        Else
            strResult = strResult & "Outlook Internal label: not applied (best effort)"
        End If
    End If

    If lngOtherCount > 0 Then
        strResult = strResult & vbCrLf
        strResult = strResult & "Other / unconfigured Team values: " & strOtherTeams
    End If

    If Len(strLocalOutputFolder) > 0 Then
        strResult = strResult & vbCrLf & vbCrLf
        strResult = strResult & "Local output:" & vbCrLf & strLocalOutputFolder
    End If

    If Len(strSharePointOutputFolder) > 0 Then
        strResult = strResult & vbCrLf & vbCrLf
        strResult = strResult & "SharePoint output:" & vbCrLf & strSharePointOutputFolder
    End If

ExitPoint:
    Set dictTeamCounts = Nothing
    Set dictConfigured = Nothing

    If errNumber = 0 Then BuildSuccessMessage = strResult
    If errNumber <> 0 Then Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "row;team", lngRow, strTeam)
    GoTo ExitPoint
End Function


