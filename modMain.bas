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
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngPreviousAutomationSecurity As Long
    Dim strLatestFondslistePath As String
    Dim strWorkspace As String
    Dim varPreviousStatusBar As Variant
    Dim xlPreviousCalculation As XlCalculation
    Dim dtRunTimestamp As Date

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousScreenUpdating = Application.ScreenUpdating
    blnPreviousEnableEvents = Application.EnableEvents
    blnPreviousDisplayAlerts = Application.DisplayAlerts
    xlPreviousCalculation = Application.Calculation
    varPreviousStatusBar = Application.StatusBar
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
    strWorkspace = CreateRunWorkspace(dtRunTimestamp)

    Application.StatusBar = "Dallas Cash Transactions: downloading Outlook attachments..."
    Set colAttachmentPaths = DownloadCashTransactionAttachments(appConfig, strWorkspace, dictAttachmentSubjects)

    Application.StatusBar = "Dallas Cash Transactions: merging input files..."
    arrCombined = MergeInputFiles(colAttachmentPaths, appConfig.InputWorksheetName, dictAttachmentSubjects)

    Application.StatusBar = "Dallas Cash Transactions: loading the latest Fondsliste..."
    strLatestFondslistePath = GetLatestFondslistePath(appConfig)
    Set dictFundTeams = BuildOpenFundTeamDictionary(strLatestFondslistePath)

    Application.StatusBar = "Dallas Cash Transactions: updating Static_Data..."
    Call UpdateStaticDataFromFondsliste(dictFundTeams)
    Set dictAccounts = BuildAccountMappingDictionary()
    Call EnrichTransactions(arrCombined, dictAccounts)

    Set colTeams = ParseSupportedTeams(appConfig.SupportedTeams)
    Set colDestinationRoots = BuildDestinationRoots(appConfig)

    Application.StatusBar = "Dallas Cash Transactions: creating daily reports..."
    Call CreateDailyReports(arrCombined, colTeams, colDestinationRoots, dtRunTimestamp, strWorkspace)

    If appConfig.UpdateHistorical Then
        Application.StatusBar = "Dallas Cash Transactions: updating historical master files..."
        Call UpdateHistoricalFiles(arrCombined, colTeams, colDestinationRoots)
    End If

ExitPoint:
    Set dictAccounts = Nothing
    Set dictAttachmentSubjects = Nothing
    Set dictFundTeams = Nothing
    Set colDestinationRoots = Nothing
    Set colTeams = Nothing
    Set colAttachmentPaths = Nothing

    If errNumber = 0 Then
        If Len(strWorkspace) > 0 Then Call DeleteFolderIfExists(strWorkspace)
    End If

    If blnStateCaptured Then
        Application.StatusBar = varPreviousStatusBar
        Application.AutomationSecurity = lngPreviousAutomationSecurity
        Application.Calculation = xlPreviousCalculation
        Application.DisplayAlerts = blnPreviousDisplayAlerts
        Application.EnableEvents = blnPreviousEnableEvents
        Application.ScreenUpdating = blnPreviousScreenUpdating
    End If

    If errNumber <> 0 Then Call ErrorManager.display
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "workspace;fondslistePath", strWorkspace, strLatestFondslistePath)
    GoTo ExitPoint
End Sub
