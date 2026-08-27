Option Explicit

Public Type TAppConfig
    SaveLocal As Boolean
    SaveSharePoint As Boolean
    UpdateHistorical As Boolean
    OutlookMailbox As String
    OutlookSourceFolder As String
    OutlookArchiveFolder As String
    OutlookSubjectPrefix As String
    OutlookAttachmentPrefix As String
    InputWorksheetName As String
    FondslisteFolder As String
    FondslistePattern As String
    FondslisteExtension As String
    OutputLocalBase As String
    OutputSharePointBase As String
    SupportedTeams As String
    CreateEmailDraft As Boolean
    MailSubject As String
    MailBody As String
    MailBodyNoTeams As String
    MailTo As String
    MailCc As String
    MailFrom As String
    MailSharePointUrl As String
End Type
