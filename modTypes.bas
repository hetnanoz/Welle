Option Explicit

Public Type TAppConfig
    SaveLocal As Boolean
    SaveSharePoint As Boolean
    UpdateHistorical As Boolean
    OutlookMailbox As String
    OutlookSourceFolder As String
    OutlookArchiveFolder As String
    FondslisteFolder As String
    FondslistePattern As String
    FondslisteExtension As String
    OutputLocalBase As String
    OutputSharePointBase As String
    SupportedTeams As String
End Type
