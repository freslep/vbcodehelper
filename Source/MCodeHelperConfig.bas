Attribute VB_Name = "MCodeHelperConfig"
'*******************************************************************************
' MODULE:       MCodeHelperConfig
' FILENAME:     C:\My Code\vb\vbch\Source\MCodeHelperConfig.bas
' AUTHOR:       Phil Fresle
' CREATED:      10-May-2000
' COPYRIGHT:    Copyright 2001-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Main module if running standalone
'
' MODIFICATION HISTORY:
' 1.0       20-Jan-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private Const MODULE_NAME As String = "MCodeHelperConfig"

'*******************************************************************************
' Main (SUB)
'*******************************************************************************
Public Sub Main()
    Dim oClient         As CEntryPoint
    Dim sTitle          As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
'    Dim oFSO As FileSystemObject
'    Dim oTS As TextStream
    
    Const PROCEDURE_NAME As String = "Main"
    
    On Error GoTo ERROR_HANDLER
    
    If App.StartMode = vbSModeStandalone Then
        If App.PrevInstance Then
            sTitle = App.Title
            App.Title = "Duplicate " & App.Title
            AppActivate sTitle
            SendKeys "%R", True
            Exit Sub
        End If
        
        Set oClient = New CEntryPoint
        oClient.ShowConfigForm True, g_sLicenseKey, g_sLicensedTo
        Set oClient = Nothing
    End If
Exit Sub
TIDY_UP:
    On Error Resume Next

    If lErrNumber <> 0 Then
        ShowUnexpectedError lErrNumber, sErrDescription, sErrSource
    End If
Exit Sub
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCEDURE_NAME)
    Resume TIDY_UP
End Sub

'Public Sub WriteDebug(sMessage As String)
'    Dim oFSO As FileSystemObject
'    Dim oTS  As TextStream
'    Static lCount As Long
'
'    On Error Resume Next
'
'#If DebugVersion = 1 Then
'    If App.StartMode = vbSModeStandalone Then
'        lCount = lCount + 1
'        Set oFSO = New FileSystemObject
'        Set oTS = oFSO.OpenTextFile(App.Path & "\vbchdebug.txt", ForAppending, False)
'        oTS.WriteLine lCount & " " & sMessage
'        oTS.Close
'        Set oTS = Nothing
'        Set oFSO = Nothing
'    End If
'#End If
'End Sub
