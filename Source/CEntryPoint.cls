VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEntryPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' MODULE:       CEntryPoint
' FILENAME:     C:\My Code\vb\vbch\Source\CEntryPoint.cls
' AUTHOR:       Phil Fresle
' CREATED:      10-May-2000
' COPYRIGHT:    Copyright 2001-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Entry point class
'
' MODIFICATION HISTORY:
' 1.0       17-Jan-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private Const MODULE_NAME As String = "CEntryPoint"

Private WithEvents m_frmConfigure As FConfigure
Attribute m_frmConfigure.VB_VarHelpID = -1

Public Event ConfigFinished(ByVal bUpdateOptions As Boolean)

'*******************************************************************************
' Class_Terminate (SUB)
'*******************************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    
    Unload m_frmConfigure
    Set m_frmConfigure = Nothing
End Sub

'*******************************************************************************
' m_frmConfigure_Unloading (SUB)
'*******************************************************************************
Private Sub m_frmConfigure_Unloading(ByVal bOK As Boolean)
    RaiseEvent ConfigFinished(bOK)
End Sub

'*******************************************************************************
' RegisterApplication (FUNCTION)
'*******************************************************************************
Public Function RegisterApplication() As Boolean
End Function

'*******************************************************************************
' ShowConfigForm (SUB)
'*******************************************************************************
Public Sub ShowConfigForm(ByVal bModal As Boolean, _
                          ByVal sLicenseKey As String, _
                          ByVal sLicensedTo As String)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "ShowConfigForm"
    
    On Error GoTo ERROR_HANDLER
    
    If m_frmConfigure Is Nothing Then
        Set m_frmConfigure = New FConfigure
    End If
    
    m_frmConfigure.Configure sLicenseKey, sLicensedTo
    
    If bModal Then
        m_frmConfigure.Show vbModal
    Else
        On Error Resume Next
        m_frmConfigure.Show
        If Err.Number = 401 Then
            Err.Clear
            On Error GoTo ERROR_HANDLER
            m_frmConfigure.Show vbModal
        Else
            SetWindowPos m_frmConfigure.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        End If
    End If
    
    SetForegroundWindow m_frmConfigure.hwnd
TIDY_UP:
    On Error Resume Next

    If lErrNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Sub
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, "ShowConfigForm")
    Resume TIDY_UP
End Sub
