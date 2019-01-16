VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMsgHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       IMsgHook
' FILENAME:     C:\Source\IMsgHook.cls
' AUTHOR:       Phil Fresle
' CREATED:      13-Mar-2001
' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' Interface class
'
' MODIFICATION HISTORY:
' 1.0       13-Mar-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

'*******************************************************************************
' BeforeMessage (SUB)
'
' This method is called before the original window procedure is invoked
'*******************************************************************************
Public Sub BeforeMessage(hwnd As Long, _
                         uMsg As Long, _
                         wParam As Long, _
                         lParam As Long, _
                         retValue As Long, _
                         wndProcAddr As Long)
    '
End Sub

'*******************************************************************************
' AfterMessage (SUB)
'
' This method is called after the original window procedure is invoked
'*******************************************************************************
Public Sub AfterMessage(hwnd As Long, _
                        uMsg As Long, _
                        wParam As Long, _
                        lParam As Long, _
                        retValue As Long)
    '
End Sub
