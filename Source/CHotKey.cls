VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       CHotKey
' FILENAME:     C:\Source\CHotKey.cls
' AUTHOR:       Phil Fresle
' CREATED:      13-Mar-2001
' COPYRIGHT:    Copyright 2001-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Based on a VBPJ article by Francesco Balena.
'
' MODIFICATION HISTORY:
' 1.0       13-Mar-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Implements IMsgHook

Public Event HotKey(iKeyCode As Integer, iShift As Integer)

Private Declare Function RegisterHotKeyAPI Lib "user32" Alias "RegisterHotKey" _
    (ByVal hwnd As Long, _
     ByVal id As Long, _
     ByVal fsModifiers As Long, _
     ByVal vk As Long) As Long
     
Private Declare Function UnregisterHotKeyAPI Lib "user32" Alias "UnregisterHotKey" _
    (ByVal hwnd As Long, _
     ByVal id As Long) As Long
     
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, _
     Source As Any, _
     ByVal NumBytes As Long)

Private Const MODULE_NAME = "CHotKey"

' The window used to hook hot keys
Private m_hWnd As Long

' The collection of hotkeys; the key in this collection is the
' pair (KeyCode,Shift), the value is the ID
Private m_HotKeys As Collection

'*******************************************************************************
' Class_Initialize (SUB)
'*******************************************************************************
Private Sub Class_Initialize()
    Set m_HotKeys = New Collection
End Sub

'*******************************************************************************
' Class_Terminate (SUB)
'*******************************************************************************
Private Sub Class_Terminate()
    Dim iCount As Integer
    
    On Error Resume Next
    
    If m_HotKeys.Count > 0 Then
        ' Unregister all active hotkeys
        For iCount = 1 To m_HotKeys.Count
            UnregisterHotKeyAPI m_hWnd, m_HotKeys(iCount)
        Next
        UnhookWindow Me
    End If
End Sub

'*******************************************************************************
' hiWord (FUNCTION)

' Return the high word of a long value
'*******************************************************************************
Private Function hiWord(ByVal lValue As Long) As Integer
    CopyMemory hiWord, ByVal VarPtr(lValue) + 2, 2
End Function

'*******************************************************************************
' IMsgHook_AfterMessage (SUB)
'
' This function is called after the message has been processed by the standard
' window procedure
'*******************************************************************************
Private Sub IMsgHook_AfterMessage(hwnd As Long, _
                                  uMsg As Long, _
                                  wParam As Long, _
                                  lParam As Long, _
                                  retValue As Long)
    Dim iValue As Integer
    Dim iValue2 As Integer
    
    On Error Resume Next
    
    Select Case uMsg
        Case WM_HOTKEY
            ' A registered hotkey has been pressed the low word of lParam
            ' holds the modifier keys
            
            iValue = LowWord(lParam)
            
            If iValue And MOD_SHIFT Then
                iValue2 = vbShiftMask
            End If
            
            If iValue And MOD_CONTROL Then
                iValue2 = iValue2 Or vbCtrlMask
            End If
            
            If iValue And MOD_ALT Then
                iValue2 = iValue2 Or vbAltMask
            End If
            
            ' The high word of lParam holds the virtual key code
            RaiseEvent HotKey(hiWord(lParam), iValue2)
    End Select
End Sub

'*******************************************************************************
' IMsgHook_BeforeMessage (SUB)
'
' This method is called before the message is processed by the standard window
' procedure.
'*******************************************************************************
Private Sub IMsgHook_BeforeMessage(hwnd As Long, _
                                   uMsg As Long, _
                                   wParam As Long, _
                                   lParam As Long, _
                                   retValue As Long, _
                                   wndProcAddr As Long)
'
End Sub

'*******************************************************************************
' LowWord (FUNCTION)
'
' Return the low word of a Long value
'*******************************************************************************
Private Function LowWord(ByVal lValue As Long) As Integer
    CopyMemory LowWord, lValue, 2
End Function

'*******************************************************************************
' RegisterHotKey (SUB)
'
' Register a hotkey for this window. IDs to hotkeys are assigned assuming that
' this is the only place in the application that registers hotkeys
'*******************************************************************************
Public Sub RegisterHotKey(ByVal hwnd As Long, iKeyCode As Integer, iShift As Integer)
    Dim iShiftKeys  As Integer
    Dim lID         As Long
    Dim sKey        As String
    Dim iCount      As Integer
    
    On Error Resume Next
    
    ' Remember the window's handle
    m_hWnd = hwnd

    ' Ensure any similar hotkey is unregistered
    UnregisterHotKey iKeyCode, iShift
    
    ' Find the first free lID
    lID = m_HotKeys.Count + 1
    For iCount = 1 To m_HotKeys.Count
        If m_HotKeys(iCount) <> iCount Then
            lID = iCount
            Exit For
        End If
    Next
    
    ' Build the modifiers parameter
    If iShift And vbShiftMask Then
        iShiftKeys = MOD_SHIFT
    End If
    
    If iShift And vbCtrlMask Then
        iShiftKeys = iShiftKeys Or MOD_CONTROL
    End If
    
    If iShift And vbAltMask Then
        iShiftKeys = iShiftKeys Or MOD_ALT
    End If
    
    ' Finally register the hotkey
    If RegisterHotKeyAPI(hwnd, lID, iShiftKeys, iKeyCode) = 0 Then
        Err.Raise vbObjectError + 999, MODULE_NAME, "Unable to register hotkey"
    End If
    
    If m_HotKeys.Count = 0 Then
        ' Do we need to start subclassing the window?
        HookWindowList Me, m_hWnd, WM_HOTKEY
    End If

    ' Add in the proper location in the collection
    ' build a unique key for the combination iKeyCode/shift
    sKey = iKeyCode & "/" & iShift
    If lID > m_HotKeys.Count Then
        m_HotKeys.Add lID, sKey
    Else
        m_HotKeys.Add lID, sKey, lID
    End If
End Sub

'*******************************************************************************
' UnregisterHotKey (SUB)
'
' Unregister a hotkey. Doesn't raise an error if the hotkey wasn't previously
' registered.
'*******************************************************************************
Public Sub UnregisterHotKey(iKeyCode As Integer, iShift As Integer)
    Dim lID     As Long
    Dim sKey    As String
        
    ' Search the hotkey in the hotkey collection
    On Error Resume Next
    
    ' Build a unique key for the combination iKeyCode/iShift
    sKey = iKeyCode & "/" & iShift
    lID = m_HotKeys(sKey)
    If Err = 0 Then
        ' If the hotkey was in the collection, unregisters and removes it
        ' from the collection
        UnregisterHotKeyAPI m_hWnd, lID
        m_HotKeys.Remove sKey
        
        ' Do we need to stop subclassing?
        If m_HotKeys.Count = 0 Then
            UnhookWindow Me
        End If
    End If
End Sub
