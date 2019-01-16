Attribute VB_Name = "MBrowseForFolder"
'*******************************************************************************
' MODULE:       MBrowseForFolder
' FILENAME:     C:\CodeHelperFree\Source\MBrowseForFolder.bas
' AUTHOR:       Phil Fresle
' CREATED:      02-May-2000
' COPYRIGHT:    Copyright 2000-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Use the shell function to browse for a folder
'
' MODIFICATION HISTORY:
' 1.2       02-May-2000
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private Const BIF_RETURNONLYFSDIRS      As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN     As Long = &H2
Private Const BIF_STATUSTEXT            As Long = &H4
Private Const BIF_RETURNFSANCESTORS     As Long = &H8
Private Const BIF_EDITBOX               As Long = &H10
Private Const BIF_VALIDATE              As Long = &H20
Private Const BIF_USENEWUI              As Long = &H40
Private Const BIF_BROWSEFORCOMPUTER     As Long = &H1000
Private Const BIF_BROWSEFORPRINTER      As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES    As Long = &H4000
Private Const MAX_PATH                  As Long = 260
Private Const WM_USER                   As Long = &H400
Private Const BFFM_INITIALIZED          As Long = 1
Private Const BFFM_SELECTIONCHANGED     As Long = 2
Private Const BFFM_ENABLEOK             As Long = (WM_USER + 101)   ' wParam should be 0
Private Const BFFM_SETSELECTION         As Long = (WM_USER + 102)   ' wParam should be 1
Private Const BFFM_SETSTATUSTEXT        As Long = (WM_USER + 100)   ' wParam Should be 0
      
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As String) As Long
    
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long
                                        
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, ByVal lpBuffer As String) As Long
                                        
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    
Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetDC Lib "user32" _
    (ByVal hwnd As Long) As Long
    
Private Declare Function ReleaseDC Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal hdc As Long) As Long

Private Declare Function PathCompactPathW Lib "shlwapi.dll" _
    (ByVal hdc As Long, _
     ByVal lpszPath As Long, _
     ByVal dx As Long) As Boolean

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private m_sStartingPoint As String

'*******************************************************************************
' BrowseCallbackProc (FUNCTION)
'
' PARAMETERS:
' (In) - hwnd   - Long -
' (In) - uMsg   - Long -
' (In) - lParam - Long -
' (In) - lpData - Long -
'
' RETURN VALUE:
' Long -
'
' DESCRIPTION:
' Windows calls this back when the shell browser initialises and when the user
' selects a new folder
'*******************************************************************************
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal lParam As Long, ByVal lpData As Long) As Long

    Dim sBuffer         As String
    Dim bResult         As Boolean
    Dim sCompactedPath  As String
    Dim lDeviceContext  As Long
    Dim oRect           As RECT
    Dim lWidth          As Long
    
    Const BORDERS = 30
    
    On Error Resume Next
    
    Select Case uMsg
        ' Set the starting directory
        Case BFFM_INITIALIZED
            SendMessage hwnd, BFFM_SETSELECTION, 1, m_sStartingPoint

        ' Set the status text with the name of the currently selected directory
        Case BFFM_SELECTIONCHANGED
            sBuffer = Space(MAX_PATH)
            If SHGetPathFromIDList(lParam, sBuffer) <> 0 Then
                ' Get size to fit path into
                GetWindowRect hwnd, oRect
                lWidth = oRect.Right - oRect.Left - BORDERS
                
                ' Compact the path into the size available
                lDeviceContext = GetDC(hwnd)
                sCompactedPath = sBuffer
                bResult = PathCompactPathW(lDeviceContext, StrPtr(sCompactedPath), lWidth)
                ReleaseDC hwnd, lDeviceContext
                If bResult Then
                    sBuffer = sCompactedPath
                End If
          
                SendMessage hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer
            End If
    End Select
End Function

'*******************************************************************************
' LongPointerToFunction (FUNCTION)
'
' PARAMETERS:
' (In/Out) - lAddress - Long -
'
' RETURN VALUE:
' Long -
'
' DESCRIPTION:
' Convert a function address to a long pointer
'*******************************************************************************
Private Function LongPointerToFunction(lAddress As Long) As Long
    LongPointerToFunction = lAddress
End Function

'*******************************************************************************
' BrowseForFolder (FUNCTION)
'
' PARAMETERS:
' (In) - lOwnerWindow   - Long   - Handle to the parent window
' (In) - sTitle         - String - Title to put in folder browser
' (In) - sStartingPoint - String - Folder to use as initial starting point
'
' RETURN VALUE:
' String - Folder selected
'
' DESCRIPTION:
' Browse for a folder using the shell function
'*******************************************************************************
Public Function BrowseForFolder(ByVal lOwnerWindow As Long, _
    ByVal sTitle As String, ByVal sStartingPoint As String) As String

    Dim udtBrowseInfo   As BrowseInfo
    Dim lIdList         As Long
    Dim sBuffer         As String
    
    m_sStartingPoint = sStartingPoint
    
    With udtBrowseInfo
        .hWndOwner = lOwnerWindow
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT
        .lpfnCallback = LongPointerToFunction(AddressOf BrowseCallbackProc)
    End With
        
    lIdList = SHBrowseForFolder(udtBrowseInfo)

    If (lIdList) Then
        sBuffer = Space(MAX_PATH)
        
        SHGetPathFromIDList lIdList, sBuffer
        
        BrowseForFolder = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        BrowseForFolder = ""
    End If
End Function

