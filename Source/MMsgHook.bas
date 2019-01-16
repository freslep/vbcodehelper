Attribute VB_Name = "MMsgHook"
'*******************************************************************************
' MODULE:       MMsgHook
' FILENAME:     C:\My Code\vb\vbch\Source\MMsgHook.bas
' AUTHOR:       Phil Fresle
' CREATED:      13-Mar-2001
' COPYRIGHT:    Copyright 2001-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Based on a VBPJ article by Francesco Balena
' Support module for subclassing. Only windows belonging to current application
' may be subclassed.
'
' MODIFICATION HISTORY:
' 1.0       13-Mar-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private Const MODULE_NAME As String = "MMsgHook"

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
     ByVal hwnd As Long, _
     ByVal Msg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long
     
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
     
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, _
     Source As Any, _
     ByVal NumBytes As Long)

Public Const GWL_WNDPROC As Long = -4

Public Const WM_DESTROY                 As Long = &H2
Public Const WM_ACTIVATEAPP             As Long = &H1C
Public Const WM_CAPTURECHANGED          As Long = &H215
Public Const WM_CHANGECBCHAIN           As Long = &H30D
Public Const WM_CHAR                    As Long = &H102
Public Const WM_CLOSE                   As Long = &H10
Public Const WM_COMMAND                 As Long = &H111
Public Const WM_COMPACTING              As Long = &H41
Public Const WM_CONTEXTMENU             As Long = &H7B
Public Const WM_COPYDATA                As Long = &H4A
Public Const WM_DEVMODECHANGE           As Long = &H1B
Public Const WM_DEVICECHANGE            As Long = &H219
Public Const WM_DISPLAYCHANGE           As Long = &H7E
Public Const WM_DRAWCLIPBOARD           As Long = &H308
Public Const WM_DROPFILES               As Long = &H233
Public Const WM_ENDSESSION              As Long = &H16
Public Const WM_ENTERMENULOOP           As Long = &H211
Public Const WM_ERASEBKGND              As Long = &H14
Public Const WM_EXITMENULOOP            As Long = &H212
Public Const WM_FONTCHANGE              As Long = &H1D
Public Const WM_GETMINMAXINFO           As Long = &H24
Public Const WM_HOTKEY                  As Long = &H312
Public Const WM_HSCROLL                 As Long = &H114
Public Const WM_KEYDOWN                 As Long = &H100
Public Const WM_KEYUP                   As Long = &H101
Public Const WM_KILLFOCUS               As Long = &H8
Public Const WM_LBUTTONDBLCLK           As Long = &H203
Public Const WM_LBUTTONDOWN             As Long = &H201
Public Const WM_LBUTTONUP               As Long = &H202
Public Const WM_MBUTTONDBLCLK           As Long = &H209
Public Const WM_MBUTTONDOWN             As Long = &H207
Public Const WM_MBUTTONUP               As Long = &H208
Public Const WM_MENUCHAR                As Long = &H120
Public Const WM_MENUSELECT              As Long = &H11F
Public Const WM_MOUSEACTIVATE           As Long = &H21
Public Const WM_MOUSEMOVE               As Long = &H200
Public Const WM_MOUSEWHEEL              As Long = &H20A
Public Const WM_MOVE                    As Long = &H3
Public Const WM_MOVING                  As Long = &H216
Public Const WM_NCACTIVATE              As Long = &H86
Public Const WM_NCHITTEST               As Long = &H84
Public Const WM_NCLBUTTONDBLCLK         As Long = &HA3
Public Const WM_NCLBUTTONDOWN           As Long = &HA1
Public Const WM_NCLBUTTONUP             As Long = &HA2
Public Const WM_NCMBUTTONDBLCLK         As Long = &HA9
Public Const WM_NCMBUTTONDOWN           As Long = &HA7
Public Const WM_NCMBUTTONUP             As Long = &HA8
Public Const WM_NCMOUSEMOVE             As Long = &HA0
Public Const WM_NCPAINT                 As Long = &H85
Public Const WM_NCRBUTTONDBLCLK         As Long = &HA6
Public Const WM_NCRBUTTONDOWN           As Long = &HA4
Public Const WM_NCRBUTTONUP             As Long = &HA5
Public Const WM_NOTIFY                  As Long = &H4E
Public Const WM_OTHERWINDOWCREATED      As Long = &H42
Public Const WM_OTHERWINDOWDESTROYED    As Long = &H43
Public Const WM_PAINT                   As Long = &HF
Public Const WM_PALETTECHANGED          As Long = &H311
Public Const WM_PALETTEISCHANGING       As Long = &H310
Public Const WM_POWER                   As Long = &H48
Public Const WM_POWERBROADCAST          As Long = &H218
Public Const WM_QUERYENDSESSION         As Long = &H11
Public Const WM_QUERYNEWPALETTE         As Long = &H30F
Public Const WM_QUERYOPEN               As Long = &H13
Public Const WM_RBUTTONDBLCLK           As Long = &H206
Public Const WM_RBUTTONDOWN             As Long = &H204
Public Const WM_RBUTTONUP               As Long = &H205
Public Const WM_SETCURSOR               As Long = &H20
Public Const WM_SETFOCUS                As Long = &H7
Public Const WM_SETTINGCHANGE           As Long = &H1A
Public Const WM_SIZE                    As Long = &H5
Public Const WM_SIZING                  As Long = &H214
Public Const WM_SPOOLERSTATUS           As Long = &H2A
Public Const WM_SYSCOLORCHANGE          As Long = &H15
Public Const WM_SYSCOMMAND              As Long = &H112
Public Const WM_SYSKEYDOWN              As Long = &H104
Public Const WM_SYSKEYUP                As Long = &H105
Public Const WM_TIMECHANGE              As Long = &H1E
Public Const WM_USERCHANGED             As Long = &H54
Public Const WM_VSCROLL                 As Long = &H115
Public Const WM_WININICHANGE            As Long = &H1A

' used by mouse messages
Public Const MK_CONTROL As Long = &H8
Public Const MK_LBUTTON As Long = &H1
Public Const MK_MBUTTON As Long = &H10
Public Const MK_RBUTTON As Long = &H2
Public Const MK_SHIFT   As Long = &H4

' return value of WM_MOUSEACTIVATE message
Public Const MA_ACTIVATE            As Long = 1
Public Const MA_ACTIVATEANDEAT      As Long = 2
Public Const MA_NOACTIVATE          As Long = 3
Public Const MA_NOACTIVATEANDEAT    As Long = 4

' used by RegisterHotKey
Public Const MOD_ALT        As Long = &H1
Public Const MOD_CONTROL    As Long = &H2
Public Const MOD_SHIFT      As Long = &H4

' Initial number of items in WndInfo
Public Const INIT_WNDINFO_SIZE = 16

' Fill ratio for the hash table
Public Const FILL_FACTOR = 4

Private Type TWndInfo
    ' handle of subclassed window - zero if none
    ' there can be multiple items with same hWnd
    hwnd As Long
    ' address of original window procedure
    wndProcAddr As Long
    ' pointer to object to be notified when a message arrives
    obj_Ptr As Long
    ' hash table of messages to be trapped
    msgTable() As Long
    ' pointers to the previous/next item in this structure that refers to the same window
    prevItem As Integer    ' 0 if first item
    nextItem As Integer    ' 0 if last item
End Type

' this holds info on each subclassed window
Private wndInfo() As TWndInfo

' this hash table is used for quick look to an index
' to WndInfo given a hWnd value
Private hashTable() As Long

'*******************************************************************************
' HookWindow (SUB)
'
' Start the subclassing of a window.
' msgs() includes all the msg numbers to be trapped
' use values >0 to invoke events AFTER regular processing
' use values <0 (negated) to invoke events BEFORE regular processing
' use special value +1/-1 to subclass all messages
'*******************************************************************************
Public Sub HookWindow(obj As IMsgHook, _
                      ByVal hwnd As Long, _
                      msgs() As Long)
    Static arrayCreated     As Boolean
    Dim Index               As Long
'VBCH    Dim prevIndex           As Long
    Dim msgSize             As Integer
    Dim msgIndex            As Integer
    Dim hashIndex           As Long
    Dim i                   As Integer
        
    On Error Resume Next
    
    ' create the array the first time
    If Not arrayCreated Then
        arrayCreated = True
        ReDim wndInfo(INIT_WNDINFO_SIZE) As TWndInfo
        ReDim hashTable(1 To INIT_WNDINFO_SIZE * FILL_FACTOR + 1) As Long
    End If
    
    ' Exit if we requested no messages
    If LBound(msgs) > UBound(msgs) Then
        Exit Sub
    End If
    
    ' search the first available slot
    Index = 1
    Do Until wndInfo(Index).hwnd = 0
        Index = Index + 1
        If Index > UBound(wndInfo) Then
            ' expand the WndInfo array if all slots are engaged
            ReDim Preserve wndInfo(Index + INIT_WNDINFO_SIZE) As TWndInfo
        End If
    Loop
        
    ' save data in local structure
    wndInfo(Index).hwnd = hwnd
    ' save pointer to object, not a reference
    ' this will not keep the object alive when it goes
    ' out of scope in the main program
    wndInfo(Index).obj_Ptr = ObjPtr(obj)
    ' clear pointers
    wndInfo(Index).prevItem = 0
    wndInfo(Index).nextItem = 0
    
    ' build a hash table of messages to trap
    ' (odd numbers minimize collisions)
    msgSize = (UBound(msgs) - LBound(msgs) + 1)
    If msgSize = 1 And Abs(msgs(LBound(msgs)) = 1) Then
        ' requested to subclass all messages
        ReDim wndInfo(Index).msgTable(1 To 1) As Long
        wndInfo(Index).msgTable(1) = msgs(LBound(msgs))
    Else
        ' subclass only a number of messages
        ReDim wndInfo(Index).msgTable(1 To msgSize * 4 + 1) As Long
        For i = LBound(msgs) To UBound(msgs)
            ' do store the same message twice
            msgIndex = Abs(MsgHashSearch(wndInfo(Index).msgTable(), msgs(i)))
            wndInfo(Index).msgTable(msgIndex) = msgs(i)
        Next
    End If
    ' check if this window is already subclassed
    hashIndex = HashSearch(hwnd)
    If hashIndex > 0 Then
        i = hashTable(hashIndex)
        hashTable(hashIndex) = Index
        ' put the new item in front of the linked list
        ' (wndInfo(index).prevItem is zero, which means "begin of list")
        wndInfo(i).prevItem = Index
        wndInfo(Index).nextItem = i
        wndInfo(Index).wndProcAddr = wndInfo(i).wndProcAddr
        Exit Sub
    End If
    
    ' the window isn't currently sublassed
    ' (-hashIndex) contains the correct index in the hash table
    hashTable(-hashIndex) = Index
    
    ' if the hash table is getting too crowded
    ' lets build a larger one
    If UBound(wndInfo) * FILL_FACTOR > UBound(hashTable) Then
        ' the hash table should be at least tice as large as WndInfo
        RehashTable UBound(wndInfo) * FILL_FACTOR + 1
    End If
    
    ' START SUBCLASSING
    ' enforce new window procedure, save old address
    wndInfo(Index).wndProcAddr = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

'*******************************************************************************
' HookWindowList (SUB)
'
' Similar to HookWindow, but messages can be specified on the command line
' use -1/+1 to trap ALL messages (before or after default window procedure)
'*******************************************************************************
Public Sub HookWindowList(obj As IMsgHook, _
                          ByVal hwnd As Long, _
                          ParamArray msgList() As Variant)
    Dim Index           As Integer
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "HookWindowList"
    
    On Error GoTo ERROR_HANDLER
    
    ' exit if no events were specified
    If UBound(msgList) < 0 Then
        Exit Sub
    End If
    
    ' build any array holding msg numbers
    ReDim msgs(LBound(msgList) To UBound(msgList)) As Long
    For Index = LBound(msgList) To UBound(msgList)
        msgs(Index) = msgList(Index)
    Next
    
    ' let HookWindow do the real job
    HookWindow obj, hwnd, msgs()
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
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCEDURE_NAME)
    Resume TIDY_UP
End Sub

'*******************************************************************************
' UnhookWindow (SUB)
'
' Stop subclassing the window associated to a given object
'*******************************************************************************
Public Sub UnhookWindow(obj As IMsgHook)
    Dim Index           As Long
    Dim obj_Ptr         As Long
    Dim hashIndex       As Integer
    Dim emptyWndInfo    As TWndInfo
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "UnhookWindow"
    
    On Error GoTo ERROR_HANDLER
    
    ' search this object in the wndInfo() array
    obj_Ptr = ObjPtr(obj)
    For Index = 1 To UBound(wndInfo)
        If wndInfo(Index).obj_Ptr = obj_Ptr Then
            Exit For
        End If
    Next
    
    ' exit if this object is not in wndInfo
    If Index > UBound(wndInfo) Then
        Exit Sub
    End If
    
    ' search the corresponding hWnd in the hash table
    ' this is the start of the linked list
    hashIndex = HashSearch(wndInfo(Index).hwnd)
    
    ' if this item belongs to a linked list, just compact the list
    If wndInfo(Index).prevItem Then
        ' it is not at the beginning of the list
        ' adject the NEXT pointer of the previous item
        wndInfo(wndInfo(Index).prevItem).nextItem = wndInfo(Index).nextItem
        If wndInfo(Index).nextItem Then
            ' if it is not the last item of the list
            ' adjust the PREV pointer of the next item
            wndInfo(wndInfo(Index).nextItem).prevItem = wndInfo(Index).prevItem
        End If
        ' reset all fields in the slot
        wndInfo(Index) = emptyWndInfo
    ElseIf wndInfo(Index).nextItem Then
        ' it is at the beginning of the list of at least two items
        hashTable(hashIndex) = wndInfo(Index).nextItem
        wndInfo(Index).prevItem = 0
        ' reset all fields in the slot
        wndInfo(Index) = emptyWndInfo
    Else
        ' it is the only object that subclass that window
        ' therefore we must stop subclassing
        SetWindowLong wndInfo(Index).hwnd, GWL_WNDPROC, wndInfo(Index).wndProcAddr
        ' reset all fields in the slot, then rebuild the hash table
        wndInfo(Index) = emptyWndInfo
        RehashTable UBound(hashTable)
    End If
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
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCEDURE_NAME)
    Resume TIDY_UP
End Sub

'*******************************************************************************
' WndProc (FUNCTION)
'*******************************************************************************
Public Function WndProc(ByVal hwnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
    Dim Index           As Long
    Dim hashIndex       As Long
    Dim msgIndex        As Integer
    Dim msgValue        As Long
    Dim wndProcAddr     As Long
    Dim retValue        As Long
    Dim nextIndex       As Long
    Dim emptyWndInfo    As TWndInfo
    Dim obj             As IMsgHook
    Dim notifyCount     As Integer
    
    On Error Resume Next
        
    hashIndex = HashSearch(hwnd)
    Index = hashTable(hashIndex)
    ' this *should* never happen - however, if it does happen
    ' we can't do much, because we lost the original procAddr of the window
    Debug.Assert Index > 0
    
    ' get the address of the original window procedure
    ' if this value is still non-null on exit, it means
    ' that no object called the original window proc manually
    wndProcAddr = wndInfo(Index).wndProcAddr
    
    ' this is the list of items that must be notified
    ' after the original window procedure has been invoked
    ReDim notifyList(1 To UBound(wndInfo)) As Integer
    
    Do
        If UBound(wndInfo(Index).msgTable) = 1 And Abs(wndInfo(Index).msgTable(1)) = 1 Then
            ' we requested to subclass all messages
            msgIndex = 1
        Else
            ' try to find the message in the msgTable()
            msgIndex = MsgHashSearch(wndInfo(Index).msgTable(), uMsg)
        End If
        
        If msgIndex > 0 Then
            ' get the value of the message
            msgValue = wndInfo(Index).msgTable(msgIndex)
            If msgValue > 0 Then
                ' message post-processing
                ' simply add this index to the notify list
                notifyCount = notifyCount + 1
                notifyList(notifyCount) = Index
            Else
                ' message pre-processing
                ' build a temporary reference to the IMsgHook object
                ' and notify it the event
                CopyMemory obj, wndInfo(Index).obj_Ptr, 4
                obj.BeforeMessage hwnd, uMsg, wParam, lParam, retValue, wndProcAddr
                If wndProcAddr Then
                    ' if the wndProcAddr parameter is still non-null
                    ' process the message in the standard window proc
                    retValue = CallWindowProc(wndProcAddr, hwnd, uMsg, wParam, lParam)
                    wndProcAddr = 0
                End If
                ' if the message value was -1, also add this index to the list
                If msgValue = -1 Then
                    notifyCount = notifyCount + 1
                    notifyList(notifyCount) = Index
                End If
            End If
        End If
        
        ' see if there is another object that refers to the same window
        Index = wndInfo(Index).nextItem
    Loop While Index
    
    ' if no one ever called the original window procedure
    If wndProcAddr Then
        retValue = CallWindowProc(wndProcAddr, hwnd, uMsg, wParam, lParam)
    End If
    
    ' send an AfterMessage notification to all object in the notify list
    For Index = 1 To notifyCount
        ' build a temporary reference to the object
        CopyMemory obj, wndInfo(notifyList(Index)).obj_Ptr, 4
        ' then notify the event to the IMsgHook object
        obj.AfterMessage hwnd, uMsg, wParam, lParam, retValue
    Next
    
    ' destroy the temporary object
    ' (if you omit this step, VB will crash )
    CopyMemory obj, 0&, 4
    
    ' if this was a destroy message, stop subclassing
    If uMsg = WM_DESTROY Then
        Index = hashTable(hashIndex)
        ' restore old procedure address
        SetWindowLong wndInfo(Index).hwnd, GWL_WNDPROC, wndInfo(Index).wndProcAddr
        ' clear items in the table
        Do While Index
            nextIndex = wndInfo(Index).nextItem
            wndInfo(Index) = emptyWndInfo
            Index = nextIndex
        Loop
    End If
    
    ' return the last return value
    WndProc = retValue
End Function

'*******************************************************************************
' CallWindowProcedure (FUNCTION)
'
' Call a window procedure
' This is suitable for calling from within a BeforeMessage method
' because automatically clears the wndProcAddr argument
'*******************************************************************************
Public Function CallWindowProcedure(wndProcAddr As Long, _
                                    ByVal hwnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long
    If wndProcAddr Then
        CallWindowProcedure = CallWindowProc(wndProcAddr, hwnd, uMsg, wParam, lParam)
        wndProcAddr = 0
    End If
End Function

'*******************************************************************************
' RehashTable (SUB)
'
' Re-hash hash table (private)
'*******************************************************************************
Private Sub RehashTable(ByVal hashSize As Long)
    ' rebuild the hash table using the data in WndInfo
    Dim Index       As Long
    Dim hashIndex   As Long
    
    ReDim hashTable(1 To hashSize) As Long
    
    For Index = 1 To UBound(wndInfo)
        ' do not take empty slot into account, nor items that
        ' are not the beginning of a linked list
        If wndInfo(Index).hwnd And wndInfo(Index).prevItem = 0 Then
            hashIndex = HashSearch(wndInfo(Index).hwnd)
            ' the result should be a negative number
            hashTable(-hashIndex) = Index
        End If
    Next
End Sub

'*******************************************************************************
' HashSearch (FUNCTION)
'
' search an item in the main hash table (hWnd)
' if found, return its index in the table
' if not found, return the negated index of the first available slot
'*******************************************************************************
Private Function HashSearch(ByVal hwnd As Long) As Integer
    Dim Index       As Long
    Dim hashSize    As Integer
    
    hashSize = UBound(hashTable)
    Index = (hwnd Mod hashSize) + 1
    Do
        If hashTable(Index) = 0 Then
            HashSearch = -Index
            Exit Function
        ElseIf wndInfo(hashTable(Index)).hwnd = hwnd Then
            HashSearch = Index
            Exit Function
        End If
        Index = Index + 1
        If Index > hashSize Then Index = 1
    Loop
End Function

'*******************************************************************************
' MsgHashSearch (FUNCTION)
'
' search an item in a message hash table (search for negated items too)
' if found, return its index in the hash table
' if not found, return the negated index of the first available slot
'*******************************************************************************
Private Function MsgHashSearch(hashTable() As Long, ByVal value As Long) As Integer
    Dim Index       As Long
    Dim hashSize    As Integer
    
    hashSize = UBound(hashTable)
    Index = (value Mod hashSize) + 1
    Do
        If hashTable(Index) = value Or hashTable(Index) = -value Then
            MsgHashSearch = Index
            Exit Function
        ElseIf hashTable(Index) = 0 Then
            MsgHashSearch = -Index
            Exit Function
        End If
        Index = Index + 1
        If Index > hashSize Then
            Index = 1
        End If
    Loop
End Function
