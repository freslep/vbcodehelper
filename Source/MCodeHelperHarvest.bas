Attribute VB_Name = "MCodeHelperHarvest"
Option Explicit

Private Const MODULE_NAME As String = "MCodeHelperHarvest"

Public Type udtParameters
    ParameterName   As String
    ParameterType   As String
    InOutBoth       As String
    IsParamArray    As Boolean
    IsOptional      As Boolean
    OptionalValue   As String
End Type

Public Enum enumMode
    PTDesign = 0
    PTRun = 1
    PTCompile = 2
End Enum

Public Enum fslSourceTypes
    fslModuleHeader
    fslProcedureHeader
    fslTimeStamp
    fslUserTemplate
End Enum

Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type COLORMAP
    from As Long
    to As Long
End Type

Public Type PICTDESC
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Declare Function CreateMappedBitmap Lib "comctl32.dll" _
    (ByVal hInstance As Long, _
    ByVal idBitmap As Long, _
    ByVal wFlags As Long, _
    lpColorMap As COLORMAP, _
    ByVal iNumMaps As Long) As Long
    
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (desc As PICTDESC, _
    RefIID As IID, _
    ByVal fPictureOwnsHandle As Long, _
    IPic As IPicture) As Long
    
Public Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long
    
Public Const COLOR_MENU    As Long = 4
Public Const COLOR_BTNFACE As Long = 15

Public Function BitmapToPicture(ByVal hBmp As Long, _
                                Optional ByVal hPal As Long = 0) As IPicture
    Dim IPic        As IPicture
    Dim picdes      As PICTDESC
    Dim iidIPicture As IID
    
    With picdes
        .Size = Len(picdes)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With
    
    With iidIPicture
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    OleCreatePictureIndirect picdes, iidIPicture, True, IPic
    
    Set BitmapToPicture = IPic
End Function

'*******************************************************************************
' LoadButtonPicture (SUB)
'*******************************************************************************
Public Sub LoadButtonPicture(ByVal lResID As Long, _
                             ByVal cmdButton As Office.CommandBarButton)
    Dim lErrorCount As Long
    Dim oPicture    As StdPicture
    Dim oMap        As COLORMAP
    
    Const MAX_ERROR_COUNT As Long = 5
    
    On Error GoTo PICTURE_ERROR
    
    lErrorCount = 0
            
    oMap.from = &H80
    oMap.to = GetSysColor(COLOR_BTNFACE)
    
    Clipboard.SetData BitmapToPicture(CreateMappedBitmap(App.hInstance, lResID, 0, oMap, 1))
    
    cmdButton.PasteFace
    
    Clipboard.Clear
Exit Sub

PICTURE_ERROR:
    If lErrorCount < MAX_ERROR_COUNT Then
        lErrorCount = lErrorCount + 1
        Sleep 10
        Resume
    Else
        Resume TIDY_UP
    End If
Exit Sub

TIDY_UP:
    Err.Clear
End Sub



