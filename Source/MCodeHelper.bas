Attribute VB_Name = "MCodeHelper"
'*******************************************************************************
' MODULE:       MCodeHelper
' FILENAME:     C:\CodeHelperFree\Source\MCodeHelper.bas
' AUTHOR:       Phil Fresle
' CREATED:      01-Dec-1999
' COPYRIGHT:    Copyright 1999-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Generic routines
'
' MODIFICATION HISTORY:
' 1.0       01-Dec-1999
'           Phil Fresle
'           Initial Version
' 1.1       10-Mar-2000
'           Phil Fresle
'           Free Version
' 2.0       21-Jan-2001
'           Phil Fresle
'           Commercial Version
' 6.0       16-Jan-2019
'           Phil Fresle
'           Open source version.
'*******************************************************************************
Option Explicit

Public g_sLicenseType   As String
Public g_lDaysLeft      As Long
Public g_sLicensedTo    As String
Public g_sLicenseKey    As String

Private Const MODULE_NAME  As String = "MCodeHelper"

Public Const HELP_FILE     As String = "vbcodehelper.hlp"

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd As Long, ByVal lpHelpFile As String, _
    ByVal wCommand As Long, ByVal dwData As Any) As Long

Public Const HELP_QUIT          As Long = &H2
Public Const HELP_FINDER        As Long = &HB&

Public Const TEMP_PERIOD                As Long = 30

Public Const PARAM_FORMAT_0             As String = "(in/out) - varname - vartype - "
Public Const PARAM_FORMAT_1             As String = "(in/out) | varname - vartype - "
' **** START CHANGE - PMF - 08-Nov-2005 07:17 ****
' New options
Public Const PARAM_FORMAT_2             As String = "ByVal/ByRef - varname - vartype - "
Public Const PARAM_FORMAT_3             As String = "ByVal/ByRef | varname - vartype - "
' **** END CHANGE - PMF - 07:17 ****

Public Const CAPTION_MODULE             As String = "Document Entire Module"
Public Const CAPTION_PROCEDURE          As String = "Document Procedure"
Public Const CAPTION_TIMESTAMP          As String = "Insert Time Stamp"
Public Const CAPTION_TEMPLATE           As String = "Templates"
Public Const CAPTION_CLOSE              As String = "Close Windows"
Public Const CAPTION_CLEAR              As String = "Clear Immediate Window"
Public Const CAPTION_ERROR1             As String = "Error Handler 1"
Public Const CAPTION_ERROR2             As String = "Error Handler 2"
Public Const CAPTION_DIM                As String = "Clean Variables in Procedure"
Public Const CAPTION_DIM_ALL            As String = "Clean Variables in All Procedures"
Public Const CAPTION_ERROR_ALL          As String = "Error Handle All Procedures"
Public Const CAPTION_INDENT             As String = "Smart Indent Procedure"
Public Const CAPTION_INDENT_ALL         As String = "Smart Indent All Procedures"
Public Const CAPTION_TAB                As String = "Smart Tab Order"
Public Const CAPTION_ZORDER             As String = "ZOrder Management"
Public Const CAPTION_STATS              As String = "Project Statistics"
Public Const CAPTION_WHITE_SPACE        As String = "Tidy White Space in Module"
Public Const CAPTION_WHITE_SPACE_ALL    As String = "Tidy White Space in All Modules"
Public Const CAPTION_PROCEDURE_LIST     As String = "Procedure List"

Public Const SHORTCUT_MODULE            As String = "Document Entire Module (Ctrl+Shift+M)"
Public Const SHORTCUT_PROCEDURE         As String = "Document Procedure (Ctrl+Shift+P)"
Public Const SHORTCUT_TIMESTAMP         As String = "Insert Time Stamp (Ctrl+Shift+S)"
Public Const SHORTCUT_TEMPLATE          As String = "Templates (Ctrl+Shift+T)"
Public Const SHORTCUT_CLOSE             As String = "Close Windows (Ctrl+Shift+F4)"
Public Const SHORTCUT_CLEAR             As String = "Clear Immediate Window (Ctrl+Shift+Bksp)"
Public Const SHORTCUT_ERROR1            As String = "Error Handler 1 (Ctrl+Alt+F1)"
Public Const SHORTCUT_ERROR2            As String = "Error Handler 2 (Ctrl+Alt+F2)"
Public Const SHORTCUT_DIM               As String = "Clean Variables in Procedure(Ctrl+Shift+V)"
Public Const SHORTCUT_DIM_ALL           As String = "Clean Variables in All Procedures(Ctrl+Alt+V)"
Public Const SHORTCUT_ERROR_ALL         As String = "Error Handle All Procedures (Ctrl+Alt+H)"
Public Const SHORTCUT_INDENT            As String = "Smart Indent Procedure (Ctrl+Shift+I)"
Public Const SHORTCUT_INDENT_ALL        As String = "Smart Indent All Procedures (Ctrl+Alt+I)"
Public Const SHORTCUT_TAB               As String = "Smart Tab Order (Ctrl+Alt+T)"
Public Const SHORTCUT_ZORDER            As String = "ZOrder Management (Ctrl+Shift+Z)"
Public Const SHORTCUT_STATS             As String = "Project Statistics (Ctrl+Alt+S)"
Public Const SHORTCUT_WHITE_SPACE       As String = "Tidy White Space in Module (Ctrl+Shift+W)"
Public Const SHORTCUT_WHITE_SPACE_ALL   As String = "Tidy White Space in All Modules (Ctrl+Alt+W)"
Public Const SHORTCUT_PROCEDURE_LIST    As String = "Procedure List (Ctrl+Alt+L)"

Public Const CAPTION_MENU_MODULE        As String = "Document Entire &Module..."
Public Const CAPTION_MENU_PROCEDURE     As String = "Document &Procedure"
Public Const CAPTION_MENU_TIMESTAMP     As String = "Insert Time &Stamp"
Public Const CAPTION_MENU_TEMPLATE      As String = "Insert &Template..."
Public Const CAPTION_MENU_CLOSE         As String = "Close &Windows"
Public Const CAPTION_MENU_CLEAR         As String = "Clear &Immediate Window"
Public Const CAPTION_MENU_ERROR1        As String = "Error Handler &1"
Public Const CAPTION_MENU_ERROR2        As String = "Error Handler &2"
Public Const CAPTION_MENU_DIM           As String = "Clean &Variables in Procedure"
Public Const CAPTION_MENU_DIM_ALL       As String = "Clean Variables in &All Procedures"
Public Const CAPTION_MENU_ERROR_ALL     As String = "Error &Handle All Procedures"
Public Const CAPTION_MENU_INDENT        As String = "Smart I&ndent Procedure"
Public Const CAPTION_MENU_INDENT_ALL    As String = "Smart In&dent All Procedures"
Public Const CAPTION_MENU_TAB           As String = "Smart Tab &Order"
Public Const CAPTION_MENU_ZORDER        As String = "&ZOrder Management"
Public Const CAPTION_MENU_STATS         As String = "Pro&ject Statistics"
Public Const CAPTION_MENU_WHITE_SPACE       As String = "Tidy White Space in Module"
Public Const CAPTION_MENU_WHITE_SPACE_ALL   As String = "Tidy White Space in All Modules"
Public Const CAPTION_MENU_PROCEDURE_LIST    As String = "Procedure &List"

Public Const CAPTION_HELP               As String = "VBCodeHelper &Help..."
Public Const CAPTION_CONFIGURE          As String = "&Configure VBCodeHelper..."
Public Const CAPTION_ABOUT              As String = "&About VBCodeHelper..."

Public Const DEFAULT_COMBO_WIDTH        As Long = 145

Public Const DATA_USER                  As String = "User"
Public Const DATA_LK                    As String = "LKey"
Public Const DATA_KEY                   As String = "DataKey"
Public Const DATA_KEY_WIDE              As String = "DataKeyW"
Public Const COMPANY_NAME               As String = "Frez Systems Limited"

Public Const REG_APP_NAME               As String = "VBCodeHelper"
Public Const REG_SETTINGS               As String = "Settings"

Public Const REG_INDENT                 As String = "Indent"
Public Const REG_NORMAL_ERRORS          As String = "NormalErrors"
Public Const REG_EVENT_ERRORS           As String = "EventErrors"
Public Const REG_PROPERTY_ERRORS        As String = "PropertyErrors"
Public Const REG_INDENT_DIM             As String = "IndentDim"
Public Const REG_BOILERPLATES           As String = "BoilerplatePath"
Public Const REG_TEMPLATES              As String = "TemplatePath"
Public Const REG_USER_TOKENS            As String = "UserTokensPath"
Public Const REG_AUTHOR                 As String = "Author"
Public Const REG_COMPANY                As String = "Company"
Public Const REG_INITIALS               As String = "Initials"
Public Const REG_TIMEFORMAT             As String = "TimeFormat"
Public Const REG_DATEFORMAT             As String = "DateFormat"
Public Const REG_PARAMFORMAT            As String = "ParamFormat"
Public Const REG_DOCBEFORE              As String = "DocBefore"
Public Const REG_CLOSEACTIVE            As String = "CloseActive"
Public Const REG_TOOLBARPOS             As String = "ToolbarPosition"
Public Const REG_TOOLBARINDEX           As String = "ToolbarIndex"
Public Const REG_TOOLBARLEFT            As String = "ToolbarLeft"
Public Const REG_TOOLBARTOP             As String = "ToolbarTop"
Public Const REG_TOOLBARSHOW            As String = "ToolbarShow"
Public Const REG_TAB_DISPLAY            As String = "TabDisplay"
Public Const REG_ZORDER_DISPLAY         As String = "ZOrderDisplay"
Public Const REG_BLANK_LINES            As String = "BlankLines"
Public Const REG_BUTTON0_SHOW           As String = "Button0Show"
Public Const REG_BUTTON1_SHOW           As String = "Button1Show"
Public Const REG_BUTTON2_SHOW           As String = "Button2Show"
Public Const REG_BUTTON3_SHOW           As String = "Button3Show"
Public Const REG_BUTTON4_SHOW           As String = "Button4Show"
Public Const REG_BUTTON5_SHOW           As String = "Button5Show"
Public Const REG_BUTTON6_SHOW           As String = "Button6Show"
Public Const REG_BUTTON7_SHOW           As String = "Button7Show"
Public Const REG_BUTTON8_SHOW           As String = "Button8Show"
Public Const REG_BUTTON9_SHOW           As String = "Button9Show"
Public Const REG_BUTTON10_SHOW          As String = "Button10Show"
Public Const REG_BUTTON11_SHOW          As String = "Button11Show"
Public Const REG_BUTTON12_SHOW          As String = "Button12Show"
Public Const REG_BUTTON13_SHOW          As String = "Button13Show"
Public Const REG_BUTTON14_SHOW          As String = "Button14Show"
Public Const REG_BUTTON15_SHOW          As String = "Button15Show"
Public Const REG_BUTTON16_SHOW          As String = "Button16Show"
Public Const REG_BUTTON17_SHOW          As String = "Button17Show"
Public Const REG_BUTTON18_SHOW          As String = "Button18Show"
Public Const REG_SHORTCUTS              As String = "Shortcuts"
Public Const REG_BUTTON0_SHORTCUT       As String = "Button0Shortcut"
Public Const REG_BUTTON1_SHORTCUT       As String = "Button1Shortcut"
Public Const REG_BUTTON2_SHORTCUT       As String = "Button2Shortcut"
Public Const REG_BUTTON3_SHORTCUT       As String = "Button3Shortcut"
Public Const REG_BUTTON4_SHORTCUT       As String = "Button4Shortcut"
Public Const REG_BUTTON5_SHORTCUT       As String = "Button5Shortcut"
Public Const REG_BUTTON6_SHORTCUT       As String = "Button6Shortcut"
Public Const REG_BUTTON7_SHORTCUT       As String = "Button7Shortcut"
Public Const REG_BUTTON8_SHORTCUT       As String = "Button8Shortcut"
Public Const REG_BUTTON9_SHORTCUT       As String = "Button9Shortcut"
Public Const REG_BUTTON10_SHORTCUT      As String = "Button10Shortcut"
Public Const REG_BUTTON11_SHORTCUT      As String = "Button11Shortcut"
Public Const REG_BUTTON12_SHORTCUT      As String = "Button12Shortcut"
Public Const REG_BUTTON13_SHORTCUT      As String = "Button13Shortcut"
Public Const REG_BUTTON14_SHORTCUT      As String = "Button14Shortcut"
Public Const REG_BUTTON15_SHORTCUT      As String = "Button15Shortcut"
Public Const REG_BUTTON16_SHORTCUT      As String = "Button16Shortcut"
Public Const REG_BUTTON17_SHORTCUT      As String = "Button17Shortcut"
Public Const REG_BUTTON18_SHORTCUT      As String = "Button18Shortcut"
Public Const REG_COMBO_WIDTH            As String = "ComboWidth"

Public Const DEF_REG_TIMEFORMAT         As String = "hh:nn"
Public Const DEF_REG_DATEFORMAT         As String = "dd-mmm-yyyy"
Public Const DEF_REG_PARAMFORMAT        As String = "0"
Public Const DEF_REG_DOCBEFORE          As String = "1"
Public Const DEF_REG_CLOSEACTIVE        As String = "1"
Public Const DEF_REG_TOOLBARINDEX       As String = "0"
Public Const DEF_REG_TOOLBARLEFT        As String = "0"
Public Const DEF_REG_TOOLBARTOP         As String = "0"
Public Const DEF_REG_NORMAL_ERRORS      As String = "1"
Public Const DEF_REG_EVENT_ERRORS       As String = "2"
Public Const DEF_REG_PROPERTY_ERRORS    As String = "1"
Public Const DEF_REG_INDENT             As String = "4"
Public Const DEF_REG_TAB_DISPLAY        As String = "0"
Public Const DEF_REG_ZORDER_DISPLAY     As String = "0"
Public Const DEF_REG_BLANK_LINES        As String = "1"
Public Const DEF_REG_INDENT_DIM         As String = "1"

Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const SYNCHRONIZE               As Long = &H100000
Private Const READ_CONTROL              As Long = &H20000
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const STANDARD_RIGHTS_READ      As Long = (READ_CONTROL)
Private Const KEY_READ                  As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE _
                                                    Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
                                                    And (Not SYNCHRONIZE))
    
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
    (ByVal hKey As Long, _
        ByVal lpClass As String, _
        lpcbClass As Long, _
        ByVal lpReserved As Long, _
        lpcSubKeys As Long, _
        lpcbMaxSubKeyLen As Long, _
        lpcbMaxClassLen As Long, _
        lpcValues As Long, _
        lpcbMaxValueNameLen As Long, _
        lpcbMaxValueLen As Long, _
        lpcbSecurityDescriptor As Long, _
        lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        ByVal lpData As String, _
        lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal cx As Long, _
     ByVal cy As Long, _
     ByVal wFlags As Long) As Long
     
Public Const SWP_NOMOVE        As Long = &H2
Public Const SWP_NOSIZE        As Long = &H1
Public Const FLAGS             As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST      As Long = -1
Public Const HWND_NOTOPMOST    As Long = -2

Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

'Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
'    (ByVal lpszPath As String, _
'     ByVal lpPrefixString As String, _
'     ByVal wUnique As Long, _
'     ByVal lpTempFileName As String) As Long
'
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
'    (ByVal nBufferLength As Long, _
'     ByVal lpBuffer As String) As Long

'*******************************************************************************
' GetInitials (FUNCTION)
'*******************************************************************************
Public Function GetInitials(ByVal sName As String) As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lPos            As Long
    Dim sInitials       As String
    
    Const PROCEDURE_NAME As String = "GetInitials"
    
    On Error GoTo ERROR_HANDLER
    
    sName = Trim(sName)
    
    If sName <> "" Then
        sInitials = Left(sName, 1)
        lPos = 1
        Do
            lPos = InStr(lPos, sName, " ")
            If lPos > 0 Then
                lPos = lPos + 1
                sInitials = sInitials & Mid(sName, lPos, 1)
            End If
        Loop Until lPos = 0
    End If
    
    GetInitials = UCase(sInitials)
TIDY_UP:
    On Error Resume Next

    If lErrNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Function
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCEDURE_NAME)
    Resume TIDY_UP
End Function

'*******************************************************************************
' DebugOutput (SUB)
'*******************************************************************************
Public Sub DebugOutput(sOutput As String)
    Dim lFile           As Long
    Dim sErrDescription As String
    Dim sErrSource      As String
    Dim lErrNumber      As Long
    Dim sMessage        As String
    
    On Error GoTo ERROR_HANDLER
    
    lFile = FreeFile
    
    Open "C:\VBCHDEBUG.TXT" For Append As #lFile
    Print #lFile, sOutput
    Close #lFile
Exit Sub
TIDY_UP:
    Err.Clear
    If lErrNumber <> 0 Then
        sMessage = "Error logging debug message to file." & vbCrLf & _
            lErrNumber & " - " & sErrDescription & vbCrLf & _
            "Debug message = " & sOutput
        MsgBox sMessage, vbExclamation, App.ProductName
    End If
Exit Sub
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, "MakeToken")
    GoTo TIDY_UP
End Sub

'*******************************************************************************
' DebugOutput (SUB)
'*******************************************************************************
Public Function GetFilenameFromPath(ByVal sPath As String) As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lLocation       As Long
    
    Const PROCEDURE_NAME As String = "GetFilenameFromPath"
    
    On Error GoTo ERROR_HANDLER
    
    lLocation = InStrRev(sPath, "\")
    
    If lLocation > 0 Then
        GetFilenameFromPath = Mid(sPath, lLocation + 1)
    Else
        GetFilenameFromPath = sPath
    End If
TIDY_UP:
    On Error Resume Next

    If lErrNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Function
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCEDURE_NAME)
    Resume TIDY_UP
End Function
