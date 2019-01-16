Attribute VB_Name = "MCodeHelperOnly"
'*******************************************************************************
' MODULE:       MCodeHelperOnly
' FILENAME:     C:\My Code\vb\vbch\Source\MCodeHelperOnly.bas
' AUTHOR:       Phil Fresle
' CREATED:      06-Jul-2001
' COPYRIGHT:    Copyright 2001-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Stuff unique to the DLL
'
' MODIFICATION HISTORY:
' 1.0       06-Jul-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private m_lMenuHandle As Long

Private Const MODULE_NAME As String = "MCodeHelperOnly"

Public Const USER_TAGS                 As String = "USER_TOKENS.dat"
Public Const USER_TAG_DELIMITER        As String = ";"
Public Const TAG_QUOTE                 As String = "%"

Public Const TAG_PROJECTNAME           As String = "%PROJECTNAME%"
Public Const TAG_PROJECTDESCRIPTION    As String = "%PROJECTDESCRIPTION%"
Public Const TAG_MODULE                As String = "%MODULE%"
Public Const TAG_FILENAME              As String = "%FILENAME%"
Public Const TAG_AUTHOR                As String = "%AUTHOR%"
Public Const TAG_CREATED               As String = "%CREATED%"
Public Const TAG_COMPANYNAME           As String = "%COMPANYNAME%"
Public Const TAG_YEAR                  As String = "%YEAR%"
Public Const TAG_DATE                  As String = "%DATE%"
Public Const TAG_PROCNAME              As String = "%PROCNAME%"
Public Const TAG_PROCTYPE              As String = "%PROCTYPE%"
Public Const TAG_PROCTYPEMAIN          As String = "%PROCTYPEMAIN%"
Public Const TAG_PROPERTYTYPE          As String = "%PROPERTYTYPE%"
Public Const TAG_PARAMS                As String = "%PARAMS%"
Public Const TAG_RETVAL                As String = "%RETVAL%"
Public Const TAG_TIMESTAMP             As String = "%TIMESTAMP%"
Public Const TAG_INITIALS              As String = "%INITIALS%"
Public Const TAG_SCOPE                 As String = "%SCOPE%"
Public Const TAG_MAJORVER              As String = "%MAJORVER%"
Public Const TAG_MINORVER              As String = "%MINORVER%"
Public Const TAG_REVISIONVER           As String = "%REVISIONVER%"
' **** START CHANGE - PMF - 08-Nov-2005 07:25 ****
' New tags
Public Const TAG_FIRSTCODELINE         As String = "%FIRSTCODELINE%"
Public Const TAG_LASTCODELINE          As String = "%LASTCODELINE%"
Public Const TAG_SELECTION             As String = "%SELECTION%"
Public Const TAG_PROCTYPE2             As String = "%PROCTYPE2%"
Public Const TAG_DESCRIPTION           As String = "%DESCRIPTION%"
Public Const TAG_REMARK                As String = "%REMARK%"
Public Const TAG_MAXOBJERROR           As String = "%MAXOBJERROR%"
Public Const TAG_NEXTOBJERROR          As String = "%NEXTOBJERROR"
' **** END CHANGE - PMF - 07:25 ****

Public Const TOOLBAR_NAME              As String = "VBCodeHelperButtons"
Public Const MENU_NAME                 As String = "VB&CodeHelper"

Public Const TEMPLATE_WILDCARD         As String = "*.tlt"

Public Const GUID_TAB                  As String = "{5795A19D-CA7B-4399-AF05-AD78175FFEDF}"
Public Const GUID_ZORDER               As String = "{95285F7B-2861-4575-84AB-88ACA5D8ACAE}"
Public Const GUID_STATS                As String = "{BB942597-DBD4-473f-BF0D-DA684B8554F7}"
Public Const GUID_TEMPLATES            As String = "{3F65CA59-39FE-47a7-AD44-BD9A52B2E9DB}"
Public Const GUID_PROCEDURE_LIST       As String = "{03FDA3D0-19DB-4b88-8983-082EC611FFD8}"

Public Const RES_ID_MODULE             As Long = 101
Public Const RES_ID_PROCEDURE          As Long = 102
Public Const RES_ID_TIMESTAMP          As Long = 103
Public Const RES_ID_CLOSE              As Long = 104
Public Const RES_ID_CLEAR              As Long = 105
Public Const RES_ID_ERROR1             As Long = 106
Public Const RES_ID_ERROR2             As Long = 107
Public Const RES_ID_CONFIGURE          As Long = 108
Public Const RES_ID_HELP               As Long = 109
Public Const RES_ID_DIM                As Long = 110
Public Const RES_ID_DIM_ALL            As Long = 111
Public Const RES_ID_ERROR_ALL          As Long = 112
Public Const RES_ID_INDENT             As Long = 113
Public Const RES_ID_INDENT_ALL         As Long = 114
Public Const RES_ID_TAB                As Long = 115
Public Const RES_ID_ZORDER             As Long = 116
Public Const RES_ID_STATS              As Long = 117
Public Const RES_ID_WHITE_SPACE        As Long = 118
Public Const RES_ID_WHITE_SPACE_ALL    As Long = 119
Public Const RES_ID_PROCEDURE_LIST     As Long = 120

Public Const HEADER_SUB                As String = "HEADER_SUB.tld"
Public Const HEADER_FUNCTION           As String = "HEADER_FUNCTION.tld"
Public Const HEADER_SET                As String = "HEADER_SET.tld"
Public Const HEADER_LET                As String = "HEADER_LET.tld"
Public Const HEADER_GET                As String = "HEADER_GET.tld"
Public Const HEADER_MODULE             As String = "HEADER_MODULE.tld"
Public Const TIMESTAMP                 As String = "TIMESTAMP.tld"
Public Const TIMESTAMP_START           As String = "TIMESTAMP_START.tld"
Public Const TIMESTAMP_END             As String = "TIMESTAMP_END.tld"
Public Const ERROR1_START              As String = "ERROR1_START.tld"
Public Const ERROR1_END                As String = "ERROR1_END.tld"
Public Const ERROR2_START              As String = "ERROR2_START.tld"
Public Const ERROR2_END                As String = "ERROR2_END.tld"

Public Const ERROR_NUMBER_1            As Long = 1
Public Const ERROR_NUMBER_2            As Long = 2

Public Const BUTTON_TAG_MODULE          As String = "VBCH:Document Entire Module"
Public Const BUTTON_TAG_PROCEDURE       As String = "VBCH:Document Procedure"
Public Const BUTTON_TAG_TIME_STAMP      As String = "VBCH:Insert Time Stamp"
Public Const BUTTON_TAG_TEMPLATE        As String = "VBCH:Code Template"
Public Const BUTTON_TAG_CLOSE           As String = "VBCH:Close"
Public Const BUTTON_TAG_CLEAR           As String = "VBCH:Clear"
Public Const BUTTON_TAG_ERROR1          As String = "VBCH:Error1"
Public Const BUTTON_TAG_ERROR2          As String = "VBCH:Error2"
Public Const BUTTON_TAG_DIM             As String = "VBCH:Dim"
Public Const BUTTON_TAG_DIM_ALL         As String = "VBCH:DimAll"
Public Const BUTTON_TAG_ERROR_ALL       As String = "VBCH:ErrorAll"
Public Const BUTTON_TAG_INDENT          As String = "VBCH:Indent"
Public Const BUTTON_TAG_INDENT_ALL      As String = "VBCH:IndentAll"
Public Const BUTTON_TAG_TAB             As String = "VBCH:Tab"
Public Const BUTTON_TAG_ZORDER          As String = "VBCH:ZOrder"
Public Const BUTTON_TAG_STATS           As String = "VBCH:Stats"
Public Const BUTTON_TAG_WHITE_SPACE     As String = "VBCH:WhiteSpace"
Public Const BUTTON_TAG_WHITE_SPACE_ALL As String = "VBCH:WhiteSpaceAll"
Public Const BUTTON_TAG_PROCEDURE_LIST  As String = "VBCH:ProcedureList"

'public Const MENU_TAG_CONFIGURE        As String = "VBCHM:Configure"
Public Const MENU_TAG_MAIN             As String = "VBCHM:Main"

Public Const MENU_TAG_MODULE           As String = "VBCHM:Document Entire Module"
Public Const MENU_TAG_PROCEDURE        As String = "VBCHM:Document Procedure"
Public Const MENU_TAG_TIME_STAMP       As String = "VBCHM:Insert Time Stamp"
Public Const MENU_TAG_TEMPLATE         As String = "VBCHM:Code Template"
Public Const MENU_TAG_CLOSE            As String = "VBCHM:Close"
Public Const MENU_TAG_CLEAR            As String = "VBCHM:Clear"
Public Const MENU_TAG_ERROR1           As String = "VBCHM:Error1"
Public Const MENU_TAG_ERROR2           As String = "VBCHM:Error2"
Public Const MENU_TAG_CONFIGURE        As String = "VBCHM:Configure"
Public Const MENU_TAG_ABOUT            As String = "VBCHM:About"
Public Const MENU_TAG_HELP             As String = "VBCHM:Help"
Public Const MENU_TAG_DIM              As String = "VBCHM:Dim"
Public Const MENU_TAG_DIM_ALL          As String = "VBCHM:DimAll"
Public Const MENU_TAG_ERROR_ALL        As String = "VBCHM:ErrorAll"
Public Const MENU_TAG_INDENT           As String = "VBCHM:Indent"
Public Const MENU_TAG_INDENT_ALL       As String = "VBCHM:IndentAll"
Public Const MENU_TAG_TAB              As String = "VBCHM:Tab"
Public Const MENU_TAG_ZORDER           As String = "VBCHM:ZOrder"
Public Const MENU_TAG_STATS            As String = "VBCHM:Stats"
Public Const MENU_TAG_WHITE_SPACE      As String = "VBCHM:WhiteSpace"
Public Const MENU_TAG_WHITE_SPACE_ALL  As String = "VBCHM:WhiteSpaceAll"
Public Const MENU_TAG_PROCEDURE_LIST   As String = "VBCHM:ProcedureList"

Public Const WM_SYSKEYDOWN  As Long = &H104
Public Const WM_SYSKEYUP    As Long = &H105
Public Const WM_SYSCHAR     As Long = &H106
Public Const VK_F           As Long = 70

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Private Declare Function SetFocus Lib "user32" _
    (ByVal hwnd As Long) As Long

Private Declare Function GetParent Lib "user32" _
    (ByVal hwnd As Long) As Long

Public Enum enumProcedureType
    PTUnknown = 0
    PTSub = 1
    PTFunction = 2
    PTGet = 3
    PTLet = 4
    PTSet = 5
End Enum

Private Const OF_READ               As Long = &H0
Private Const OF_SHARE_DENY_NONE    As Long = &H40
Private Const OFS_MAXPATHNAME       As Long = 128

Private Type OFSTRUCTREC
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIMEREC
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIMEREC
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToSystemTime Lib "kernel32" _
    (lpFileTime As FILETIMEREC, _
     lpSystemTime As SYSTEMTIMEREC) As Long

Private Declare Function GetFileTime Lib "kernel32" _
    (ByVal hFile As Long, _
     lpCreationTime As FILETIMEREC, _
     lpLastAccessTime As FILETIMEREC, _
     lpLastWriteTime As FILETIMEREC) As Long

Private Declare Function OpenFile Lib "kernel32" _
    (ByVal lpFileName As String, _
     lpReOpenBuff As OFSTRUCTREC, _
     ByVal wStyle As Long) As Long

Private Declare Function hread Lib "kernel32" Alias "_hread" _
    (ByVal hFile As Long, _
     lpBuffer As Any, _
     ByVal lBytes As Long) As Long

Private Declare Function lclose Lib "kernel32" Alias "_lclose" _
    (ByVal hFile As Long) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" _
    (lpFileTime As FILETIMEREC, _
     lpLocalFileTime As FILETIMEREC) As Long

'*******************************************************************************
' CleanUpLine (FUNCTION)
'*******************************************************************************
Public Function CleanUpLine(ByVal sLine As String) As String
    Dim lQuoteCount     As Long
    Dim lCount          As Long
    Dim sChar           As String
    Dim sPrevChar       As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "CleanUpLine"
    
    On Error GoTo ERROR_HANDLER
    
    ' Starts with Rem it is a comment
    sLine = Trim(sLine)
    If Left(sLine, 3) = "Rem" Then
        CleanUpLine = ""
        Exit Function
    End If
    
    ' Starts with ' it is a comment
    If Left(sLine, 1) = "'" Then
        CleanUpLine = ""
        Exit Function
    End If
    
    ' Contains ' may end in a comment, so test if it is a comment or in the
    ' body of a string
    If InStr(sLine, " '") > 0 Then
        sPrevChar = " "
        lQuoteCount = 0
        
        For lCount = 1 To Len(sLine)
            sChar = Mid(sLine, lCount, 1)
            
            ' If we found " '" then an even number of " characters in front
            ' means it is the start of a comment, and odd number means it is
            ' part of a string
            If sChar = "'" And sPrevChar = " " Then
                If lQuoteCount Mod 2 = 0 Then
                    sLine = Trim(Left(sLine, lCount - 1))
                    Exit For
                End If
            ElseIf sChar = """" Then
                lQuoteCount = lQuoteCount + 1
            End If
            sPrevChar = sChar
        Next lCount
    End If
        
    CleanUpLine = sLine
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
' GetLineNumber (FUNCTION)
'*******************************************************************************
Public Function GetLineNumber(ByVal sLine As String) As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim vSplit()        As String
    Dim lCount          As Long
    Dim sTest           As String
    
    Const PROCEDURE_NAME As String = "GetLineNumber"
    
    On Error GoTo ERROR_HANDLER
    
    ' Remove leading and trailing spaces
    sLine = Trim(sLine)
    
    If sLine = "" Then
        GetLineNumber = ""
    Else
        vSplit = Split(sLine, " ")
        sTest = Trim(vSplit(0))
        
        If Right(sTest, 1) = ":" Then
            sTest = Left(sTest, Len(sTest) - 1)
        End If
        
        If IsNumeric(sTest) Then
            For lCount = 1 To Len(sTest)
                If Mid(sTest, lCount, 1) < "0" Or Mid(sTest, lCount, 1) > "9" Then
                    GetLineNumber = ""
                    Exit Function
                End If
            Next
            GetLineNumber = sTest
        Else
            GetLineNumber = ""
        End If
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

'*******************************************************************************
' TrimLine (FUNCTION)
'*******************************************************************************
Public Function TrimLine(ByVal sLine As String) As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim vSplit()        As String
    Dim lCount          As Long
    
    Const PROCEDURE_NAME As String = "TrimLine"
    
    On Error GoTo ERROR_HANDLER
    
    ' Remove leading and trailing spaces
    sLine = Trim(sLine)
    
    If sLine = "" Then
        TrimLine = ""
    Else
        vSplit = Split(sLine, " ")
        
        If IsNumeric(vSplit(0)) Then
            For lCount = 1 To Len(vSplit(0))
                If Mid(vSplit(0), lCount, 1) < "0" Or Mid(vSplit(0), lCount, 1) > "9" Then
                    TrimLine = sLine
                    Exit Function
                End If
            Next
            TrimLine = Trim(Mid(sLine, Len(vSplit(0)) + 1))
        Else
            TrimLine = sLine
        End If
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

'*******************************************************************************
' FindMenuHandle (FUNCTION)
'*******************************************************************************
Private Function FindMenuHandle(hwnd As Long) As Long
    Dim h As Long
    
    On Error Resume Next

RECURSE_LOOP:
    h = GetParent(hwnd)
    If h = 0 Then
        FindMenuHandle = hwnd
        Exit Function
    End If
    hwnd = h
    GoTo RECURSE_LOOP
End Function

'*******************************************************************************
' GetDeclarationLine (FUNCTION)
'
' PARAMETERS:
' (In) - lStartLine - Long - The first line of the declaration
'
' RETURN VALUE:
' String - The declaration
'
' DESCRIPTION:
' Given a line number, return the full declaration for the procedure the line
' number appear in
'*******************************************************************************
Public Function GetDeclarationLine(oVB As VBIDE.VBE, _
                                   ByVal lStartLine As Long, _
                                   lEndLine As Long) As String
    Dim sText           As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lRealStartLine  As Long
    
    Const PROCEDURE_NAME As String = "GetDeclarationLine"
    
    On Error GoTo ERROR_HANDLER
    
    lRealStartLine = lStartLine
    lEndLine = lStartLine
    
    With oVB.ActiveCodePane.CodeModule
        sText = ""
        Do
            If Right(sText, 2) = " _" Then
                sText = Left(sText, Len(sText) - 1)
            End If
            sText = sText & CleanUpLine(.Lines(lStartLine, 1))
            lStartLine = lStartLine + 1
        Loop Until Right(sText, 2) <> " _" Or lStartLine > .CountOfLines
    End With
    
    lEndLine = lStartLine - 1
    
    GetDeclarationLine = sText
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
' GetFileContents (FUNCTION)
'*******************************************************************************
Public Function GetFileContents(ByVal sFilename As String) As String
'    Dim oFSO            As FileSystemObject
'    Dim oTextStream     As TextStream
'    Dim sFileContents   As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lFile           As Long
    Dim sLines          As String
    
    Const PROCEDURE_NAME As String = "GetFileContents"
    
    On Error GoTo ERROR_HANDLER
    
'    Set oFSO = New FileSystemObject
'
'    If oFSO.FileExists(sFilename) Then
'        Set oTextStream = oFSO.OpenTextFile(sFilename)
'        sFileContents = oTextStream.ReadAll
'        oTextStream.Close
'        Set oTextStream = Nothing
'    End If
'    Set oFSO = Nothing
'
'    GetFileContents = sFileContents

    If Len(Dir(sFilename)) > 0 Then
        lFile = FreeFile
        
        Open sFilename For Input As #lFile
        
        sLines = StrConv(InputB(LOF(lFile), lFile), vbUnicode)
        
        Close #lFile
        
        GetFileContents = sLines
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

'*******************************************************************************
' GetFileCreationDate (FUNCTION)
'*******************************************************************************
Public Function GetFileCreationDate(sInpFile As String, sFormat As String) As String
    Dim hFile           As Integer
    Dim oFileStruct     As OFSTRUCTREC
    Dim iRC             As Integer
    Dim oCreationTime   As FILETIMEREC
    Dim oLastAccessTime As FILETIMEREC
    Dim oLastWriteTime  As FILETIMEREC
    Dim oLocalTime      As FILETIMEREC
    Dim oSystemTime     As SYSTEMTIMEREC
    Dim sDate           As String
    
    If Len(Dir(sInpFile)) > 0 Then
        hFile = OpenFile(sInpFile, oFileStruct, OF_READ Or OF_SHARE_DENY_NONE)
        
        If hFile <> 0 Then
            
            If GetFileTime(hFile, oCreationTime, oLastAccessTime, oLastWriteTime) <> 0 Then
                
                If FileTimeToLocalFileTime(oCreationTime, oLocalTime) <> 0 Then
                    
                    If FileTimeToSystemTime(oLocalTime, oSystemTime) <> 0 Then
                        sDate = Right("0000" & oSystemTime.wYear, 4) & "-"
                        sDate = sDate & Right("0" & oSystemTime.wMonth, 2) & "-"
                        sDate = sDate & Right("0" & oSystemTime.wDay, 2) & " "
                        sDate = sDate & Right("0" & oSystemTime.wHour, 2) & ":"
                        sDate = sDate & Right("0" & oSystemTime.wMinute, 2) & ":"
                        sDate = sDate & Right("0" & oSystemTime.wSecond, 2)
                        
                        GetFileCreationDate = Format(CDate(sDate), sFormat)
                    End If
                End If
            End If
            
            iRC = lclose(hFile)
        End If
    End If
End Function

'*******************************************************************************
' GetMethodType (FUNCTION)
'*******************************************************************************
Public Function GetMethodType(oVB As VBIDE.VBE, _
                              ByVal lTopLine As Long, _
                              sProcedureName As String) As enumProcedureType
    Dim sText           As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndLine        As Long
    
    Const PROCEDURE_NAME As String = "GetMethodType"
    
    On Error GoTo ERROR_HANDLER
    
    sText = GetDeclarationLine(oVB, lTopLine, lEndLine)

    If InStr(sText, "Sub " & sProcedureName & "(") > 0 Then
        GetMethodType = PTSub
    Else
        GetMethodType = PTFunction
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

'*******************************************************************************
' GetOtherTags (SUB)
'*******************************************************************************
Public Sub GetOtherTags(sAuthor As String, _
                        sCompany As String, _
                        sInitials As String, _
                        sTimeFormat As String, _
                        sDateFormat As String, _
                        bDocumentBeforeProcedure As Boolean, _
                        bCloseActive As Boolean)
                        
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim sDocBefore      As String
    Dim sCloseActive    As String
    
    Const PROCEDURE_NAME As String = "GetOtherTags"
    
    On Error GoTo ERROR_HANDLER
    
    'GetOwnerAndCompany sAuthor, sCompany
    
    sAuthor = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_AUTHOR, sAuthor)
    sCompany = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_COMPANY, sCompany)
    sInitials = GetInitials(sAuthor)
    sInitials = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_INITIALS, sInitials)
    sTimeFormat = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TIMEFORMAT, DEF_REG_TIMEFORMAT)
    sDateFormat = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_DATEFORMAT, DEF_REG_DATEFORMAT)
    sDocBefore = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_DOCBEFORE, DEF_REG_DOCBEFORE)
    If sDocBefore = "0" Then
        bDocumentBeforeProcedure = False
    Else
        bDocumentBeforeProcedure = True
    End If
    sCloseActive = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_CLOSEACTIVE, DEF_REG_CLOSEACTIVE)
    If sCloseActive = "0" Then
        bCloseActive = False
    Else
        bCloseActive = True
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
' GetPaths (SUB)
'*******************************************************************************
Public Sub GetPaths(sTemplatePath As String, _
                    sBoilerplatePath As String)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim sTempT          As String
    
    Const PROCEDURE_NAME As String = "GetPaths"
    
    On Error GoTo ERROR_HANDLER
 
    sTempT = App.Path
    If Right(sTempT, 1) = "\" Then
        sTempT = sTempT & "Templates"
    Else
        sTempT = sTempT & "\Templates"
    End If
    
    ' Get settings from registry
    sBoilerplatePath = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BOILERPLATES, sTempT))
    sTemplatePath = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TEMPLATES, sTempT))
    If Right(sBoilerplatePath, 1) <> "\" Then
        sBoilerplatePath = sBoilerplatePath & "\"
    End If
    If Right(sTemplatePath, 1) <> "\" Then
        sTemplatePath = sTemplatePath & "\"
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
' GetProcedureDetails (SUB)
'*******************************************************************************
Public Sub GetProcedureDetails(oVB As VBIDE.VBE, _
                               ByVal lStartLine As Long, _
                               oMember As Member, _
                               eptType As enumProcedureType, _
                               lTopLine As Long, _
                               lEndDecLine As Long)
    Dim sProcedureName  As String
    Dim lProcStartLine  As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lStartingLine   As Long
    Dim sText           As String
    
    Const PROCEDURE_NAME As String = "GetProcedureDetails"
    
    On Error GoTo ERROR_HANDLER

    With oVB.ActiveCodePane.CodeModule
        eptType = PTUnknown
        lTopLine = 0
        sProcedureName = .ProcOfLine(lStartLine, vbext_pk_Proc)
        
        If sProcedureName <> "" Then
            Set oMember = .Members(sProcedureName)
            
            If oMember.Type = vbext_mt_Method Then
                lTopLine = .ProcBodyLine(sProcedureName, vbext_pk_Proc)
                eptType = GetMethodType(oVB, lTopLine, sProcedureName)
            
            ElseIf oMember.Type = vbext_mt_Property Then
                On Error Resume Next
                
                ' Is it a Get, Set or Let?
                lProcStartLine = .ProcStartLine(sProcedureName, vbext_pk_Get)
                If Err.Number <> 0 Then
                    lProcStartLine = 0
                    Err.Clear
                End If
                If lProcStartLine > lTopLine And lProcStartLine <= lStartLine Then
                    lTopLine = .ProcBodyLine(sProcedureName, vbext_pk_Get)
                    eptType = PTGet
                End If
                
                lProcStartLine = .ProcStartLine(sProcedureName, vbext_pk_Let)
                If Err.Number <> 0 Then
                    lProcStartLine = 0
                    Err.Clear
                End If
                If lProcStartLine > lTopLine And lProcStartLine <= lStartLine Then
                    lTopLine = .ProcBodyLine(sProcedureName, vbext_pk_Let)
                    eptType = PTLet
                End If
                
                lProcStartLine = .ProcStartLine(sProcedureName, vbext_pk_Set)
                If Err.Number <> 0 Then
                    lProcStartLine = 0
                    Err.Clear
                End If
                If lProcStartLine > lTopLine And lProcStartLine <= lStartLine Then
                    lTopLine = .ProcBodyLine(sProcedureName, vbext_pk_Set)
                    eptType = PTSet
                End If
                On Error GoTo 0
            End If
            
            lEndDecLine = lTopLine
            lStartingLine = lTopLine
            Do
                sText = CleanUpLine(.Lines(lStartingLine, 1))
                lStartingLine = lStartingLine + 1
            Loop Until Right(sText, 2) <> " _" Or lStartingLine > .CountOfLines
    
            lEndDecLine = lStartingLine - 1
        Else
            Set oMember = Nothing
        End If
    End With
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
' GetUserTags (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Read user-defined tags file. This contains tags in the format:
' TAG_NAME;TAG_PROMPT;TAG_VALUE
' For instance:
'  %SUB_NAME%;Enter the sub procedure name
'  %DEPT%;;Personnel Department
'  %SECTION%;Enter section;IT Section
' With these examples, entry of %SUB_NAME% in a template would result in a
' prompt for the sub procedure name whenever the template was inserted;
' entry of %DEPT% in a template would result in the straight forward
' substitution for the text 'Personnel Department';
' entry of %SECTION% in a template would result in a prompt for the section
' whenever the template was inserted but with a default of 'IT Section'
'*******************************************************************************
Public Sub GetUserTags(colUserTags As Collection)
    Dim sUserTokensPath     As String
    Dim sUserTagContents    As String
    Dim lCount              As Long
    Dim oUserTag            As CUserTag
    Dim sUserTags()         As String
    Dim sTagParts()         As String
    Dim sTag                As String
    Dim lLower              As Long
    Dim lErrNumber          As Long
    Dim sErrSource          As String
    Dim sErrDescription     As String
    Dim sTemplate           As String
    Dim sBoilerplate        As String
    
    Const PROCEDURE_NAME As String = "GetUserTags"
    
    On Error GoTo ERROR_HANDLER
    
    GetPaths sTemplate, sBoilerplate
    
    sUserTokensPath = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_USER_TOKENS, sTemplate))
    If Right(sUserTokensPath, 1) <> "\" Then
        sUserTokensPath = sUserTokensPath & "\"
    End If
            
    ' Start with an empty collection
    If Not colUserTags Is Nothing Then
        For lCount = 1 To colUserTags.Count
            colUserTags.Remove 1
        Next
        Set colUserTags = Nothing
    End If
    Set colUserTags = New Collection
    
    ' Read in the user tags file
    sUserTagContents = Trim(GetFileContents(sUserTokensPath & USER_TAGS))
    If sUserTagContents <> "" Then
        sUserTags = Split(sUserTagContents, vbCrLf)
        
        ' Go through line by line
        For lCount = LBound(sUserTags) To UBound(sUserTags)
            sTag = Trim(sUserTags(lCount))
            
            ' If not an empty line
            If sTag <> "" Then
            
                ' If not a comment line
                If Left(sTag, 1) <> "'" Then
                
                    ' If it is a valid tag format, remember it
                    If InStr(sTag, USER_TAG_DELIMITER) > 0 Then
                        sTagParts = Split(sTag, USER_TAG_DELIMITER)
                        Set oUserTag = New CUserTag
                        With oUserTag
                            lLower = LBound(sTagParts)
                            .TagName = Trim(sTagParts(lLower))
                            .TagPrompt = Trim(sTagParts(lLower + 1))
                            If UBound(sTagParts) >= lLower + 2 Then
                                .TagValue = Trim(sTagParts(lLower + 2))
                            Else
                                .TagValue = ""
                            End If
                            If Len(.TagName) <> 0 And Len(.TagName) > 2 Then
                                If Left(.TagName, 1) = TAG_QUOTE And Right(.TagName, 1) = TAG_QUOTE Then
                                    colUserTags.Add oUserTag
                                End If
                            End If
                        End With
                    End If
                End If
            End If
        Next
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

'*******************************************************************************
' GetVersionInfo (SUB)
'*******************************************************************************
Public Sub GetVersionInfo(oVB As VBIDE.VBE, _
                          sMajorVer As String, _
                          sMinorVer As String, _
                          sRevisionVer As String)
                           
'    Dim oFSO            As FileSystemObject
'    Dim oTS             As TextStream
    Dim sProject        As String
    Dim sContents       As String
    Dim sLines()        As String
    Dim lCount          As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROJ_MAJOR_VER    As String = "MajorVer="
    Const PROJ_MINOR_VER    As String = "MinorVer="
    Const PROJ_REVISION_VER As String = "RevisionVer="
    
    Const PROCEDURE_NAME As String = "GetVersionInfo"
    
    On Error GoTo ERROR_HANDLER
    
    sProject = oVB.ActiveVBProject.FileName
    
    sMajorVer = "1"
    sMinorVer = "0"
    sRevisionVer = "0"
        
    If sProject <> "" Then
        If Dir(sProject) <> "" Then
'            Set oFSO = New FileSystemObject
'            Set oTS = oFSO.OpenTextFile(sProject)
'            sContents = oTS.ReadAll
'            oTS.Close
            
            sContents = GetFileContents(sProject)
            
            sLines = Split(sContents, vbCrLf)
            
            For lCount = 0 To UBound(sLines)
                If Left(sLines(lCount), Len(PROJ_MAJOR_VER)) = PROJ_MAJOR_VER Then
                    sMajorVer = Mid(sLines(lCount), Len(PROJ_MAJOR_VER) + 1)
                ElseIf Left(sLines(lCount), Len(PROJ_MINOR_VER)) = PROJ_MINOR_VER Then
                    sMinorVer = Mid(sLines(lCount), Len(PROJ_MINOR_VER) + 1)
                ElseIf Left(sLines(lCount), Len(PROJ_REVISION_VER)) = PROJ_REVISION_VER Then
                    sRevisionVer = Mid(sLines(lCount), Len(PROJ_REVISION_VER) + 1)
                End If
            Next
            
'            Set oTS = Nothing
'            Set oFSO = Nothing
        End If
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
' HandleKeyDown (SUB)
'*******************************************************************************
Public Sub HandleKeyDown(oVB As VBIDE.VBE, _
                         oUserDocument As Object, _
                         iKeyCode As Integer, _
                         iShift As Integer)
    On Error Resume Next
    
    If iShift <> vbAltMask Then
        Exit Sub
    End If
    
    If iKeyCode < 65 Or iKeyCode > 90 Then
        Exit Sub
    End If
    
    If oVB.DisplayModel = vbext_dm_SDI Then
        Exit Sub
    End If
    
    If m_lMenuHandle = 0 Then
        m_lMenuHandle = FindMenuHandle(oUserDocument.hwnd)
    End If
    
    PostMessage m_lMenuHandle, WM_SYSKEYDOWN, iKeyCode, &H20000000
    iKeyCode = 0
    SetFocus m_lMenuHandle
    
    Err.Clear
End Sub

'*******************************************************************************
' InRunMode (FUNCTION)
'*******************************************************************************
Public Function InRunMode(oVB As VBIDE.VBE) As Boolean
    Dim oMenuBar As CommandBar
    
    On Error Resume Next
    Err.Clear
    Set oMenuBar = oVB.CommandBars("Menu Bar")
    
    If Err.Number <> 0 Then
        Set oMenuBar = oVB.CommandBars(1)
        
        If oMenuBar.Type <> msoBarTypeMenuBar Then
            For Each oMenuBar In oVB.CommandBars
                If oMenuBar.Type = msoBarTypeMenuBar Then
                    Exit For
                End If
            Next
        End If
    End If
    InRunMode = (oMenuBar.Controls(1).Controls(1).Enabled = False)
'    InRunMode = (oVB.CommandBars("File").Controls(1).Enabled = False)
End Function

'*******************************************************************************
' ReplaceProcedureTags (SUB)
'*******************************************************************************
Public Sub ReplaceProcedureTags(sSource As String, _
                                ByVal eptType As enumProcedureType, _
                                ByVal oMember As Member)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "ReplaceProcedureTags"
    
    On Error GoTo ERROR_HANDLER
    
    ' Make sure we have at least one % otherwise it is not worth searching for tags
    If InStr(sSource, TAG_QUOTE) > 0 Then
        
        If InStr(sSource, TAG_PROCTYPE) > 0 Then
            Select Case eptType
                Case PTSub
                    ReplaceTag sSource, TAG_PROCTYPE, "SUB"
                    ReplaceTag sSource, TAG_PROCTYPE2, "SUB"
                Case PTFunction
                    ReplaceTag sSource, TAG_PROCTYPE, "FUNCTION"
                    ReplaceTag sSource, TAG_PROCTYPE2, "FUNCTION"
                Case PTGet
                    ReplaceTag sSource, TAG_PROCTYPE, "PROPERTY GET"
                    ReplaceTag sSource, TAG_PROCTYPE2, "PROPERTY"
                Case PTLet
                    ReplaceTag sSource, TAG_PROCTYPE, "PROPERTY LET"
                    ReplaceTag sSource, TAG_PROCTYPE2, "PROPERTY"
                Case PTSet
                    ReplaceTag sSource, TAG_PROCTYPE, "PROPERTY SET"
                    ReplaceTag sSource, TAG_PROCTYPE2, "PROPERTY"
                Case Else
                    ReplaceTag sSource, TAG_PROCTYPE, "UNKNOWN"
                    ReplaceTag sSource, TAG_PROCTYPE2, "UNKNOWN"
            End Select
        End If
        
        If InStr(sSource, TAG_PROPERTYTYPE) > 0 Then
            Select Case eptType
                Case PTGet
                    ReplaceTag sSource, TAG_PROPERTYTYPE, "(GET)"
                Case PTLet
                    ReplaceTag sSource, TAG_PROPERTYTYPE, "(LET)"
                Case PTSet
                    ReplaceTag sSource, TAG_PROPERTYTYPE, "(SET)"
                Case Else
                    ReplaceTag sSource, TAG_PROPERTYTYPE, ""
            End Select
        End If
        
        If InStr(sSource, TAG_PROCTYPEMAIN) > 0 Then
            Select Case eptType
                Case PTSub
                    ReplaceTag sSource, TAG_PROCTYPEMAIN, "Sub"
                Case PTFunction
                    ReplaceTag sSource, TAG_PROCTYPEMAIN, "Function"
                Case PTGet, PTLet, PTSet
                    ReplaceTag sSource, TAG_PROCTYPEMAIN, "Property"
                Case Else
                    ReplaceTag sSource, TAG_PROCTYPEMAIN, ""
            End Select
        End If
        
        If InStr(sSource, TAG_SCOPE) > 0 Then
            Select Case oMember.Scope
                Case vbext_Private
                    ReplaceTag sSource, TAG_SCOPE, "PRIVATE"
                Case vbext_Public
                    ReplaceTag sSource, TAG_SCOPE, "PUBLIC"
                Case vbext_Friend
                    ReplaceTag sSource, TAG_SCOPE, "FRIEND"
                Case Else
                    ReplaceTag sSource, TAG_SCOPE, ""
            End Select
        End If
        
        If oMember Is Nothing Then
            ReplaceTag sSource, TAG_PROCNAME, ""
        Else
            ReplaceTag sSource, TAG_PROCNAME, oMember.Name
        End If
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
' ReplaceTag (SUB)
'*******************************************************************************
Public Sub ReplaceTag(sSource As String, _
                      sTag As String, _
                      sReplacement As String)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "ReplaceTag"
    
    On Error GoTo ERROR_HANDLER
    
    If InStr(sSource, sTag) > 0 Then
        sSource = Replace(sSource, sTag, sReplacement)
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
' ReplaceTags (SUB)
'*******************************************************************************
Public Sub ReplaceTags(oVB As VBIDE.VBE, _
                       sSource As String, _
                       sAuthor As String, _
                       sCompany As String, _
                       sDateFormat As String, _
                       sTimeFormat As String, _
                       sInitials As String, _
                       colUserTags As Collection)
                       
    Dim oTag            As CUserTag
'    Dim oFSO            As FileSystemObject
'    Dim oFile           As File
    Dim sValue          As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim sMajorVer       As String
    Dim sMinorVer       As String
    Dim sRevisionVer    As String
    
    Const PROCEDURE_NAME As String = "ReplaceTags"
    
    On Error GoTo ERROR_HANDLER
    
    ' Make sure we have at least one % otherwise it is not worth searching for tags
    If InStr(sSource, TAG_QUOTE) > 0 Then
    
        ReplaceTag sSource, TAG_PROJECTNAME, oVB.ActiveVBProject.Name
        ReplaceTag sSource, TAG_PROJECTDESCRIPTION, oVB.ActiveVBProject.Description
        ReplaceTag sSource, TAG_MODULE, oVB.ActiveCodePane.CodeModule.Name
        ReplaceTag sSource, TAG_FILENAME, oVB.SelectedVBComponent.FileNames(1)
        ReplaceTag sSource, TAG_AUTHOR, sAuthor
        ReplaceTag sSource, TAG_COMPANYNAME, sCompany
        ReplaceTag sSource, TAG_YEAR, Format(Date, "yyyy")
        ReplaceTag sSource, TAG_DATE, Format(Now, sDateFormat)
        ReplaceTag sSource, TAG_TIMESTAMP, Format(Now, sTimeFormat)
        ReplaceTag sSource, TAG_INITIALS, sInitials
        
        If InStr(sSource, TAG_CREATED) > 0 Then
            If oVB.SelectedVBComponent.FileNames(1) <> "" Then
'                Set oFSO = New FileSystemObject
'
'                Set oFile = oFSO.GetFile(oVB.SelectedVBComponent.FileNames(1))
'                ReplaceTag sSource, TAG_CREATED, Format(oFile.DateCreated, sDateFormat)
                ReplaceTag sSource, TAG_CREATED, GetFileCreationDate(oVB.SelectedVBComponent.FileNames(1), sDateFormat)
                
'                Set oFile = Nothing
'                Set oFSO = Nothing
            Else
                ReplaceTag sSource, TAG_CREATED, ""
            End If
        End If
                
        If InStr(sSource, TAG_MAJORVER) > 0 Or _
            InStr(sSource, TAG_MINORVER) > 0 Or _
            InStr(sSource, TAG_REVISIONVER) > 0 Then
            
            GetVersionInfo oVB, sMajorVer, sMinorVer, sRevisionVer
        
            ReplaceTag sSource, TAG_MAJORVER, sMajorVer
            ReplaceTag sSource, TAG_MINORVER, sMinorVer
            ReplaceTag sSource, TAG_REVISIONVER, sRevisionVer
        End If
        
        ' Replace all user-defined tags
        For Each oTag In colUserTags
            With oTag
                If InStr(sSource, .TagName) > 0 Then
                    If .TagPrompt = "" Then
                        sValue = .TagValue
                    Else
                        sValue = InputBox(.TagPrompt, App.ProductName, .TagValue)
                    End If
                    ReplaceTag sSource, .TagName, sValue
                End If
            End With
        Next
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
