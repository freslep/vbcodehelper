Attribute VB_Name = "MRegistry"
'*******************************************************************************
' MODULE:       MRegistry
' FILENAME:     C:\My Code\vb\vbch\Source\MRegistry.bas
' AUTHOR:       Phil Fresle
' CREATED:      07-Mar-2002
' COPYRIGHT:    Copyright 2002-2019 Frez Systems Limited.
'*******************************************************************************
Option Explicit

Public Const REG_SZ     As Long = 1

Private Const MODULE_NAME As String = "MRegistry"

Private Const BASE_KEY  As String = "SOFTWARE"

Public Const REG_DWORD  As Long = 4

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
   (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long
    
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal ulOptions As Long, _
     ByVal samDesired As Long, _
     phkResult As Long) As Long
     
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, ByVal lpData As String, _
     lpcbData As Long) As Long
     
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Long, _
     lpcbData As Long) As Long
     
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     ByVal lpData As Long, _
     lpcbData As Long) As Long
     
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     ByVal lpValue As String, _
     ByVal cbData As Long) As Long
     
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpValue As Long, _
     ByVal cbData As Long) As Long

'*******************************************************************************
' SetValueEx (FUNCTION)
'*******************************************************************************
Private Function SetValueEx(ByVal hKey As Long, _
                           sValueName As String, _
                           lType As Long, _
                           vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function

'*******************************************************************************
' QueryValueEx (FUNCTION)
'*******************************************************************************
Private Function QueryValueEx(ByVal lhKey As Long, _
                              ByVal szValueName As String, _
                              vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then
        Error 5
    End If

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
            
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then
                vValue = lValue
            End If
        
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
Exit Function

QueryValueExError:
    Resume QueryValueExExit
End Function

'*******************************************************************************
' SaveStringSetting (SUB)
'*******************************************************************************
Public Sub SaveStringSetting(ByVal sAppName As String, _
                             ByVal sSection As String, _
                             ByVal sKey As String, _
                             ByVal sSetting As String)
    Dim lErrNumber      As Long
    Dim sErrDescription As String
    Dim sErrSource      As String
    
    On Error Resume Next
    
    Err.Clear
    
    ' Try local machine first
    SaveStringSettingPrivate sAppName, sSection, sKey, sSetting, HKEY_LOCAL_MACHINE
    
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ERROR_HANDLER
        
        ' If local machine failed, try current user
        SaveStringSettingPrivate sAppName, sSection, sKey, sSetting, HKEY_CURRENT_USER
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
    sErrSource = Err.Source
    sErrDescription = Err.Description
    Resume TIDY_UP
End Sub

'*******************************************************************************
' GetStringSetting (FUNCTION)
'*******************************************************************************
Public Function GetStringSetting(ByVal sAppName As String, _
                                 ByVal sSection As String, _
                                 ByVal sKey As String, _
                                 Optional ByVal sDefault As String) As String
    Dim bDefaultUsed    As Boolean
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "GetStringSetting"
    
    On Error Resume Next
    
    Err.Clear
    
    ' Try current user first
    GetStringSetting = GetStringSettingPrivate(sAppName, sSection, sKey, sDefault, HKEY_CURRENT_USER, bDefaultUsed)
    
    If Err.Number <> 0 Or bDefaultUsed Then
        Err.Clear
        On Error GoTo ERROR_HANDLER
        
        ' If current user failed, try local machine
        GetStringSetting = GetStringSettingPrivate(sAppName, sSection, sKey, sDefault, HKEY_LOCAL_MACHINE, bDefaultUsed)
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
' GetStringSettingPrivate (FUNCTION)
'*******************************************************************************
Public Function GetStringSettingPrivate(ByVal sAppName As String, _
                                        ByVal sSection As String, _
                                        ByVal sKey As String, _
                                        ByVal sDefault As String, _
                                        ByVal lMainKey As Long, _
                                        bDefaultUsed As Boolean) As String
    Dim lRetVal         As Long
    Dim sFullKey        As String
    Dim lHandle         As Long
    Dim sValue          As String
    Dim lErrNumber      As Long
    Dim sErrDescription As String
    Dim sErrSource      As String
    
    On Error GoTo ERROR_HANDLER
    
    bDefaultUsed = False

    If Trim(sAppName) = "" Then
        Err.Raise vbObjectError + 1000, , "AppName may not be empty"
    End If
    If Trim(sSection) = "" Then
        Err.Raise vbObjectError + 1001, , "Section may not be empty"
    End If
    If Trim(sKey) = "" Then
        Err.Raise vbObjectError + 1002, , "Key may not be empty"
    End If
    
    sFullKey = BASE_KEY & "\" & Trim(sAppName) & "\" & Trim(sSection)

    ' Open up the key
    lRetVal = RegOpenKeyEx(lMainKey, sFullKey, 0, KEY_QUERY_VALUE, lHandle)
    If lRetVal <> ERROR_NONE Then
        If lRetVal = ERROR_BADKEY Then
            bDefaultUsed = True
            GetStringSettingPrivate = sDefault
            Exit Function
        Else
            Err.Raise vbObjectError + 2000 + lRetVal, , _
                "Could not open registry section"
        End If
    End If
    
    lRetVal = QueryValueEx(lHandle, sKey, sValue)
    If lRetVal = ERROR_NONE Then
        GetStringSettingPrivate = sValue
    Else
        bDefaultUsed = True
        GetStringSettingPrivate = sDefault
    End If
TIDY_UP:
    On Error Resume Next
    
    RegCloseKey lHandle
    
    If lErrNumber <> 0 Then
        On Error GoTo 0
        
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Function

ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrSource = Err.Source
    sErrDescription = Err.Description
    Resume TIDY_UP
End Function

'*******************************************************************************
' SaveStringSettingPrivate (SUB)
'*******************************************************************************
Private Sub SaveStringSettingPrivate(ByVal sAppName As String, _
                                     ByVal sSection As String, _
                                     ByVal sKey As String, _
                                     ByVal sSetting As String, _
                                     ByVal lMainKey As Long)
    Dim lRetVal         As Long
    Dim sNewKey         As String
    Dim lDisposition    As Long
    Dim lHandle         As Long
    Dim lErrNumber      As Long
    Dim sErrDescription As String
    Dim sErrSource      As String
    
    On Error GoTo ERROR_HANDLER
    
    If Trim(sAppName) = "" Then
        Err.Raise vbObjectError + 1000, , "AppName may not be empty"
    End If
    If Trim(sSection) = "" Then
        Err.Raise vbObjectError + 1001, , "Section may not be empty"
    End If
    If Trim(sKey) = "" Then
        Err.Raise vbObjectError + 1002, , "Key may not be empty"
    End If
    
    sNewKey = BASE_KEY & "\" & Trim(sAppName) & "\" & Trim(sSection)
    
    ' Create the key or open it if it already exists
    lRetVal = RegCreateKeyEx(lMainKey, sNewKey, 0, vbNullString, 0, _
        KEY_ALL_ACCESS, 0, lHandle, lDisposition)
        
    If lRetVal <> ERROR_NONE Then
        Err.Raise vbObjectError + 2000 + lRetVal, , _
            "Could not open/create registry section"
    End If
    
    ' Set the key value
    lRetVal = SetValueEx(lHandle, sKey, REG_SZ, sSetting)
    
    If lRetVal <> ERROR_NONE Then
        Err.Raise vbObjectError + 2000 + lRetVal, , "Could not set key value"
    End If
    
TIDY_UP:
    On Error Resume Next
    
    RegCloseKey lHandle
    
    If lErrNumber <> 0 Then
        On Error GoTo 0
        
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Sub

ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrSource = Err.Source
    sErrDescription = Err.Description
    Resume TIDY_UP
End Sub
