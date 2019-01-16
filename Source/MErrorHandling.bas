Attribute VB_Name = "MErrorHandling"
'*******************************************************************************
' MODULE:       MErrorHandling
' FILENAME:
' AUTHOR:       Phil Fresle
' CREATED:
' COPYRIGHT:    Copyright 2001-2019 Frez Systems Limited
'
' DESCRIPTION:
' Utility functions used in handling errors
'
' MODIFICATION HISTORY:
' 1.0       17-Jan-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

'*******************************************************************************
' FormatErrorSource (FUNCTION)
'
' PARAMETERS:
' (In) - sErrSource - String - The current Err.Source
' (In) - sModule    - String - The module where it occurred
' (In) - sFunction  - String - The function where it occurred
'
' RETURN VALUE:
' String - A formatted string to return as the error source
'
' DESCRIPTION:
' Takes the error source and appends module and function information so it can
' be used to trace the stack.
'*******************************************************************************
Public Function FormatErrorSource(ByVal sErrSource As String, _
                                  ByVal sModule As String, _
                                  ByVal sFunction As String) As String
    
    Static s_sDefaultErrorSource As String
    
    On Error Resume Next
    
    If LenB(s_sDefaultErrorSource) = 0 Then
        Err.Raise vbObjectError
        s_sDefaultErrorSource = Err.Source
        Err.Clear
    End If
    
    If sErrSource = s_sDefaultErrorSource Then
        FormatErrorSource = App.ProductName & "." & sModule & "." & sFunction
    Else
        FormatErrorSource = sErrSource & vbCrLf & _
            App.ProductName & "." & sModule & "." & sFunction
    End If
End Function

'*******************************************************************************
' ShowUnexpectedError (SUB)
'
' PARAMETERS:
' (In) - lNumber      - Long   - Error number
' (In) - sDescription - String - Error description
' (In) - sLocation    - String - Error source
'
' DESCRIPTION:
' If an unexpected error is found, this routine gets called to display the error
' to the user.
'*******************************************************************************
Public Sub ShowUnexpectedError(ByVal lNumber As Long, _
                               ByVal sDescription As String, _
                               ByVal sLocation As String)
                               
    Dim sMessage As String
    
    On Error Resume Next
    
    sMessage = "Unexpected error found " & lNumber & vbCrLf & _
        sDescription & vbCrLf & _
        sLocation
    Debug.Print sMessage
    MsgBox sMessage, vbCritical, App.ProductName
End Sub
