VERSION 5.00
Begin VB.Form FRegistration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBCodeHelper Licence Registration"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "FRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   3315
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2580
      TabIndex        =   11
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3900
      TabIndex        =   12
      Top             =   1860
      Width           =   1215
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1380
      Width           =   675
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1380
      Width           =   675
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1380
      Width           =   675
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Registered name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4260
      TabIndex        =   9
      Top             =   1380
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3420
      TabIndex        =   7
      Top             =   1380
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2580
      TabIndex        =   5
      Top             =   1380
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your licence key below and the name you registered the product under."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      Caption         =   "Licence key:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1110
   End
End
Attribute VB_Name = "FRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FRegistration
' FILENAME:     C:\My Code\vb\vbch\Source\FRegistration.frm
' AUTHOR:       Phil Fresle
' CREATED:      20-Jan-2001
' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' Register the product
'
' MODIFICATION HISTORY:
' 1.0       20-Jan-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private Const MODULE_NAME As String = "FRegistration"

Private m_bNewLicense As Boolean

'*******************************************************************************
' cmdCancel_Click (SUB)
'*******************************************************************************
Private Sub cmdCancel_Click()
    On Error Resume Next
    
    Unload Me
End Sub

'*******************************************************************************
' cmdOK_Click (SUB)
'*******************************************************************************
Private Sub cmdOK_Click()
    Dim sKey            As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim sToken          As String
    Dim lTokenData      As Long
    Dim oRegistration   As CRegistration
    Dim sSoundex        As String
    Dim sPart1          As String
    Dim sChar1          As String
    Dim sChar3          As String
    Dim sChar5          As String
    
    Const PROCEDURE_NAME As String = "cmdOK_Click"
    
    On Error GoTo ERROR_HANDLER
    
    sKey = Trim(txtKey(0).Text) & Trim(txtKey(1).Text) & Trim(txtKey(2).Text) _
        & Trim(txtKey(3).Text)
    sKey = UCase(sKey)
    sKey = Replace(sKey, "O", "0")
    
    If sKey <> "" Then
        Set oRegistration = New CRegistration
    
        If oRegistration.IsKeyOK(sKey) Then
        
            sChar1 = Mid(sKey, 1, 1)
            sChar3 = Mid(sKey, 3, 1)
            sChar5 = Mid(sKey, 5, 1)
            
            ' Is it a component source key?
            If sChar1 = "8" And sChar3 = "9" And sChar5 = "7" Then
                SaveStringSetting FREZ_DATA, REG_APP_NAME, DATA_USER, Trim(txtName.Text)
                SaveStringSetting FREZ_DATA, REG_APP_NAME, DATA_LK, sKey
                
                sToken = MakeToken(LICENSE_PERM, Date, sKey & Trim(txtName.Text))
                
                g_sLicensedTo = txtName.Text
                g_sLicenseKey = sKey
                SaveStringSetting FREZ_DATA, REG_APP_NAME, DATA_KEY_WIDE, sToken
                MsgBox "The licence key has been loaded.", vbInformation, App.ProductName
                m_bNewLicense = True
                Unload Me
            Else
                sSoundex = oRegistration.Soundex(txtName.Text)
                sPart1 = Mid(sSoundex, 2, 3) & Asc(Left(sSoundex, 1))
                
                If Left(sKey, 5) = sPart1 Then
                    SaveStringSetting FREZ_DATA, REG_APP_NAME, DATA_USER, Trim(txtName.Text)
                    SaveStringSetting FREZ_DATA, REG_APP_NAME, DATA_LK, sKey
                    
                    sToken = MakeToken(LICENSE_PERM, Date, sKey & Trim(txtName.Text))
                    
                    g_sLicensedTo = txtName.Text
                    g_sLicenseKey = sKey
                    SaveStringSetting FREZ_DATA, REG_APP_NAME, DATA_KEY_WIDE, sToken
                    MsgBox "The licence key has been loaded.", vbInformation, App.ProductName
                    m_bNewLicense = True
                    Unload Me
                Else
                    MsgBox "The licence key you entered was invalid, please check it carefully and try again.", _
                        vbCritical, App.ProductName
                End If
            End If
        Else
            MsgBox "The licence key you entered was invalid, please check it carefully and try again.", _
                vbCritical, App.ProductName
        End If
        
        Set oRegistration = Nothing
    Else
        Exit Sub
    End If
Exit Sub
TIDY_UP:
    On Error Resume Next

    If lErrNumber <> 0 Then
        MsgBox "There was a problem processing the licence key, please check it carefully and try again.", _
            vbCritical, App.ProductName
    End If
Exit Sub
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCEDURE_NAME)
    Resume TIDY_UP
End Sub

'*******************************************************************************
' Form_Load (SUB)
'*******************************************************************************
Private Sub Form_Load()
    On Error Resume Next
    
    m_bNewLicense = False
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

'*******************************************************************************
' NewLicense (PROPERTY GET)
'*******************************************************************************
Public Property Get NewLicense() As Boolean
    NewLicense = m_bNewLicense
End Property

'*******************************************************************************
' txtKey_KeyPress (SUB)
'*******************************************************************************
Private Sub txtKey_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
