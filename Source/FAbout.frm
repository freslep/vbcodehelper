VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VBCodeHelper"
   ClientHeight    =   4695
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5910
   ClipControls    =   0   'False
   Icon            =   "FAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240.572
   ScaleMode       =   0  'User
   ScaleWidth      =   5549.795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   4560
      TabIndex        =   0
      Top             =   4140
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   $"FAbout.frx":0442
      Height          =   1035
      Left            =   60
      TabIndex        =   6
      Top             =   3000
      Width           =   5775
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
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
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   5925
   End
   Begin VB.Label lblLicenseType 
      Alignment       =   2  'Center
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
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   5925
   End
   Begin VB.Label lblLicenseKey 
      Alignment       =   2  'Center
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
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   5925
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      Caption         =   "Copyright 1999-2019 Frez Systems"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5925
   End
   Begin VB.Label lblAbout 
      Caption         =   "VBCodeHelper is a multi-function Add-In for Visual Basic 6.0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   1860
      Width           =   5775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.582
      Y2              =   1687.582
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.935
      Y2              =   1697.935
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FAbout
' FILENAME:     C:\My Code\vb\vbch\Source\FAbout.frm
' AUTHOR:       Phil Fresle
' CREATED:      12-Mar-2001
' COPYRIGHT:    Copyright 2001-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Standard 'about'
'
' MODIFICATION HISTORY:
' 1.0       13-Mar-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

'*******************************************************************************
' cmdOK_Click (SUB)
'*******************************************************************************
Private Sub cmdOK_Click()
    On Error Resume Next
    
    Unload Me
End Sub

'*******************************************************************************
' Form_Load (SUB)
'*******************************************************************************
Private Sub Form_Load()
    On Error Resume Next
    
    lblCopyright.Caption = App.LegalCopyright
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

'*******************************************************************************
' lblHyperlink_Click (SUB)
'*******************************************************************************
Private Sub lblHyperlink_Click()
    On Error Resume Next
    
    'ShellExecute Me.hwnd, "open", lblHyperlink.Caption, vbNullString, "C:\", SW_SHOWNORMAL
End Sub
