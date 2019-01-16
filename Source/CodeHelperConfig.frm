VERSION 5.00
Begin VB.Form FConfigure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBCodeHelper Configuration"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "CodeHelperConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optConfig 
      Caption         =   "About VBCodeHelper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame fraConfig 
      Height          =   5355
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   660
      Width           =   9255
      Begin VB.ComboBox cboParamFormat 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3780
         Width           =   2715
      End
      Begin VB.TextBox txtAuthor 
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   480
         Width           =   4515
      End
      Begin VB.TextBox txtCompany 
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1140
         Width           =   4515
      End
      Begin VB.TextBox txtInitials 
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   1800
         Width           =   675
      End
      Begin VB.TextBox txtTimeFormat 
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   2460
         Width           =   1815
      End
      Begin VB.TextBox txtDateFormat 
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Format to use for %PARAMS% when documenting parameters"
         Height          =   195
         Index           =   20
         Left            =   180
         TabIndex        =   21
         Top             =   3540
         Width           =   4365
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Name of author to substitute for the %AUTHOR% token:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   3960
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Name of company to substitute for the %COMPANYNAME% token:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   4725
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Initials of author to substitute for the %INITIALS% token:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   1560
         Width           =   3945
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Time format to use when substituting the %TIMESTAMP% token:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   2220
         Width           =   4575
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Date format to use when substituting the %DATE% and %CREATED% tokens:"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   19
         Top             =   2880
         Width           =   5505
      End
   End
   Begin VB.Frame fraConfig 
      Height          =   5355
      Index           =   2
      Left            =   60
      TabIndex        =   23
      Top             =   660
      Width           =   9255
      Begin VB.Frame fraCloseAll 
         Caption         =   "Close All Windows"
         Height          =   735
         Left            =   180
         TabIndex        =   51
         Top             =   4440
         Width           =   8895
         Begin VB.OptionButton optCloseActive 
            Caption         =   "Yes"
            Height          =   315
            Index           =   0
            Left            =   4860
            TabIndex        =   53
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optCloseActive 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   6000
            TabIndex        =   54
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Close the Active Window when Closing All Windows"
            Height          =   195
            Index           =   19
            Left            =   360
            TabIndex        =   52
            Top             =   300
            Width           =   3690
         End
      End
      Begin VB.Frame fraDoc 
         Caption         =   "Documentation Options"
         Height          =   735
         Left            =   180
         TabIndex        =   47
         Top             =   3540
         Width           =   8895
         Begin VB.OptionButton optBefore 
            Caption         =   "After"
            Height          =   315
            Index           =   1
            Left            =   6000
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optBefore 
            Caption         =   "Before"
            Height          =   315
            Index           =   0
            Left            =   4860
            TabIndex        =   49
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Add Documention Before or After Procedure Declaration?"
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   48
            Top             =   300
            Width           =   4065
         End
      End
      Begin VB.Frame fraWhiteSpace 
         Caption         =   "White Space Options"
         Height          =   735
         Left            =   180
         TabIndex        =   44
         Top             =   2700
         Width           =   8895
         Begin VB.TextBox txtWhite 
            Height          =   315
            Left            =   4860
            TabIndex        =   46
            Text            =   "1"
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Number of Consecutive Lines of White Space Allowed:"
            Height          =   195
            Index           =   16
            Left            =   360
            TabIndex        =   45
            Top             =   300
            Width           =   3885
         End
      End
      Begin VB.Frame fraIndent 
         Caption         =   "Options When Smart Indenting"
         Height          =   1035
         Left            =   180
         TabIndex        =   37
         Top             =   1620
         Width           =   8895
         Begin VB.TextBox txtIndent 
            Height          =   315
            Left            =   4860
            TabIndex        =   39
            Text            =   "4"
            Top             =   240
            Width           =   675
         End
         Begin VB.Frame fraIndentDim 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   4740
            TabIndex        =   41
            Top             =   480
            Width           =   3255
            Begin VB.OptionButton optIndentDim 
               Caption         =   "No"
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   43
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optIndentDim 
               Caption         =   "Yes"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   42
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Indent Declaration Statements (Dim, Static, Const):"
            Height          =   195
            Index           =   14
            Left            =   360
            TabIndex        =   40
            Top             =   660
            Width           =   3585
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Number of Spaces For Each Indent:"
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   38
            Top             =   300
            Width           =   2550
         End
      End
      Begin VB.Frame fraErrorOptions 
         Caption         =   "Options When Error Handling ALL Procedures"
         Height          =   1395
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   8895
         Begin VB.Frame fraProperty 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   4740
            TabIndex        =   34
            Top             =   840
            Width           =   3255
            Begin VB.OptionButton optProperty 
               Caption         =   "Yes"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton optProperty 
               Caption         =   "No"
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   36
               Top             =   120
               Width           =   675
            End
         End
         Begin VB.Frame fraEvent 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   4740
            TabIndex        =   30
            Top             =   480
            Width           =   3255
            Begin VB.OptionButton optEvent 
               Caption         =   "1"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   31
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optEvent 
               Caption         =   "2"
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   32
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
         End
         Begin VB.Frame fraNormal 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   4740
            TabIndex        =   26
            Top             =   120
            Width           =   3255
            Begin VB.OptionButton optNormal 
               Caption         =   "2"
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   28
               Top             =   120
               Width           =   675
            End
            Begin VB.OptionButton optNormal 
               Caption         =   "1"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   120
               Value           =   -1  'True
               Width           =   675
            End
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Add Error Handlers to Property Procedures:"
            Height          =   195
            Index           =   18
            Left            =   360
            TabIndex        =   33
            Top             =   1020
            Width           =   3045
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Error Handler if Procedure Name Contains Underscore (_):"
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   29
            Top             =   660
            Width           =   4095
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            Caption         =   "Normal Error Handler:"
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   25
            Top             =   300
            Width           =   1515
         End
      End
   End
   Begin VB.Frame fraConfig 
      Height          =   5355
      Index           =   4
      Left            =   60
      TabIndex        =   60
      Top             =   660
      Width           =   9255
      Begin VB.CheckBox chkShortcuts 
         Height          =   195
         Left            =   2040
         TabIndex        =   62
         Top             =   240
         Width           =   315
      End
      Begin VB.ListBox lstShortcuts 
         Height          =   2760
         Left            =   2040
         Style           =   1  'Checkbox
         TabIndex        =   64
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Enable Shortcuts"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Shortcuts to include"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   63
         Top             =   600
         Width           =   1410
      End
   End
   Begin VB.Frame fraConfig 
      Height          =   5355
      Index           =   3
      Left            =   60
      TabIndex        =   55
      Top             =   660
      Width           =   9255
      Begin VB.ListBox lstButtons 
         Height          =   2760
         Left            =   2040
         Style           =   1  'Checkbox
         TabIndex        =   59
         Top             =   600
         Width           =   4575
      End
      Begin VB.CheckBox chkToolbar 
         Height          =   195
         Left            =   2040
         TabIndex        =   57
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Buttons to include"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   58
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Show toolbar in IDE"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   56
         Top             =   240
         Width           =   1410
      End
   End
   Begin VB.Frame fraConfig 
      Height          =   5355
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   9255
      Begin VB.CommandButton cmdTemplateBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   8820
         TabIndex        =   3
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdBoilerplateBrowser 
         Caption         =   "..."
         Height          =   315
         Left            =   8820
         TabIndex        =   6
         Top             =   1140
         Width           =   315
      End
      Begin VB.CommandButton cmdUserTokens 
         Caption         =   "..."
         Height          =   315
         Left            =   8820
         TabIndex        =   9
         Top             =   1800
         Width           =   315
      End
      Begin VB.TextBox txtTemplatePath 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   8595
      End
      Begin VB.TextBox txtBoilerplatePath 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1140
         Width           =   8595
      End
      Begin VB.TextBox txtUserTokens 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   1800
         Width           =   8595
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Path to code templates (*.tlt):"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Path to boilerplates used to document code (*.tld):"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   4
         Top             =   900
         Width           =   3525
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Path to USER_TOKENS.dat"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   7
         Top             =   1560
         Width           =   2025
      End
   End
   Begin VB.OptionButton optConfig 
      Caption         =   "Shortcuts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton optConfig 
      Caption         =   "Toolbar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton optConfig 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton optConfig 
      Caption         =   "Tokens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
   End
   Begin VB.OptionButton optConfig 
      Caption         =   "File Locations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   0
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   5280
      TabIndex        =   71
      Top             =   6120
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   8040
      TabIndex        =   73
      Top             =   6120
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6660
      TabIndex        =   72
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Frame fraConfig 
      Height          =   5355
      Index           =   5
      Left            =   60
      TabIndex        =   65
      Top             =   660
      Width           =   9255
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
         Left            =   120
         TabIndex        =   68
         Top             =   1320
         Width           =   9045
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
         Left            =   120
         TabIndex        =   67
         Top             =   1080
         Width           =   9045
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
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   9045
      End
      Begin VB.Label Label2 
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
         Height          =   1275
         Left            =   1140
         TabIndex        =   70
         Top             =   2100
         Width           =   6975
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
         Left            =   120
         TabIndex        =   69
         Top             =   1680
         Width           =   9045
      End
   End
End
Attribute VB_Name = "FConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FConfigure
' FILENAME:     C:\My Code\vb\AddIns\CodeHelperFree\CodeHelperConfig.frm
' AUTHOR:       Phil Fresle
' CREATED:      01-Dec-1999
' COPYRIGHT:    Copyright 1999-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Add setting to registry for CodeHelper Add-In
'
' MODIFICATION HISTORY:
' 1.0.0     01-Dec-1999
'           Phil Fresle
'           Initial Version
' 1.0.1     10-Mar-2000
'           Phil Fresle
'           Free Version
' 1.0.2     02-May-2000
'           Phil Fresle
'           Added code to browse for folders and so a click on the URL will
'           jump to my home page and a click on the email will prepare an
'           email to send to me.
' 2.0.0     10-May-2000
'           Phil Fresle
'           Added a date format and changed the labels. Changed to an ActiveX
'           exe so it can be called from the Add-In.
' 2.0.1     21-Jan-2001
'           Phil Fresle
'           Commercial version that includes registration
' 2.0.3     30-Apr-2001
'           Phil Fresle
'           Added shorcuts configuration option to allow user to specify
'           whether shortcuts should be enabled on the menu.
' 6.0       16-Jan-2019
'           Phil Fresle
'           Open source version
'*******************************************************************************
Option Explicit

Private Const MODULE_NAME As String = "FConfigure"

Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Declare Function DLLSelfRegister Lib "vb6stkit.dll" _
    (ByVal lpDllName As String) As Integer

Private m_bOK As Boolean

Public Event Unloading(ByVal bOK As Boolean)

'*******************************************************************************
' Configure (SUB)
'*******************************************************************************
Public Sub Configure(ByVal sLicenseKey As String, _
                     ByVal sLicensedTo As String)
    Dim sType           As String
    Dim sTypeName       As String
    Dim sVersion        As String
    Dim sCaption        As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "Configure"
    
    On Error GoTo ERROR_HANDLER
    
    sTypeName = ""
    sVersion = ""

    sVersion = App.Major & "." & App.Minor & "." & App.Revision
    

    lblVersion.Caption = "Version " & sVersion
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
' cmdBoilerplateBrowser_Click (SUB)
'*******************************************************************************
Private Sub cmdBoilerplateBrowser_Click()
    Dim sPath           As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "cmdBoilerplateBrowser_Click"
    
    On Error GoTo ERROR_HANDLER
    
    sPath = BrowseForFolder(Me.hwnd, "Select path to boilerplates", txtBoilerplatePath.Text)
    If sPath <> "" Then
        txtBoilerplatePath.Text = sPath
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
' cmdCancel_Click (SUB)
'*******************************************************************************
Private Sub cmdCancel_Click()
    On Error Resume Next
    
    m_bOK = False
    Unload Me
End Sub

'*******************************************************************************
' cmdHelp_Click (SUB)
'*******************************************************************************
Private Sub cmdHelp_Click()
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "cmdHelp_Click"
    
    On Error GoTo ERROR_HANDLER
    
    If App.HelpFile <> "" Then
        WinHelp Me.hwnd, App.HelpFile, HELP_FINDER, ""
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
' cmdOK_Click (SUB)
'*******************************************************************************
Private Sub cmdOK_Click()
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "cmdOK_Click"
    
    On Error GoTo ERROR_HANDLER
    
    RegisterAddIn
    
    m_bOK = True
    Unload Me
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
' RegisterAddIn (SUB)
'*******************************************************************************
Private Sub RegisterAddIn()
    Dim lReturn         As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "RegisterAddIn"
    
    On Error GoTo ERROR_HANDLER
    
    ' Save the stuff the user entered in the registry
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TEMPLATES, Trim(txtTemplatePath.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BOILERPLATES, Trim(txtBoilerplatePath.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_USER_TOKENS, Trim(txtUserTokens.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_AUTHOR, Trim(txtAuthor.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_COMPANY, Trim(txtCompany.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_INITIALS, Trim(txtInitials.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TIMEFORMAT, Trim(txtTimeFormat.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_DATEFORMAT, Trim(txtDateFormat.Text)
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_PARAMFORMAT, CStr(cboParamFormat.ListIndex)
    If optNormal(0).Value Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_NORMAL_ERRORS, "1"
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_NORMAL_ERRORS, "2"
    End If
    If optEvent(0).Value Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_EVENT_ERRORS, "1"
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_EVENT_ERRORS, "2"
    End If
    If optProperty(0).Value Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_PROPERTY_ERRORS, "1"
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_PROPERTY_ERRORS, "2"
    End If
    If optBefore(0).Value Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_DOCBEFORE, "1"
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_DOCBEFORE, "0"
    End If
    If optCloseActive(0).Value Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_CLOSEACTIVE, "1"
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_CLOSEACTIVE, "0"
    End If
    
    If Val(txtIndent.Text) < 1 Or Val(txtIndent.Text) > 8 Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_INDENT, DEF_REG_INDENT
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_INDENT, CStr(CLng(Val(txtIndent.Text)))
    End If
    
    If Val(txtWhite.Text) < 1 Or Val(txtWhite.Text) > 8 Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BLANK_LINES, DEF_REG_BLANK_LINES
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BLANK_LINES, CStr(CLng(Val(txtWhite.Text)))
    End If
    
    If optIndentDim(0).Value Then
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_INDENT_DIM, "1"
    Else
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_INDENT_DIM, "2"
    End If
    
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TOOLBARSHOW, chkToolbar.Value
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHOW, CLng(lstButtons.Selected(0))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHOW, CLng(lstButtons.Selected(1))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHOW, CLng(lstButtons.Selected(2))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHOW, CLng(lstButtons.Selected(3))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHOW, CLng(lstButtons.Selected(4))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHOW, CLng(lstButtons.Selected(5))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHOW, CLng(lstButtons.Selected(6))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHOW, CLng(lstButtons.Selected(7))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHOW, CLng(lstButtons.Selected(8))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHOW, CLng(lstButtons.Selected(9))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHOW, CLng(lstButtons.Selected(10))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHOW, CLng(lstButtons.Selected(11))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHOW, CLng(lstButtons.Selected(12))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHOW, CLng(lstButtons.Selected(13))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHOW, CLng(lstButtons.Selected(14))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHOW, CLng(lstButtons.Selected(15))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHOW, CLng(lstButtons.Selected(16))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHOW, CLng(lstButtons.Selected(17))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHOW, CLng(lstButtons.Selected(18))
    
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_SHORTCUTS, chkShortcuts.Value
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHORTCUT, CLng(lstShortcuts.Selected(0))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHORTCUT, CLng(lstShortcuts.Selected(1))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHORTCUT, CLng(lstShortcuts.Selected(2))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHORTCUT, CLng(lstShortcuts.Selected(3))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHORTCUT, CLng(lstShortcuts.Selected(4))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHORTCUT, CLng(lstShortcuts.Selected(5))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHORTCUT, CLng(lstShortcuts.Selected(6))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHORTCUT, CLng(lstShortcuts.Selected(7))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHORTCUT, CLng(lstShortcuts.Selected(8))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHORTCUT, CLng(lstShortcuts.Selected(9))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHORTCUT, CLng(lstShortcuts.Selected(10))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHORTCUT, CLng(lstShortcuts.Selected(11))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHORTCUT, CLng(lstShortcuts.Selected(12))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHORTCUT, CLng(lstShortcuts.Selected(13))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHORTCUT, CLng(lstShortcuts.Selected(14))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHORTCUT, CLng(lstShortcuts.Selected(15))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHORTCUT, CLng(lstShortcuts.Selected(16))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHORTCUT, CLng(lstShortcuts.Selected(17))
    SaveSetting REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHORTCUT, CLng(lstShortcuts.Selected(18))
    
    If App.StartMode = vbSModeStandalone Then
        ' Register as an add-in with VB
        lReturn = WritePrivateProfileString("Add-Ins32", "VBCodeHelper.CCodeHelper", "3", "VBADDIN.INI")
        If lReturn <> 0 Then
            MsgBox "VBCodeHelper Add-In successfully configured", vbInformation, App.ProductName
        Else
            MsgBox "Failed to configure VBCodeHelper Add-In", vbCritical, App.ProductName
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
' cmdTemplateBrowse_Click (SUB)
'*******************************************************************************
Private Sub cmdTemplateBrowse_Click()
    Dim sPath As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "cmdTemplateBrowse_Click"
    
    On Error GoTo ERROR_HANDLER
    
    sPath = BrowseForFolder(Me.hwnd, "Select path to code templates", txtTemplatePath.Text)
    If sPath <> "" Then
        txtTemplatePath.Text = sPath
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
' cmdUserTokens_Click (SUB)
'*******************************************************************************
Private Sub cmdUserTokens_Click()
    Dim sPath           As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "cmdUserTokens_Click"
    
    On Error GoTo ERROR_HANDLER
    
    sPath = BrowseForFolder(Me.hwnd, "Select path to USER_TOKENS.dat", txtUserTokens.Text)
    If sPath <> "" Then
        txtUserTokens.Text = sPath
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
' x (FUNCTION)
'
' PARAMETERS:
' (In/Out) - y - Integer -
'
' RETURN VALUE:
' Variant -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Private Function x(y As Integer)

End Function

'*******************************************************************************
' Form_Load (SUB)
'*******************************************************************************
Private Sub Form_Load()
    Dim oCodeHelper     As Object
    Dim sDLL            As String
    Dim iReturn         As Integer
    Dim sAuthor         As String
    Dim sCompany        As String
    Dim sInitials       As String
    Dim sTemplatesPath  As String
    Dim sHelpfile       As String
    Dim sParamFormat    As String
    
    On Error Resume Next
    
    lblCopyright.Caption = App.LegalCopyright
    
    sTemplatesPath = App.Path
    If Right(sTemplatesPath, 1) = "\" Then
        sTemplatesPath = sTemplatesPath & "Templates"
    Else
        sTemplatesPath = sTemplatesPath & "\Templates"
    End If
    
    If Right(Trim(App.Path), 1) = "\" Then
        sHelpfile = App.Path & HELP_FILE
    Else
        sHelpfile = App.Path & "\" & HELP_FILE
    End If
    
    If Dir(sHelpfile) <> "" Then
        App.HelpFile = sHelpfile
    Else
        App.HelpFile = ""
    End If
    
    fraConfig(0).Visible = True
    fraConfig(1).Visible = False
    fraConfig(2).Visible = False
    fraConfig(3).Visible = False
    fraConfig(4).Visible = False
    fraConfig(5).Visible = False
    optConfig(0).Value = True

    ' Put default values in the form
    'GetOwnerAndCompany sAuthor, sCompany
    txtTemplatePath.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TEMPLATES, sTemplatesPath))
    txtBoilerplatePath.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BOILERPLATES, sTemplatesPath))
    txtUserTokens.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_USER_TOKENS, sTemplatesPath))
    txtAuthor.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_AUTHOR, sAuthor))
    txtCompany.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_COMPANY, sCompany))
    sInitials = GetInitials(txtAuthor.Text)
    txtInitials.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_INITIALS, sInitials))
    txtTimeFormat.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TIMEFORMAT, DEF_REG_TIMEFORMAT))
    txtDateFormat.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_DATEFORMAT, DEF_REG_DATEFORMAT))
    sParamFormat = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_PARAMFORMAT, DEF_REG_PARAMFORMAT))
    
    If Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_NORMAL_ERRORS, DEF_REG_NORMAL_ERRORS)) = "2" Then
        optNormal(1).Value = True
    Else
        optNormal(0).Value = True
    End If
    If Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_EVENT_ERRORS, DEF_REG_EVENT_ERRORS)) = "1" Then
        optEvent(0).Value = True
    Else
        optEvent(1).Value = True
    End If
    If Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_PROPERTY_ERRORS, DEF_REG_PROPERTY_ERRORS)) = "2" Then
        optProperty(1).Value = True
    Else
        optProperty(0).Value = True
    End If
    If Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_DOCBEFORE, DEF_REG_DOCBEFORE)) = "0" Then
        optBefore(1).Value = True
    Else
        optBefore(0).Value = True
    End If
    If Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_CLOSEACTIVE, DEF_REG_CLOSEACTIVE)) = "0" Then
        optCloseActive(1).Value = True
    Else
        optCloseActive(0).Value = True
    End If
    
    If Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_INDENT_DIM, DEF_REG_INDENT_DIM)) = "2" Then
        optIndentDim(1).Value = True
    Else
        optIndentDim(0).Value = True
    End If
    
    txtIndent.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_INDENT, DEF_REG_INDENT))
            
    txtWhite.Text = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BLANK_LINES, DEF_REG_BLANK_LINES))
            
    chkToolbar.Value = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TOOLBARSHOW, vbChecked)
    
    lstButtons.Clear
    lstButtons.AddItem CAPTION_MODULE
    lstButtons.AddItem CAPTION_PROCEDURE
    lstButtons.AddItem CAPTION_TIMESTAMP
    lstButtons.AddItem CAPTION_TEMPLATE
    lstButtons.AddItem CAPTION_CLOSE
    lstButtons.AddItem CAPTION_CLEAR
    lstButtons.AddItem CAPTION_ERROR1
    lstButtons.AddItem CAPTION_ERROR2
    lstButtons.AddItem CAPTION_ERROR_ALL
    lstButtons.AddItem CAPTION_DIM
    lstButtons.AddItem CAPTION_DIM_ALL
    lstButtons.AddItem CAPTION_INDENT
    lstButtons.AddItem CAPTION_INDENT_ALL
    lstButtons.AddItem CAPTION_TAB
    lstButtons.AddItem CAPTION_ZORDER
    lstButtons.AddItem CAPTION_STATS
    lstButtons.AddItem CAPTION_WHITE_SPACE
    lstButtons.AddItem CAPTION_WHITE_SPACE_ALL
    lstButtons.AddItem CAPTION_PROCEDURE_LIST
        
    lstButtons.Selected(0) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHOW, True))
    lstButtons.Selected(1) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHOW, True))
    lstButtons.Selected(2) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHOW, True))
    lstButtons.Selected(3) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHOW, True))
    lstButtons.Selected(4) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHOW, True))
    lstButtons.Selected(5) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHOW, True))
    lstButtons.Selected(6) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHOW, True))
    lstButtons.Selected(7) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHOW, True))
    lstButtons.Selected(8) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHOW, True))
    lstButtons.Selected(9) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHOW, True))
    lstButtons.Selected(10) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHOW, True))
    lstButtons.Selected(11) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHOW, True))
    lstButtons.Selected(12) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHOW, True))
    lstButtons.Selected(13) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHOW, True))
    lstButtons.Selected(14) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHOW, True))
    lstButtons.Selected(15) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHOW, True))
    lstButtons.Selected(16) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHOW, True))
    lstButtons.Selected(17) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHOW, True))
    lstButtons.Selected(18) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHOW, True))
    lstButtons.ListIndex = -1

    chkShortcuts.Value = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_SHORTCUTS, vbChecked)
    
    lstShortcuts.Clear
    lstShortcuts.AddItem SHORTCUT_MODULE
    lstShortcuts.AddItem SHORTCUT_PROCEDURE
    lstShortcuts.AddItem SHORTCUT_TIMESTAMP
    lstShortcuts.AddItem SHORTCUT_TEMPLATE
    lstShortcuts.AddItem SHORTCUT_CLOSE
    lstShortcuts.AddItem SHORTCUT_CLEAR
    lstShortcuts.AddItem SHORTCUT_ERROR1
    lstShortcuts.AddItem SHORTCUT_ERROR2
    lstShortcuts.AddItem SHORTCUT_ERROR_ALL
    lstShortcuts.AddItem SHORTCUT_DIM
    lstShortcuts.AddItem SHORTCUT_DIM_ALL
    lstShortcuts.AddItem SHORTCUT_INDENT
    lstShortcuts.AddItem SHORTCUT_INDENT_ALL
    lstShortcuts.AddItem SHORTCUT_TAB
    lstShortcuts.AddItem SHORTCUT_ZORDER
    lstShortcuts.AddItem SHORTCUT_STATS
    lstShortcuts.AddItem SHORTCUT_WHITE_SPACE
    lstShortcuts.AddItem SHORTCUT_WHITE_SPACE_ALL
    lstShortcuts.AddItem SHORTCUT_PROCEDURE_LIST
        
    lstShortcuts.Selected(0) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHORTCUT, True))
    lstShortcuts.Selected(1) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHORTCUT, True))
    lstShortcuts.Selected(2) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHORTCUT, True))
    lstShortcuts.Selected(3) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHORTCUT, True))
    lstShortcuts.Selected(4) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHORTCUT, True))
    lstShortcuts.Selected(5) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHORTCUT, True))
    lstShortcuts.Selected(6) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHORTCUT, True))
    lstShortcuts.Selected(7) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHORTCUT, True))
    lstShortcuts.Selected(8) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHORTCUT, True))
    lstShortcuts.Selected(9) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHORTCUT, True))
    lstShortcuts.Selected(10) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHORTCUT, True))
    lstShortcuts.Selected(11) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHORTCUT, True))
    lstShortcuts.Selected(12) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHORTCUT, True))
    lstShortcuts.Selected(13) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHORTCUT, True))
    lstShortcuts.Selected(14) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHORTCUT, True))
    lstShortcuts.Selected(15) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHORTCUT, True))
    lstShortcuts.Selected(16) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHORTCUT, True))
    lstShortcuts.Selected(17) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHORTCUT, True))
    lstShortcuts.Selected(18) = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHORTCUT, True))
    lstShortcuts.ListIndex = -1

    With cboParamFormat
        .Clear
        .AddItem PARAM_FORMAT_0
        .AddItem PARAM_FORMAT_1
    End With
    
    Debug.Print sParamFormat
    
    If sParamFormat = "1" Then
        cboParamFormat.ListIndex = 1
    Else
        cboParamFormat.ListIndex = 0
    End If
End Sub

'*******************************************************************************
' Form_Unload (SUB)
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Me.Hide
    RaiseEvent Unloading(m_bOK)
End Sub

'*******************************************************************************
' lblHyperlink_Click (SUB)
'*******************************************************************************
Private Sub lblHyperlink_Click()
    On Error Resume Next
    
    'ShellExecute Me.hwnd, "open", lblHyperlink.Caption, vbNullString, _
    '    "C:\", SW_SHOWNORMAL
End Sub

'*******************************************************************************
' optConfig_Click (SUB)
'*******************************************************************************
Private Sub optConfig_Click(Index As Integer)
    Dim lCount          As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "optConfig_Click"
    
    On Error GoTo ERROR_HANDLER
    
    fraConfig(Index).Visible = True
    
    For lCount = fraConfig.LBound To fraConfig.UBound
        If lCount <> Index Then
            fraConfig(lCount).Visible = False
        End If
    Next
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
' txtIndent_Validate (SUB)
'*******************************************************************************
Private Sub txtIndent_Validate(Cancel As Boolean)
    Dim dIndent         As Double
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "txtIndent_Validate"
    
    On Error GoTo ERROR_HANDLER
    
    If Not IsNumeric(txtIndent.Text) Then
        Cancel = True
        MsgBox "The indent value must be numeric between 1 and 8", vbCritical, App.ProductName
    Else
        dIndent = Val(txtIndent.Text)
        If dIndent < 1 Or dIndent > 8 Then
            Cancel = True
            MsgBox "The indent value must be numeric between 1 and 8", vbCritical, App.ProductName
        End If
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
' txtWhite_Validate (SUB)
'*******************************************************************************
Private Sub txtWhite_Validate(Cancel As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim dWhite          As Double
    
    Const PROCEDURE_NAME As String = "txtWhite_Validate"
    
    On Error GoTo ERROR_HANDLER
    
    If Not IsNumeric(txtWhite.Text) Then
        Cancel = True
        MsgBox "The white space value must be numeric between 1 and 8", vbCritical, App.ProductName
    Else
        dWhite = Val(txtWhite.Text)
        If dWhite < 1 Or dWhite > 8 Then
            Cancel = True
            MsgBox "The white space value must be numeric between 1 and 8", vbCritical, App.ProductName
        End If
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
