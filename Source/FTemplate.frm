VERSION 5.00
Begin VB.Form FTemplate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Template"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FTemplate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2100
      TabIndex        =   3
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   3420
      TabIndex        =   2
      Top             =   3060
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "Select the template to insert:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3735
   End
End
Attribute VB_Name = "FTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FTemplate
' FILENAME:     C:\My Code\vb\vbch\Source\FTemplate.frm
' AUTHOR:       Phil Fresle
' CREATED:      12-Mar-2001
' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' Pick a template, any template
'
' MODIFICATION HISTORY:
' 1.0       13-Mar-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private m_sFile         As String
Private m_sTemplatePath As String

'*******************************************************************************
' cmdCancel_Click (SUB)
'*******************************************************************************
Private Sub cmdCancel_Click()
    On Error Resume Next
    
    m_sFile = ""
    Unload Me
End Sub

'*******************************************************************************
' cmdOK_Click (SUB)
'*******************************************************************************
Private Sub cmdOK_Click()
    On Error Resume Next
    
    If Right(m_sTemplatePath, 1) = "\" Then
        m_sFile = m_sTemplatePath & File1.FileName
    Else
        m_sFile = m_sTemplatePath & "\" & File1.FileName
    End If
    Unload Me
End Sub

'*******************************************************************************
' File1_Click (SUB)
'*******************************************************************************
Private Sub File1_Click()
    cmdOK.Enabled = True
End Sub

'*******************************************************************************
' Form_Load (SUB)
'*******************************************************************************
Private Sub Form_Load()
    Dim sTemplatesPath  As String
    
    On Error Resume Next
    
    File1.Pattern = "*.tlt"
    
    sTemplatesPath = App.Path
    If Right(sTemplatesPath, 1) = "\" Then
        sTemplatesPath = sTemplatesPath & "Templates"
    Else
        sTemplatesPath = sTemplatesPath & "\Templates"
    End If
    
    ' Get settings from registry
    m_sTemplatePath = Trim(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TEMPLATES, sTemplatesPath))
    
    File1.Path = m_sTemplatePath
End Sub

'*******************************************************************************
' FileName (PROPERTY GET)
'*******************************************************************************
Public Property Get FileName() As String
    FileName = m_sFile
End Property
