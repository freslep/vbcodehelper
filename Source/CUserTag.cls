VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       CUserTag
' FILENAME:     C:\My Code\vb\vbch\Source\CUserTag.cls
' AUTHOR:       Phil Fresle
' CREATED:      17-Jan-2001
' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' User tag
'
' MODIFICATION HISTORY:
' 1.0       21-Jan-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private m_sTagName As String
Private m_sTagPrompt As String
Private m_sTagValue As String

'*******************************************************************************
' TagName (PROPERTY GET)
'*******************************************************************************
Public Property Get TagName() As String
    TagName = m_sTagName
End Property

'*******************************************************************************
' TagName (PROPERTY LET)
'*******************************************************************************
Public Property Let TagName(ByVal NewValue As String)
    m_sTagName = NewValue
End Property

'*******************************************************************************
' TagPrompt (PROPERTY GET)
'*******************************************************************************
Public Property Get TagPrompt() As String
    TagPrompt = m_sTagPrompt
End Property

'*******************************************************************************
' TagPrompt (PROPERTY LET)
'*******************************************************************************
Public Property Let TagPrompt(ByVal NewValue As String)
    m_sTagPrompt = NewValue
End Property

'*******************************************************************************
' TagValue (PROPERTY GET)
'*******************************************************************************
Public Property Get TagValue() As String
    TagValue = m_sTagValue
End Property

'*******************************************************************************
' TagValue (PROPERTY LET)
'*******************************************************************************
Public Property Let TagValue(ByVal NewValue As String)
    m_sTagValue = NewValue
End Property
