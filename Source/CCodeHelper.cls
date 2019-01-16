VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCodeHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Add-In VBCodeHelper"
'*******************************************************************************
' MODULE:       CCodeHelper
' FILENAME:     C:\CodeHelperFree\Source\CCodeHelper.cls
' AUTHOR:       Phil Fresle
' CREATED:      01-Dec-1999
' COPYRIGHT:    Copyright 1999-2019 Frez Systems Limited.
'
' DESCRIPTION:
' Add-In to document the VB code in the active window
'
' MODIFICATION HISTORY:
' 1.0.0     01-Dec-1999
'           Phil Fresle
'           Initial Version
' 1.0.1     10-Mar-2000
'           Phil Fresle
'           Free Version
' 1.0.2     12-Mar-2000
'           Phil Fresle
'           Fixed bug with Let/Set/Get where logic reversed.
' 1.0.3     22-Mar-2000
'           Phil Fresle
'           Fixed bug with OnConnection where initials were defaulting to author
'           rather than the initials worked out from the author's name
' 1.0.4     08-Apr-2000
'           Phil Fresle
'           Fixed bug with OnConnection where path rather than TemplatePath and
'           BoilerplatePath was trying to be obtained.  Thanks to Tore Bostrup
'           for finding this error.  Also set toolbar destination to a constant
'           to allow it to be changed quickly.
' 1.0.5     14-Apr-2000
'           Phil Fresle
'           Added two more buttons, one to close all code and designer windows
'           and the other to clear the immediate window
' 1.0.6     28-Apr-2000
'           Phil Fresle
'           Changed the time stamp code so if user selected a bunch of lines it
'           puts a START CHANGE and END CHANGE comment around them.
' 1.0.7     30-Apr-2000
'           Phil Fresle
'           Following an idea received by email from Eric Eggermann, changed the
'           time stamp code to use templates for flexibility
' 1.0.8     02-May-2000
'           Phil Fresle
'           The Add-In now uses its own toolbar and remembers its position
'           between invocations
' 2.0.0     10-May-2000
'           Phil Fresle
'           Adds a menu option on the Add-In menu so we can configure ourselves,
'           uses the new date format supplied by the config program, reads in
'           data from USER_TAGS.dat to use for prompts when inserting templates.
' 2.0.1     17-Jan-2001
'           Phil Fresle
'           Commercial version.
'           This version includes new error handling insertion buttons,
'           and rejuvinated code.
' 4.0.3     Sep-2001
'           Phil Fresle
'           Tries to prevent errors loading images where user is using Terminal
'           Services by introducing a delay and trying several time to paste
'           from the clipboard
' 4.0.4     Oct-2001
'           Phil Fresle
'           No longer a dependance on the Microsoft Scripting Runtime that was
'           giving some users errors necessitaing a reinstall
' 4.0.5     Oct-2001
'           Phil Fresle
'           Timestamp in the middle of a line with continuation characters now
'           inserts the timestamp before and after the line, not in the middle
' 4.1.0     10-Oct-2001
'           Phil Fresle
'           This version now has buttons with transparent backgrounds and
'           tooltips. Can now also configure so that procedure declarations are
'           not indented during the smart indent process.
' 4.1.1     15-Oct-2001
'           Phil Fresle
'           Added 'Tidy White Space' functionality and a stub for 'Procedure List'
' 5.0.0     30-Oct-2001
'           Phil Fresle
'           Added procedure list functionality. Auto tab now also auto tabs the
'           declaration area.
' 5.4.3     Code to work in German (foreign languages) - do not rely on name of
'           MenuBar in collection
' 5.4.4     Different way of masking the ocons to make them transparent
' 6.0       16-Jan-2019
'           Phil Fresle
'           Open source version.
'*******************************************************************************
Option Explicit

Implements IDTExtensibility

Private Const MODULE_NAME               As String = "CCodeHelper"

Private WithEvents m_evtProjects        As VBProjectsEvents
Attribute m_evtProjects.VB_VarHelpID = -1
Private WithEvents m_evtComponents      As VBComponentsEvents
Attribute m_evtComponents.VB_VarHelpID = -1
Private WithEvents m_evtControls        As VBControlsEvents
Attribute m_evtControls.VB_VarHelpID = -1

Private WithEvents m_evtButtonClose     As CommandBarEvents
Attribute m_evtButtonClose.VB_VarHelpID = -1
Private WithEvents m_evtButtonClear     As CommandBarEvents
Attribute m_evtButtonClear.VB_VarHelpID = -1
Private WithEvents m_evtButtonModule    As CommandBarEvents
Attribute m_evtButtonModule.VB_VarHelpID = -1
Private WithEvents m_evtButtonProcedure As CommandBarEvents
Attribute m_evtButtonProcedure.VB_VarHelpID = -1
Private WithEvents m_evtButtonTimeStamp As CommandBarEvents
Attribute m_evtButtonTimeStamp.VB_VarHelpID = -1
Private WithEvents m_evtComboTemplate   As CommandBarEvents
Attribute m_evtComboTemplate.VB_VarHelpID = -1
Private WithEvents m_evtButtonError1    As CommandBarEvents
Attribute m_evtButtonError1.VB_VarHelpID = -1
Private WithEvents m_evtButtonError2    As CommandBarEvents
Attribute m_evtButtonError2.VB_VarHelpID = -1
Private WithEvents m_evtConfigure       As CommandBarEvents
Attribute m_evtConfigure.VB_VarHelpID = -1
Private WithEvents m_evtDim             As CommandBarEvents
Attribute m_evtDim.VB_VarHelpID = -1
Private WithEvents m_evtDimAll          As CommandBarEvents
Attribute m_evtDimAll.VB_VarHelpID = -1
Private WithEvents m_evtErrorAll        As CommandBarEvents
Attribute m_evtErrorAll.VB_VarHelpID = -1
Private WithEvents m_evtIndent          As CommandBarEvents
Attribute m_evtIndent.VB_VarHelpID = -1
Private WithEvents m_evtIndentAll       As CommandBarEvents
Attribute m_evtIndentAll.VB_VarHelpID = -1
Private WithEvents m_evtTab             As CommandBarEvents
Attribute m_evtTab.VB_VarHelpID = -1
Private WithEvents m_evtZOrder          As CommandBarEvents
Attribute m_evtZOrder.VB_VarHelpID = -1
Private WithEvents m_evtStats           As CommandBarEvents
Attribute m_evtStats.VB_VarHelpID = -1
Private WithEvents m_evtWhiteSpace      As CommandBarEvents
Attribute m_evtWhiteSpace.VB_VarHelpID = -1
Private WithEvents m_evtWhiteSpaceAll   As CommandBarEvents
Attribute m_evtWhiteSpaceAll.VB_VarHelpID = -1
Private WithEvents m_evtProcedureList   As CommandBarEvents
Attribute m_evtProcedureList.VB_VarHelpID = -1

Private WithEvents m_evtMenuClose           As CommandBarEvents
Attribute m_evtMenuClose.VB_VarHelpID = -1
Private WithEvents m_evtMenuClear           As CommandBarEvents
Attribute m_evtMenuClear.VB_VarHelpID = -1
Private WithEvents m_evtMenuModule          As CommandBarEvents
Attribute m_evtMenuModule.VB_VarHelpID = -1
Private WithEvents m_evtMenuProcedure       As CommandBarEvents
Attribute m_evtMenuProcedure.VB_VarHelpID = -1
Private WithEvents m_evtMenuTimeStamp       As CommandBarEvents
Attribute m_evtMenuTimeStamp.VB_VarHelpID = -1
Private WithEvents m_evtMenuTemplate        As CommandBarEvents
Attribute m_evtMenuTemplate.VB_VarHelpID = -1
Private WithEvents m_evtMenuError1          As CommandBarEvents
Attribute m_evtMenuError1.VB_VarHelpID = -1
Private WithEvents m_evtMenuError2          As CommandBarEvents
Attribute m_evtMenuError2.VB_VarHelpID = -1
Private WithEvents m_evtMenuAbout           As CommandBarEvents
Attribute m_evtMenuAbout.VB_VarHelpID = -1
Private WithEvents m_evtMenuHelp            As CommandBarEvents
Attribute m_evtMenuHelp.VB_VarHelpID = -1
Private WithEvents m_evtMenuDim             As CommandBarEvents
Attribute m_evtMenuDim.VB_VarHelpID = -1
Private WithEvents m_evtMenuDimAll          As CommandBarEvents
Attribute m_evtMenuDimAll.VB_VarHelpID = -1
Private WithEvents m_evtMenuErrorAll        As CommandBarEvents
Attribute m_evtMenuErrorAll.VB_VarHelpID = -1
Private WithEvents m_evtMenuIndent          As CommandBarEvents
Attribute m_evtMenuIndent.VB_VarHelpID = -1
Private WithEvents m_evtMenuIndentAll       As CommandBarEvents
Attribute m_evtMenuIndentAll.VB_VarHelpID = -1
Private WithEvents m_evtMenuTab             As CommandBarEvents
Attribute m_evtMenuTab.VB_VarHelpID = -1
Private WithEvents m_evtMenuZOrder          As CommandBarEvents
Attribute m_evtMenuZOrder.VB_VarHelpID = -1
Private WithEvents m_evtMenuStats           As CommandBarEvents
Attribute m_evtMenuStats.VB_VarHelpID = -1
Private WithEvents m_evtMenuWhiteSpace      As CommandBarEvents
Attribute m_evtMenuWhiteSpace.VB_VarHelpID = -1
Private WithEvents m_evtMenuWhiteSpaceAll   As CommandBarEvents
Attribute m_evtMenuWhiteSpaceAll.VB_VarHelpID = -1
Private WithEvents m_evtMenuProcedureList   As CommandBarEvents
Attribute m_evtMenuProcedureList.VB_VarHelpID = -1

Private m_cmdClose                      As Office.CommandBarButton
Private m_cmdClear                      As Office.CommandBarButton
Private m_cmdModule                     As Office.CommandBarButton
Private m_cmdProcedure                  As Office.CommandBarButton
Private m_cmdTimeStamp                  As Office.CommandBarButton
Private m_cboTemplate                   As Office.CommandBarComboBox
Private m_cmdError1                     As Office.CommandBarButton
Private m_cmdError2                     As Office.CommandBarButton
Private m_cmdDim                        As Office.CommandBarButton
Private m_cmdDimAll                     As Office.CommandBarButton
Private m_cmdErrorAll                   As Office.CommandBarButton
Private m_cmdIndent                     As Office.CommandBarButton
Private m_cmdIndentAll                  As Office.CommandBarButton
Private m_cmdTab                        As Office.CommandBarButton
Private m_cmdZOrder                     As Office.CommandBarButton
Private m_cmdStats                      As Office.CommandBarButton
Private m_cmdWhiteSpace                 As Office.CommandBarButton
Private m_cmdWhiteSpaceAll              As Office.CommandBarButton
Private m_cmdProcedureList              As Office.CommandBarButton

Private m_mnuClose                      As Office.CommandBarButton
Private m_mnuClear                      As Office.CommandBarButton
Private m_mnuModule                     As Office.CommandBarButton
Private m_mnuProcedure                  As Office.CommandBarButton
Private m_mnuTimeStamp                  As Office.CommandBarButton
Private m_mnuTemplate                   As Office.CommandBarButton
Private m_mnuError1                     As Office.CommandBarButton
Private m_mnuError2                     As Office.CommandBarButton
Private m_mnuAbout                      As Office.CommandBarButton
Private m_mnuHelp                       As Office.CommandBarButton
Private m_mnuDim                        As Office.CommandBarButton
Private m_mnuDimAll                     As Office.CommandBarButton
Private m_mnuErrorAll                   As Office.CommandBarButton
Private m_mnuIndent                     As Office.CommandBarButton
Private m_mnuIndentAll                  As Office.CommandBarButton
Private m_mnuTab                        As Office.CommandBarButton
Private m_mnuZOrder                     As Office.CommandBarButton
Private m_mnuStats                      As Office.CommandBarButton
Private m_mnuConfigure                  As Office.CommandBarButton
Private m_mnuWhiteSpace                 As Office.CommandBarButton
Private m_mnuWhiteSpaceAll              As Office.CommandBarButton
Private m_mnuProcedureList              As Office.CommandBarButton

Private m_VBInst                        As VBIDE.VBE

Private m_StatsWindow                   As VBIDE.Window
Private m_TabWindow                     As VBIDE.Window
Private m_ZOrderWindow                  As VBIDE.Window
Private m_TemplateWindow                As VBIDE.Window
Private m_ProcedureListWindow           As VBIDE.Window
Private m_docTab                        As docTabOrder
Private m_docZOrder                     As docZOrder
Private m_docStats                      As docStatistics
Private m_docTemplates                  As docTemplates
Private m_docProcedureList              As docProcedureList

Private m_bStatsLoaded                  As Boolean
Private m_bTabLoaded                    As Boolean
Private m_bZOrderLoaded                 As Boolean
Private m_bTemplatesLoaded              As Boolean
Private m_bProcedureListLoaded          As Boolean

Private m_sParamFormat                  As String
Private m_sZOrderDisplay                As String
Private m_sTabDisplay                   As String
Private m_sIndent                       As String
Private m_sIndentDim                    As String
Private m_sErrorText                    As String
Private m_sNormalErrors                 As String
Private m_sEventErrors                  As String
Private m_sPropertyErrors               As String
Private m_sCompany                      As String
Private m_sAuthor                       As String
Private m_sInitials                     As String
Private m_sBoilerplatePath              As String
Private m_sTemplatePath                 As String
Private m_bDocumentBeforeProcedure      As Boolean
Private m_bCloseActive                  As Boolean
Private m_sTimeFormat                   As String
Private m_sDateFormat                   As String
Private m_lBlankLines                   As Long
Private m_colUserTags                   As Collection
Private m_frmAbout                      As FAbout

Private WithEvents m_oConfig        As VBCodeHelperConfig.CEntryPoint
Attribute m_oConfig.VB_VarHelpID = -1
Private WithEvents m_oKeyHandler    As CHotKey
Attribute m_oKeyHandler.VB_VarHelpID = -1


Private Sub GetProcedureNameAndType(ByVal lStartLine As Long, _
                                    sProcedureName As String, _
                                    eptType As enumProcedureType)
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim lProcStartLine  As Long
    
    With m_VBInst.ActiveCodePane.CodeModule
        sProcedureName = .ProcOfLine(lStartLine, vbext_pk_Proc)
        
        eptType = PTUnknown
        
        If Len(sProcedureName) > 0 Then
            Set oMember = .Members(sProcedureName)
            
            If oMember.Type = vbext_mt_Method Then
                lTopLine = .ProcBodyLine(sProcedureName, vbext_pk_Proc)
                eptType = GetMethodType(m_VBInst, lTopLine, sProcedureName)
            
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
        End If
    End With
End Sub

Private Sub PrepareProcedureTree()
    ' Loop through all modules and build a list of procedures, what
    ' their modules is, whether they are public, private or friend,
    ' what the line number of each one is? Sort alphabetically and
    ' add to a treeview
    
    ' Then for each procedure; if private then look in its own module
    ' to see if any other procedures call it, otherwise, look in all
    ' other modules to see if anyone calls it. For each thing that calls it
    ' add a child to its node in the treeview
End Sub

Private Function FoundVariable(ByVal sVariableName As String, _
                               ByVal sCode As String) As Boolean

    Dim sLook1          As String
    Dim sLook2          As String
    Dim sLook3          As String
    Dim sLook4          As String
    Dim sLook5          As String
    Dim sLook6          As String
    Dim sLook7          As String
    Dim sLook8          As String
    Dim sLook9          As String
    Dim sLook10         As String
    Dim sLook11         As String
    Dim sLook12         As String
    Dim sLook13         As String
    Dim sLook14         As String
    
    If sVariableName <> "" Then
        sLook1 = " " & sVariableName & " "
        sLook2 = "(" & sVariableName & ","
        sLook3 = " " & sVariableName & ")"
        sLook4 = " " & sVariableName & ","
        sLook5 = " " & sVariableName & ":"
        sLook6 = "(" & sVariableName & ")"
        sLook7 = " " & sVariableName & "("
        sLook8 = "(" & sVariableName & "("
        sLook9 = " " & sVariableName & "."
        sLook10 = "(" & sVariableName & "."
        sLook11 = ":=" & sVariableName & ","
        sLook12 = ":=" & sVariableName & " "
        sLook13 = ":=" & sVariableName & ")"
        sLook14 = ":=" & sVariableName & ":"
        
        If InStr(sCode, sVariableName) > 0 Then
            If InStr(sCode, sLook1) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook2) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook3) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook4) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook5) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook6) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook7) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook8) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook9) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook10) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook11) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook12) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook13) > 0 Then
                FoundVariable = True
            ElseIf InStr(sCode, sLook14) > 0 Then
                FoundVariable = True
            Else
                FoundVariable = False
            End If
        Else
            FoundVariable = False
        End If
    Else
        FoundVariable = False
    End If
End Function

Private Function GetProcedureVariables(ByVal lLineNumber As Long, _
                                       ByVal oMember As Member, _
                                       ByVal eptType As enumProcedureType, _
                                       sVariables() As String) As Boolean

    Dim sEndTest        As String
    Dim eProcKind       As vbext_ProcKind
    Dim lTop            As Long
    Dim sText           As String
    Dim sProcName       As String
    Dim lLinesCount     As Long
    Dim bFoundEnd       As Boolean
    Dim sTextNext       As String
    Dim lStart          As Long
    Dim lVariables      As Long
    Dim lEnd            As Long
    Dim lVarLine        As Long
    Dim sTemp           As String
    Dim sVars()         As String
    Dim lVarCount       As Long
    Dim sVariableName   As String
    Dim bFoundOne       As Boolean
    
    Select Case eptType
        Case PTFunction
            sEndTest = "End Function"
            eProcKind = vbext_pk_Proc
            
        Case PTGet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Get
        
        Case PTLet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Let
        
        Case PTSet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Set
        
        Case PTSub
            sEndTest = "End Sub"
            eProcKind = vbext_pk_Proc
        
        Case Else
            Exit Function
    End Select

    lTop = lLineNumber
    With m_VBInst.ActiveCodePane.CodeModule
        sProcName = .ProcOfLine(lLineNumber, eProcKind)
        lLinesCount = .CountOfLines
        
        ' Find first line after declaration by checking continuation characters
        bFoundEnd = False
        Do
            sText = CleanUpLine(.Lines(lTop, 1))
            lTop = lTop + 1
            If sText = sEndTest Then
                bFoundEnd = True
            ElseIf sText = "End _" Then
                sTextNext = CleanUpLine(.Lines(lTop, 1))
                If sTextNext = Mid(sEndTest, 5) Then
                    bFoundEnd = True
                End If
            End If
        Loop Until Right(sText, 2) <> " _" Or lTop > lLinesCount Or bFoundEnd
        
        If lTop > lLinesCount And Not bFoundEnd Then
            Exit Function
        End If

        ' Loop through the procedure lines looking for for variable declarations
        lStart = lTop
        bFoundEnd = False
        lVariables = 0
        Do
            lEnd = lStart
            sText = CleanUpLine(.Lines(lStart, 1))
            lVarLine = lStart
            lStart = lStart + 1
            
            If sText = sEndTest Then
                bFoundEnd = True
            ElseIf sText = "End _" Then
                sTextNext = CleanUpLine(.Lines(lStart, 1))
                If sTextNext = Mid(sEndTest, 5) Then
                    bFoundEnd = True
                    Exit Do
                End If
            End If
            
            If Left(sText, 4) = "Dim " Or Left(sText, 7) = "Static " Then
                ' Put together the full declaration line
                If InStr(sText, ":") > 0 Then
                    sTemp = Left(sText, InStr(sText, ":") - 1)
                Else
                    If Right(sText, 2) = " _" Then
                        sTemp = ""
                        Do
                            sTemp = sTemp & Left(sText, Len(sText) - 1)
                            sText = CleanUpLine(.Lines(lStart, 1))
                            lStart = lStart + 1
                        Loop Until Right(sText, 2) <> " _" Or bFoundEnd Or lStart > lLinesCount
                        sTemp = sTemp & sText
                        If InStr(sTemp, ":") > 0 Then
                            sTemp = Left(sTemp, InStr(sTemp, ":") - 1)
                        End If
                    Else
                        sTemp = sText
                    End If
                End If
                
                If Left(sTemp, 4) = "Dim " Then
                    sVars = SplitDeclaration(Mid(sTemp, 5))
                ElseIf Left(sTemp, 7) = "Static " Then
                    sVars = SplitDeclaration(Mid(sTemp, 8))
                End If
            
                For lVarCount = LBound(sVars) To UBound(sVars)
                    ReDim Preserve sVariables(1, lVariables)
                    sTemp = Trim(sVars(lVarCount))
                    If InStr(sTemp, "(") > 0 Then
                        sTemp = Left(sTemp, InStr(sTemp, "(") - 1)
                    ElseIf InStr(sTemp, " ") > 0 Then
                        sTemp = Left(sTemp, InStr(sTemp, " ") - 1)
                    End If
                    sVariables(0, lVariables) = sTemp
                    sVariables(1, lVariables) = lVarLine
                    lVariables = lVariables + 1
                Next
            End If
        Loop Until bFoundEnd Or lStart > lLinesCount
        
        If lVariables = 0 Then
            Exit Function
        End If
        
        ' Loop through lines of code, and for each line of code loop through
        ' array of variables clearing entry in array if it is found in the code
        bFoundOne = False
        lStart = lTop
        bFoundEnd = False
        Do
            sText = " " & CleanUpLine(.Lines(lStart, 1)) & " "
            lStart = lStart + 1
        
            If Left(sText, 5) <> " Dim " And Left(sText, 8) <> " Static " And Not bFoundOne Then
                For lVarCount = LBound(sVariables, 2) To UBound(sVariables, 2)
                    sVariableName = sVariables(0, lVarCount)
                    If InStr(sVariableName, "(") > 0 Then
                        sVariableName = Left(sVariableName, InStr(sVariableName, "(") - 1)
                    End If
                    
                    If FoundVariable(sVariableName, sText) Then
                        sVariables(0, lVarCount) = vbNullString
                    End If
                Next
            Else
                If bFoundOne Then
                    If Right(sText, 3) <> " _ " Then
                        bFoundOne = False
                    End If
                Else
                    If (Left(sText, 5) = " Dim " Or Left(sText, 8) = " Static ") And Right(sText, 3) = " _ " Then
                        bFoundOne = True
                    End If
                End If
            End If
        Loop Until lStart > lEnd Or lStart > lLinesCount
    End With
End Function

Private Function FindUsedVariables(ByVal sCode As String, _
                                   ByVal lStart As Long, _
                                   ByVal lEnd As Long, _
                                   ByVal lSubStart As Long, _
                                   ByVal lSubEnd As Long, _
                                   lFoundOutside As Long, _
                                   lFoundInside As Long, _
                                   sVariables() As String) As String()
    Dim sFoundVariables()   As String
    Dim bFoundOne           As Boolean
    Dim bFoundEnd           As Boolean
    Dim sText               As String
    Dim lVarCount           As Long
    Dim sVariableName       As String
    
    lFoundOutside = 0
    lFoundInside = 0
    
    ReDim sFoundVariables(3, LBound(sVariables, 2) To UBound(sVariables, 2))
    
    For lVarCount = LBound(sVariables, 2) To UBound(sVariables, 2)
        sFoundVariables(0, lVarCount) = sVariables(0, lVarCount)
        sFoundVariables(1, lVarCount) = sVariables(1, lVarCount)
    Next
    
    bFoundOne = False
    bFoundEnd = False
    
    ' Loop through lines of code, and for each line of code loop through
    ' array of variables clearing entry in array if it is found in the code
    Do
        sText = " " & CleanUpLine(m_VBInst.ActiveCodePane.CodeModule.Lines(lStart, 1)) & " "
        lStart = lStart + 1
    
        If Left(sText, 5) <> " Dim " And Left(sText, 8) <> " Static " And Not bFoundOne Then
            For lVarCount = LBound(sFoundVariables, 2) To UBound(sFoundVariables, 2)
                sVariableName = sFoundVariables(0, lVarCount)
                If InStr(sVariableName, "(") > 0 Then
                    sVariableName = Left(sVariableName, InStr(sVariableName, "(") - 1)
                End If
                
                If FoundVariable(sVariableName, sText) Then
                    If (lStart - 1) >= lSubStart And (lStart - 1) <= lSubEnd Then
                        sFoundVariables(3, lVarCount) = "Y"
                        lFoundInside = lFoundInside + 1
                    Else
                        sFoundVariables(2, lVarCount) = "Y"
                        lFoundOutside = lFoundOutside + 1
                    End If
                    lVarCount = lVarCount + 1
'                    sVariables(0, lVarCount) = vbNullString
                End If
            Next
        Else
            If bFoundOne Then
                If Right(sText, 3) <> " _ " Then
                    bFoundOne = False
                End If
            Else
                If (Left(sText, 5) = " Dim " Or Left(sText, 8) = " Static ") And Right(sText, 3) = " _ " Then
                    bFoundOne = True
                End If
            End If
        End If
    Loop Until lStart > lEnd
    
    FindUsedVariables = sFoundVariables
End Function

'*******************************************************************************
' RemoveUnusedVariablesInProcedure (SUB)
'*******************************************************************************
Private Sub RemoveUnusedVariablesInProcedure(ByVal lLineNumber As Long, _
                                             ByVal oMember As Member, _
                                             ByVal eptType As enumProcedureType)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lQuoteCount     As Long
    Dim sVariables()    As String
    Dim lVarLineCount   As Long
    Dim lCount          As Long
    Dim sAfterColon     As String
    Dim sChar           As String
    Dim lLastLine       As Long
    Dim sOldText        As String
    Dim sNewText        As String
    Dim sParts()        As String
    Dim sFirstPart      As String
    Dim lSpaceCount     As Long
    Dim lPartCount      As Long
    Dim lPartsReplaced  As Long
    Dim sMinorPart      As String
    Dim sVarNames()     As String
    Dim sText           As String
    Dim lVarLine        As Long
    Dim sTemp           As String
    Dim bFoundOne       As Boolean
    Dim lVarCount       As Long
    
    Const PROCEDURE_NAME As String = "RemoveUnusedVariablesInProcedure"
    
    On Error GoTo ERROR_HANDLER
    
    If GetProcedureVariables(lLineNumber, oMember, eptType, sVariables) Then
        
        With m_VBInst.ActiveCodePane.CodeModule
        
            ' Now remove unused
            For lVarCount = UBound(sVariables, 2) To LBound(sVariables, 2) Step -1
                If sVariables(0, lVarCount) <> "" Then
                    lVarLine = sVariables(1, lVarCount)
                    lLastLine = lVarLine - 1
                    
                    ' Get the whole line including continues
                    Do
                        lLastLine = lLastLine + 1
                        sText = .Lines(lLastLine, 1)
                        sTemp = CleanUpLine(sText)
                        If lLastLine = lVarLine Then
                            sOldText = sText
                        Else
                            sOldText = sOldText & vbCrLf & sText
                        End If
                    Loop Until InStr(sTemp, ":") > 0 Or Right(sTemp, 2) <> " _"
                    
                    sAfterColon = ""
                    If InStr(sOldText, ":") > 0 Then
                        ' We need to strip after the colon, correct before the colon
                        ' then put the after colon stuff back
                        bFoundOne = False
                        lCount = 0
                        lQuoteCount = 0
                        Do
                            lCount = lCount + 1
                            sChar = Mid(sOldText, lCount, 1)
                            If sChar = """" Then
                                lQuoteCount = lQuoteCount + 1
                            ElseIf sChar = ":" And (lQuoteCount Mod 2 = 0) Then
                                bFoundOne = True
                            End If
                        Loop Until bFoundOne
                        sAfterColon = Mid(sOldText, lCount)
                        sOldText = Left(sOldText, lCount - 1)
                    End If
                    
                    ' Now correct the line
                    lVarLineCount = 0
                    For lCount = LBound(sVariables, 2) To UBound(sVariables, 2)
                        If sVariables(1, lCount) = lVarLine Then
                            lVarLineCount = lVarLineCount + 1
                        End If
                    Next
                    
                    ' Comment existing lines
                    For lCount = lVarLine To lLastLine
                        sTemp = .Lines(lCount, 1)
                        If Left(sTemp, 1) <> "'" Then
                            If lCount = lLastLine And Right(sTemp, 2) = " _" Then
                                sTemp = Left(sTemp, Len(sTemp) - 2)
                            End If
                            If Left(sTemp, 1) = " " Then
                                .ReplaceLine lCount, "'VBCH" & sTemp
                            Else
                                .ReplaceLine lCount, "'VBCH " & sTemp
                            End If
                        End If
                    Next
                    
                    ' Is there just one variable declared on this line?
                    If lVarLineCount = 1 Then
                        ' Was there a colon?
                        If sAfterColon <> "" Then
                            lCount = Len(sTemp) - Len(LTrim(sTemp))
                            .InsertLines lLastLine + 1, _
                                String(lCount, " ") & Trim(Mid(sAfterColon, 2))
                        End If
                    
                    ' More than one variable on this line
                    Else
                        ' Step through all the variables on this line
                        ' removing them if necessary from sOldText
                        lSpaceCount = Len(sOldText) - Len(LTrim(sOldText))
                        If Left(LTrim(sOldText), 4) = "Dim " Then
                            sFirstPart = String(lSpaceCount, " ") & "Dim "
                            sOldText = Mid(LTrim(sOldText), 5)
                        Else
                            sFirstPart = String(lSpaceCount, " ") & "Static "
                            sOldText = Mid(LTrim(sOldText), 8)
                        End If
                        
                        sParts = SplitDeclaration(sOldText)
                        lPartsReplaced = 0
                        sVarNames = GetVarNames(sParts)
                        
                        For lCount = UBound(sVariables, 2) To LBound(sVariables, 2) Step -1
                            If sVariables(1, lCount) = lVarLine Then
                                If sVariables(0, lCount) <> "" Then
                                    For lPartCount = UBound(sParts) To LBound(sParts) Step -1
                                        If sParts(lPartCount) <> "" Then
                                            sMinorPart = " " & Replace(sVarNames(lPartCount), " _" & vbCrLf, "") & " "
                                            If InStr(sMinorPart, " " & sVariables(0, lCount) & " ") > 0 Then
                                                lPartsReplaced = lPartsReplaced + 1
                                                sParts(lPartCount) = ""
                                                sVariables(0, lCount) = ""
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                        
                        ' If all parts were not removed we need to put them back
                        sNewText = ""
                        If (UBound(sParts) - LBound(sParts) + 1) <> lPartsReplaced Then
                            sNewText = sFirstPart
                            For lPartCount = LBound(sParts) To UBound(sParts)
                                If sParts(lPartCount) <> "" Then
                                    sNewText = sNewText & sParts(lPartCount) & ","
                                End If
                            Next
                            sNewText = Left(sNewText, Len(sNewText) - 1)
                        End If
                        
                        ' Was there a colon?
                        If sAfterColon <> "" Then
                            lCount = Len(sTemp) - Len(LTrim(sTemp))
                            If sNewText <> "" Then
                                .InsertLines lLastLine + 1, sNewText & vbCrLf & _
                                    String(lCount, " ") & Trim(Mid(sAfterColon, 2))
                            Else
                                .InsertLines lLastLine + 1, _
                                    String(lCount, " ") & Trim(Mid(sAfterColon, 2))
                            End If
                        Else
                            If sNewText <> "" Then
                                .InsertLines lLastLine + 1, sNewText
                            End If
                        End If
                    End If
                End If
            Next
        End With
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
' CalculateStats (SUB)
'*******************************************************************************
Public Sub CalculateStats()
    Dim oProject    As VBProject
    Dim bDirty      As Boolean
    Dim lReply      As Long
    
    On Error Resume Next
    
    bDirty = False
    For Each oProject In m_VBInst.VBProjects
        If oProject.IsDirty Then
            bDirty = True
            Exit For
        End If
    Next
    
    If bDirty Then
        lReply = MsgBox("It is recommended that you save the project(s) before" & _
            " obtaining project statistics, do you want to continue anyway?", _
            vbQuestion + vbYesNo, App.ProductName)
        If lReply = vbNo Then
            Exit Sub
        End If
    End If
    
    m_StatsWindow.Visible = True
    m_docStats.RefreshList
End Sub

'*******************************************************************************
' CleanUnusedVariables (SUB)
'*******************************************************************************
Public Sub CleanUnusedVariablesProcedure()
    Dim lStartLine      As Long
    Dim lDummy          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "CleanUnusedVariablesProcedure"
    
    On Error GoTo ERROR_HANDLER
    
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "Select a code pane first", _
            vbCritical, App.ProductName
    Else
        ' Find out where the cursor is at the moment
        m_VBInst.ActiveCodePane.GetSelection lStartLine, lDummy, lDummy, lDummy
        
        With m_VBInst.ActiveCodePane.CodeModule
            GetProcedureDetails m_VBInst, lStartLine, oMember, eptType, lTopLine, lEndDecLine
            
            If lTopLine <> 0 Then
                RemoveUnusedVariablesInProcedure lTopLine, oMember, eptType
            Else
                MsgBox "You must be within a procedure in order to remove unused variables", _
                    vbCritical, App.ProductName
            End If
        End With
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
' ClearAllWhiteSpace (SUB)
'*******************************************************************************
Private Sub ClearAllWhiteSpace()
    Dim oProject    As VBProject
    Dim bDirty      As Boolean
    Dim lReply      As Long
    Dim oComponent  As VBComponent
    
    On Error Resume Next
    
    bDirty = False
    For Each oProject In m_VBInst.VBProjects
        If oProject.IsDirty Then
            bDirty = True
            Exit For
        End If
    Next
    
    If bDirty Then
        lReply = MsgBox("It is recommended that you save the project(s) before" & _
            " removing white space in all modules, do you want to continue anyway?", _
            vbQuestion + vbYesNo, App.ProductName)
        If lReply = vbNo Then
            Exit Sub
        End If
    End If
        
    #If DEBUGMODE = 1 Then
        DebugOutput "Starting ClearAllWhiteSpace..."
    #End If
    
    For Each oProject In m_VBInst.VBProjects
        #If DEBUGMODE = 1 Then
            DebugOutput "  Stepping into project " & oProject.Name & "..."
        #End If
    
        For Each oComponent In oProject.VBComponents
            #If DEBUGMODE = 1 Then
                DebugOutput "    Stepping into component " & oComponent.Name & "..."
                DebugOutput "    Component is of type " & oComponent.Type
            #End If
        
            Select Case oComponent.Type
                Case vbext_ct_ClassModule, vbext_ct_DocObject, vbext_ct_PropPage, _
                    vbext_ct_StdModule, vbext_ct_UserControl, vbext_ct_VBForm, _
                    vbext_ct_VBMDIForm, vbext_ct_ActiveXDesigner
                    #If DEBUGMODE = 1 Then
                        DebugOutput "    About to clear space in " & oComponent.Name & "..."
                    #End If
                    If Not oComponent.CodeModule Is Nothing Then
                        ClearComponentWhiteSpace oComponent.CodeModule
                    End If
                Case vbext_ct_MSForm, vbext_ct_ResFile, vbext_ct_RelatedDocument
                    #If DEBUGMODE = 1 Then
                        DebugOutput "    Component is invalid type for clearing white space."
                    #End If
                    ' Do Nothing
                Case Else
                    #If DEBUGMODE = 1 Then
                        DebugOutput "    Component is unknown type for clearing white space."
                    #End If
                    ' Do Nothing
            End Select
        Next
    Next

    #If DEBUGMODE = 1 Then
        DebugOutput "Finished clearing all white space"
    #End If
End Sub

'*******************************************************************************
' ClearComponentWhiteSpace (SUB)
'*******************************************************************************
Private Sub ClearComponentWhiteSpace(oComponent As CodeModule)
    Dim lLines          As Long
    Dim lLineCount      As Long
    Dim lBlankCount     As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "ClearComponentWhiteSpace"
    
    On Error GoTo ERROR_HANDLER
    
    #If DEBUGMODE = 1 Then
        DebugOutput "      Entering ClearComponentWhiteSpace for " & oComponent.Name
    #End If
    
    If Not oComponent Is Nothing Then
        
        #If DEBUGMODE = 1 Then
            DebugOutput "      Removing blank lines from end of " & oComponent.Name & "..."
        #End If
    
        ' Remove any blank lines from the end of the module
        lLines = oComponent.CountOfLines
        If lLines > 1 Then
            For lLineCount = lLines To 1 Step -1
                If Trim(oComponent.Lines(lLineCount, 1)) = "" Then
                    oComponent.DeleteLines lLineCount, 1
                Else
                    Exit For
                End If
            Next
        End If
        
        #If DEBUGMODE = 1 Then
            DebugOutput "      Removing blank lines from top of " & oComponent.Name & "..."
        #End If
        
        ' Remove any blank lines from the top of the module
        lLines = oComponent.CountOfLines
        If lLines > 1 Then
            lBlankCount = 0
            For lLineCount = 1 To lLines
                If Trim(oComponent.Lines(lLineCount, 1)) = "" Then
                    lBlankCount = lBlankCount + 1
                Else
                    Exit For
                End If
            Next
            If lBlankCount > 0 Then
                oComponent.DeleteLines 1, lBlankCount
            End If
        End If
        
        #If DEBUGMODE = 1 Then
            DebugOutput "      Removing repeating blank lines from " & oComponent.Name & "..."
        #End If
        
        ' Remove any repeating blank lines
        lLines = oComponent.CountOfLines
        If lLines > 1 Then
            lBlankCount = 0
            For lLineCount = lLines To 1 Step -1
                If Trim(oComponent.Lines(lLineCount, 1)) = "" Then
                    If lBlankCount >= m_lBlankLines Then
                        oComponent.DeleteLines lLineCount, 1
                        lBlankCount = lBlankCount + 1
                    Else
                        lBlankCount = lBlankCount + 1
                    End If
                Else
                    lBlankCount = 0
                End If
            Next
        End If
    End If

    #If DEBUGMODE = 1 Then
        DebugOutput "      Finished ClearComponentWhiteSpace for " & oComponent.Name & "..."
    #End If
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
' ClearCurrentModuleWhiteSpace (SUB)
'*******************************************************************************
Private Sub ClearCurrentModuleWhiteSpace()
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "ClearCurrentModuleWhiteSpace"
    
    On Error GoTo ERROR_HANDLER
    
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
    Else
        ClearComponentWhiteSpace m_VBInst.ActiveCodePane.CodeModule
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
' ClearImmediate (SUB)
'*******************************************************************************
Private Sub ClearImmediate()
    Dim oWindow As VBIDE.Window
    
    Const CTRL_HOME         As String = "^{Home}"
    Const CTRL_SHIFT_END    As String = "^+{End}"
    Const DELETE_KEY        As String = "{Del}"
    Const IMMEDIATE_WINDOW  As String = "Immediate"
    
    On Error Resume Next
    
    ' manually "clear" the control key
    ' get current state of all 256 virtual keys
    Dim keystat(0 To 255) As Byte
    
    Set oWindow = m_VBInst.ActiveWindow
    m_VBInst.Windows(IMMEDIATE_WINDOW).SetFocus
    
    GetKeyboardState keystat(0)
    ' modify bit 0 of the relevant item, and store back
    keystat(vbKeyControl) = (keystat(vbKeyControl) And &H7F)
    keystat(vbKeyShift) = (keystat(vbKeyShift) And &H7F)
    SetKeyboardState keystat(0)
    
    SendKeys CTRL_HOME, True
    
    GetKeyboardState keystat(0)
    ' modify bit 0 of the relevant item, and store back
    keystat(vbKeyControl) = (keystat(vbKeyControl) And &H7F)
    keystat(vbKeyShift) = (keystat(vbKeyShift) And &H7F)
    SetKeyboardState keystat(0)
    
    SendKeys CTRL_SHIFT_END, True
    
    GetKeyboardState keystat(0)
    ' modify bit 0 of the relevant item, and store back
    keystat(vbKeyControl) = (keystat(vbKeyControl) And &H7F)
    keystat(vbKeyShift) = (keystat(vbKeyShift) And &H7F)
    SetKeyboardState keystat(0)
    
    SendKeys DELETE_KEY, True
    oWindow.SetFocus
End Sub

'*******************************************************************************
' CloseWindows (SUB)
'*******************************************************************************
Private Sub CloseWindows()
    Dim oWindow     As VBIDE.Window
    Dim oCurrent    As VBIDE.Window
    
    On Error Resume Next
    
    Set oCurrent = m_VBInst.ActiveWindow
    
    For Each oWindow In m_VBInst.Windows
        If oWindow.Type = vbext_wt_CodeWindow Or oWindow.Type = vbext_wt_Designer Then
            If m_bCloseActive Or (Not oWindow Is oCurrent) Then
                oWindow.Close
            End If
        End If
    Next
End Sub

'*******************************************************************************
' CreateMenuItems (SUB)
'*******************************************************************************
Private Sub CreateMenuItems()
    Dim sToolbar            As String
    Dim lToolbarPosition    As Long
    Dim lToolbarIndex       As Long
    Dim lToolbarTop         As Long
    Dim lToolbarLeft        As Long
    Dim bToolbar            As Boolean
    Dim bButton0            As Boolean
    Dim bButton1            As Boolean
    Dim bButton2            As Boolean
    Dim bButton3            As Boolean
    Dim bButton4            As Boolean
    Dim bButton5            As Boolean
    Dim bButton6            As Boolean
    Dim bButton7            As Boolean
    Dim bButton8            As Boolean
    Dim bButton9            As Boolean
    Dim bButton10           As Boolean
    Dim bButton11           As Boolean
    Dim bButton12           As Boolean
    Dim bButton13           As Boolean
    Dim bButton14           As Boolean
    Dim bButton15           As Boolean
    Dim bButton16           As Boolean
    Dim bButton17           As Boolean
    Dim bButton18           As Boolean
    Dim lComboWidth         As Long
    Dim oCommandBarPopup    As CommandBarPopup
    Dim lErrNumber          As Long
    Dim sErrSource          As String
    Dim sErrDescription     As String
    Dim bShortcuts          As Boolean
    Dim bShortcut0          As Boolean
    Dim bShortcut1          As Boolean
    Dim bShortcut2          As Boolean
    Dim bShortcut3          As Boolean
    Dim bShortcut4          As Boolean
    Dim bShortcut5          As Boolean
    Dim bShortcut6          As Boolean
    Dim bShortcut7          As Boolean
    Dim bShortcut8          As Boolean
    Dim bShortcut9          As Boolean
    Dim bShortcut10         As Boolean
    Dim bShortcut11         As Boolean
    Dim bShortcut12         As Boolean
    Dim bShortcut13         As Boolean
    Dim bShortcut14         As Boolean
    Dim bShortcut15         As Boolean
    Dim bShortcut16         As Boolean
    Dim bShortcut17         As Boolean
    Dim bShortcut18         As Boolean
    Dim sMessage            As String
    Dim lPosition           As Long
    Dim oMenuBar            As CommandBar
    
    Const PROCEDURE_NAME As String = "CreateMenuItems"
    
    On Error GoTo ERROR_HANDLER
    
    #If DEBUGMODE = 1 Then
        MsgBox "Running in debug mode.", vbInformation, App.ProductName
    #End If
    
    sMessage = ""
        
    sMessage = "Calling Reset Options"
    ResetOptions
    
    sMessage = "Getting registry settings"
    lToolbarPosition = CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TOOLBARPOS, CStr(msoBarTop)))
    lToolbarIndex = CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TOOLBARINDEX, DEF_REG_TOOLBARINDEX))
    lToolbarLeft = CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TOOLBARLEFT, DEF_REG_TOOLBARLEFT))
    lToolbarTop = CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TOOLBARTOP, DEF_REG_TOOLBARTOP))
    
    bToolbar = (CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TOOLBARSHOW, vbChecked)) = vbChecked)
    bButton0 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHOW, True))
    bButton1 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHOW, True))
    bButton2 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHOW, True))
    bButton3 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHOW, True))
    bButton4 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHOW, True))
    bButton5 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHOW, True))
    bButton6 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHOW, True))
    bButton7 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHOW, True))
    bButton8 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHOW, True))
    bButton9 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHOW, True))
    bButton10 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHOW, True))
    bButton11 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHOW, True))
    bButton12 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHOW, True))
    bButton13 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHOW, True))
    bButton14 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHOW, True))
    bButton15 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHOW, True))
    bButton16 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHOW, True))
    bButton17 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHOW, True))
    bButton18 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHOW, True))
    
    lComboWidth = CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_COMBO_WIDTH, DEFAULT_COMBO_WIDTH))
    
    bShortcuts = (CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_SHORTCUTS, vbChecked)) = vbChecked)
    bShortcut0 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHORTCUT, True))
    bShortcut1 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHORTCUT, True))
    bShortcut2 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHORTCUT, True))
    bShortcut3 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHORTCUT, True))
    bShortcut4 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHORTCUT, True))
    bShortcut5 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHORTCUT, True))
    bShortcut6 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHORTCUT, True))
    bShortcut7 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHORTCUT, True))
    bShortcut8 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHORTCUT, True))
    bShortcut9 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHORTCUT, True))
    bShortcut10 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHORTCUT, True))
    bShortcut11 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHORTCUT, True))
    bShortcut12 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHORTCUT, True))
    bShortcut13 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHORTCUT, True))
    bShortcut14 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHORTCUT, True))
    bShortcut15 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHORTCUT, True))
    bShortcut16 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHORTCUT, True))
    bShortcut17 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHORTCUT, True))
    bShortcut18 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHORTCUT, True))
    
    sMessage = "Creating a key handler"
    ' Create the key handlers
    Set m_oKeyHandler = New CHotKey
    
    If bShortcuts Then
        ' Ctrl+Shift+M
        If bShortcut0 Then
            sMessage = "Registering a hot key"
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyM, vbCtrlMask + vbShiftMask
        End If
        ' Ctrl+Shift+P
        If bShortcut1 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyP, vbCtrlMask + vbShiftMask
        End If
        ' Ctrl+Shift+S
        If bShortcut2 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyS, vbCtrlMask + vbShiftMask
        End If
        ' Ctrl+Shift+T
        If bShortcut3 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyT, vbCtrlMask + vbShiftMask
        End If
        ' Ctrl+Shift+F4
        If bShortcut4 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyF4, vbCtrlMask + vbShiftMask
        End If
        ' Ctrl+Shift+Bksp
        If bShortcut5 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyBack, vbCtrlMask + vbShiftMask
        End If
        ' Ctrl+Alt+F1
        If bShortcut6 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyF1, vbCtrlMask + vbAltMask
        End If
        ' Ctrl+Alt+F2
        If bShortcut7 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyF2, vbCtrlMask + vbAltMask
        End If
        ' Ctrl+Shift+V
        If bShortcut8 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyV, vbCtrlMask + vbShiftMask
        End If
        If bShortcut9 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyV, vbCtrlMask + vbAltMask
        End If
        If bShortcut10 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyH, vbCtrlMask + vbAltMask
        End If
        If bShortcut11 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyI, vbCtrlMask + vbShiftMask
        End If
        If bShortcut12 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyI, vbCtrlMask + vbAltMask
        End If
        If bShortcut13 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyT, vbCtrlMask + vbAltMask
        End If
        If bShortcut14 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyZ, vbCtrlMask + vbShiftMask
        End If
        If bShortcut15 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyS, vbCtrlMask + vbAltMask
        End If
        If bShortcut16 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyW, vbCtrlMask + vbShiftMask
        End If
        If bShortcut17 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyW, vbCtrlMask + vbAltMask
        End If
        If bShortcut18 Then
            m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyL, vbCtrlMask + vbAltMask
        End If
    End If
    
    sMessage = "Creating the toolbar"
    On Error Resume Next
    sToolbar = m_VBInst.CommandBars(TOOLBAR_NAME).Name
    If Err.Number <> 0 Then
        m_VBInst.CommandBars.Add TOOLBAR_NAME, lToolbarPosition
    End If
    
    With m_VBInst.CommandBars(TOOLBAR_NAME)
        .Position = lToolbarPosition
    
        Select Case lToolbarPosition
            Case msoBarTop, msoBarBottom
                .RowIndex = lToolbarIndex
                .Left = lToolbarLeft
            Case msoBarLeft, msoBarRight
                .RowIndex = lToolbarIndex
                .Top = lToolbarTop
            Case Else
                .Top = lToolbarTop
                .Left = lToolbarLeft
        End Select
        
        .Visible = bToolbar
    End With
    
    Err.Clear
    On Error GoTo ERROR_HANDLER
    
    sMessage = "Adding buttons to the toolbar"
    ' Add the buttons to the edit menu
    With m_VBInst.CommandBars(TOOLBAR_NAME).Controls
        sMessage = "Adding module button to the toolbar"
        Set m_cmdModule = .Add(msoControlButton)
        sMessage = "Adding module button caption on the toolbar"
        m_cmdModule.Caption = CAPTION_MODULE
        sMessage = "Loading a picture for the module button on the toolbar"
        LoadButtonPicture RES_ID_MODULE, m_cmdModule
        sMessage = "Beginning a group for the module button on the toolbar"
        m_cmdModule.BeginGroup = True
        sMessage = "Adding a tag for the module button on the toolbar"
        m_cmdModule.Tag = BUTTON_TAG_MODULE
        sMessage = "Making the module button on the toolbar visible"
        m_cmdModule.Visible = bButton0
        
        Set m_cmdProcedure = .Add(msoControlButton)
        m_cmdProcedure.Caption = CAPTION_PROCEDURE
        LoadButtonPicture RES_ID_PROCEDURE, m_cmdProcedure
        m_cmdProcedure.Tag = BUTTON_TAG_PROCEDURE
        m_cmdProcedure.Visible = bButton1
        
        Set m_cmdTimeStamp = .Add(msoControlButton)
        m_cmdTimeStamp.Caption = CAPTION_TIMESTAMP
        LoadButtonPicture RES_ID_TIMESTAMP, m_cmdTimeStamp
        m_cmdTimeStamp.Tag = BUTTON_TAG_TIME_STAMP
        m_cmdTimeStamp.Visible = bButton2
    
        Set m_cboTemplate = .Add(msoControlDropdown)
        m_cboTemplate.Caption = CAPTION_TEMPLATE
        m_cboTemplate.Style = msoComboLabel
        m_cboTemplate.Tag = BUTTON_TAG_TEMPLATE
        m_cboTemplate.Visible = bButton3
        m_cboTemplate.Width = lComboWidth
        
        Set m_cmdClose = .Add(msoControlButton)
        m_cmdClose.Caption = CAPTION_CLOSE
        LoadButtonPicture RES_ID_CLOSE, m_cmdClose
        m_cmdClose.Tag = BUTTON_TAG_CLOSE
        m_cmdClose.Visible = bButton4
        
        Set m_cmdClear = .Add(msoControlButton)
        m_cmdClear.Caption = CAPTION_CLEAR
        LoadButtonPicture RES_ID_CLEAR, m_cmdClear
        m_cmdClear.Tag = BUTTON_TAG_CLEAR
        m_cmdClear.Visible = bButton5
        
        Set m_cmdError1 = .Add(msoControlButton)
        m_cmdError1.Caption = CAPTION_ERROR1
        LoadButtonPicture RES_ID_ERROR1, m_cmdError1
        m_cmdError1.Tag = BUTTON_TAG_ERROR1
        m_cmdError1.Visible = bButton6
        
        Set m_cmdError2 = .Add(msoControlButton)
        m_cmdError2.Caption = CAPTION_ERROR2
        LoadButtonPicture RES_ID_ERROR2, m_cmdError2
        m_cmdError2.Tag = BUTTON_TAG_ERROR2
        m_cmdError2.Visible = bButton7
        
        Set m_cmdErrorAll = .Add(msoControlButton)
        m_cmdErrorAll.Caption = CAPTION_ERROR_ALL
        LoadButtonPicture RES_ID_ERROR_ALL, m_cmdErrorAll
        m_cmdErrorAll.Tag = BUTTON_TAG_ERROR_ALL
        m_cmdErrorAll.Visible = bButton10
        
        Set m_cmdDim = .Add(msoControlButton)
        m_cmdDim.Caption = CAPTION_DIM
        LoadButtonPicture RES_ID_DIM, m_cmdDim
        m_cmdDim.Tag = BUTTON_TAG_DIM
        m_cmdDim.Visible = bButton8
        
        Set m_cmdDimAll = .Add(msoControlButton)
        m_cmdDimAll.Caption = CAPTION_DIM_ALL
        LoadButtonPicture RES_ID_DIM_ALL, m_cmdDimAll
        m_cmdDimAll.Tag = BUTTON_TAG_DIM_ALL
        m_cmdDimAll.Visible = bButton9
        
        Set m_cmdIndent = .Add(msoControlButton)
        m_cmdIndent.Caption = CAPTION_INDENT
        LoadButtonPicture RES_ID_INDENT, m_cmdIndent
        m_cmdIndent.Tag = BUTTON_TAG_INDENT
        m_cmdIndent.Visible = bButton11
        
        Set m_cmdIndentAll = .Add(msoControlButton)
        m_cmdIndentAll.Caption = CAPTION_INDENT_ALL
        LoadButtonPicture RES_ID_INDENT_ALL, m_cmdIndentAll
        m_cmdIndentAll.Tag = BUTTON_TAG_INDENT_ALL
        m_cmdIndentAll.Visible = bButton12
        
        Set m_cmdTab = .Add(msoControlButton)
        m_cmdTab.Caption = CAPTION_TAB
        LoadButtonPicture RES_ID_TAB, m_cmdTab
        m_cmdTab.Tag = BUTTON_TAG_TAB
        m_cmdTab.Visible = bButton13
        
        Set m_cmdZOrder = .Add(msoControlButton)
        m_cmdZOrder.Caption = CAPTION_ZORDER
        LoadButtonPicture RES_ID_ZORDER, m_cmdZOrder
        m_cmdZOrder.Tag = BUTTON_TAG_ZORDER
        m_cmdZOrder.Visible = bButton14
        
        Set m_cmdStats = .Add(msoControlButton)
        m_cmdStats.Caption = CAPTION_STATS
        LoadButtonPicture RES_ID_STATS, m_cmdStats
        m_cmdStats.Tag = BUTTON_TAG_STATS
        m_cmdStats.Visible = bButton15
        
        Set m_cmdWhiteSpace = .Add(msoControlButton)
        m_cmdWhiteSpace.Caption = CAPTION_WHITE_SPACE
        LoadButtonPicture RES_ID_WHITE_SPACE, m_cmdWhiteSpace
        m_cmdWhiteSpace.Tag = BUTTON_TAG_WHITE_SPACE
        m_cmdWhiteSpace.Visible = bButton16
        
        Set m_cmdWhiteSpaceAll = .Add(msoControlButton)
        m_cmdWhiteSpaceAll.Caption = CAPTION_WHITE_SPACE_ALL
        LoadButtonPicture RES_ID_WHITE_SPACE_ALL, m_cmdWhiteSpaceAll
        m_cmdWhiteSpaceAll.Tag = BUTTON_TAG_WHITE_SPACE_ALL
        m_cmdWhiteSpaceAll.Visible = bButton17
        
        Set m_cmdProcedureList = .Add(msoControlButton)
        m_cmdProcedureList.Caption = CAPTION_PROCEDURE_LIST
        LoadButtonPicture RES_ID_PROCEDURE_LIST, m_cmdProcedureList
        m_cmdProcedureList.Tag = BUTTON_TAG_PROCEDURE_LIST
        m_cmdProcedureList.Visible = bButton18
    End With

    sMessage = "Filling the template combo with a list of templates"
    ' Add list of templates to combo
    FillTemplateCombo
    
    sMessage = "Adding menu bar and items"
    ' Add menu and items
    
    On Error Resume Next
    Err.Clear
    Set oMenuBar = m_VBInst.CommandBars("Menu Bar")
    
    If Err.Number <> 0 Then
        Set oMenuBar = m_VBInst.CommandBars(1)
        
        If oMenuBar.Type <> msoBarTypeMenuBar Then
            For Each oMenuBar In m_VBInst.CommandBars
                If oMenuBar.Type = msoBarTypeMenuBar Then
                    Exit For
                End If
            Next
        End If
    End If
    On Error GoTo ERROR_HANDLER
    
    With oMenuBar
        sMessage = "Add a popup menu"
        lPosition = 0
        On Error Resume Next
        lPosition = .Controls("Window").Index
        On Error GoTo ERROR_HANDLER
        If lPosition = 0 Then
            sMessage = "Add a popup menu, no default position"
            Set oCommandBarPopup = .Controls.Add(Type:=msoControlPopup, Temporary:=True)
        Else
            sMessage = "Add a popup menu, before the Window menu"
            Set oCommandBarPopup = .Controls.Add(Type:=msoControlPopup, Before:=lPosition, Temporary:=True)
        End If
        sMessage = "Add a caption to the popup menu"
        oCommandBarPopup.Caption = MENU_NAME
        sMessage = "Add a tag to the popup menu"
        oCommandBarPopup.Tag = MENU_TAG_MAIN
    End With
    
    sMessage = "Add an item to the menu"
    Set m_mnuModule = oCommandBarPopup.Controls.Add
    sMessage = "Give the menu item a caption"
    m_mnuModule.Caption = CAPTION_MENU_MODULE
    sMessage = "Give the menu item a picture"
    LoadButtonPicture RES_ID_MODULE, m_mnuModule
    sMessage = "Give the menu item a tage"
    m_mnuModule.Tag = MENU_TAG_MODULE
    If bShortcut0 Then
        sMessage = "Give the menu item some shortcut text"
        m_mnuModule.ShortcutText = "Ctrl+Shift+M"
    End If
    
    Set m_mnuProcedure = oCommandBarPopup.Controls.Add
    m_mnuProcedure.Caption = CAPTION_MENU_PROCEDURE
    LoadButtonPicture RES_ID_PROCEDURE, m_mnuProcedure
    m_mnuProcedure.Tag = MENU_TAG_PROCEDURE
    If bShortcut1 Then
        m_mnuProcedure.ShortcutText = "Ctrl+Shift+P"
    End If
    
    Set m_mnuTimeStamp = oCommandBarPopup.Controls.Add
    m_mnuTimeStamp.Caption = CAPTION_MENU_TIMESTAMP
    LoadButtonPicture RES_ID_TIMESTAMP, m_mnuTimeStamp
    m_mnuTimeStamp.Tag = MENU_TAG_TIME_STAMP
    sMessage = "Give the menu item a group start"
    m_mnuTimeStamp.BeginGroup = True
    If bShortcut2 Then
        m_mnuTimeStamp.ShortcutText = "Ctrl+Shift+S"
    End If
    
    Set m_mnuTemplate = oCommandBarPopup.Controls.Add
    m_mnuTemplate.Caption = CAPTION_MENU_TEMPLATE
    m_mnuTemplate.Tag = MENU_TAG_TEMPLATE
    If bShortcut3 Then
        m_mnuTemplate.ShortcutText = "Ctrl+Shift+T"
    End If
    
    Set m_mnuClose = oCommandBarPopup.Controls.Add
    m_mnuClose.Caption = CAPTION_MENU_CLOSE
    LoadButtonPicture RES_ID_CLOSE, m_mnuClose
    m_mnuClose.Tag = MENU_TAG_CLOSE
    m_mnuClose.BeginGroup = True
    If bShortcut4 Then
        m_mnuClose.ShortcutText = "Ctrl+Shift+F4"
    End If
    
    Set m_mnuClear = oCommandBarPopup.Controls.Add
    m_mnuClear.Caption = CAPTION_MENU_CLEAR
    LoadButtonPicture RES_ID_CLEAR, m_mnuClear
    m_mnuClear.Tag = MENU_TAG_CLEAR
    If bShortcut5 Then
        m_mnuClear.ShortcutText = "Ctrl+Shift+Bksp"
    End If
    
    Set m_mnuError1 = oCommandBarPopup.Controls.Add
    m_mnuError1.Caption = CAPTION_MENU_ERROR1
    LoadButtonPicture RES_ID_ERROR1, m_mnuError1
    m_mnuError1.Tag = MENU_TAG_ERROR1
    m_mnuError1.BeginGroup = True
    If bShortcut6 Then
        m_mnuError1.ShortcutText = "Ctrl+Alt+F1"
    End If
    
    Set m_mnuError2 = oCommandBarPopup.Controls.Add
    m_mnuError2.Caption = CAPTION_MENU_ERROR2
    LoadButtonPicture RES_ID_ERROR2, m_mnuError2
    m_mnuError2.Tag = MENU_TAG_ERROR2
    If bShortcut7 Then
        m_mnuError2.ShortcutText = "Ctrl+Alt+F2"
    End If
    
    Set m_mnuErrorAll = oCommandBarPopup.Controls.Add
    m_mnuErrorAll.Caption = CAPTION_MENU_ERROR_ALL
    LoadButtonPicture RES_ID_ERROR_ALL, m_mnuErrorAll
    m_mnuErrorAll.Tag = MENU_TAG_ERROR_ALL
    If bShortcut10 Then
        m_mnuErrorAll.ShortcutText = "Ctrl+Alt+H"
    End If
    
    Set m_mnuDim = oCommandBarPopup.Controls.Add
    m_mnuDim.Caption = CAPTION_MENU_DIM
    LoadButtonPicture RES_ID_DIM, m_mnuDim
    m_mnuDim.Tag = MENU_TAG_DIM
    m_mnuDim.BeginGroup = True
    If bShortcut8 Then
        m_mnuDim.ShortcutText = "Ctrl+Shift+V"
    End If
    
    Set m_mnuDimAll = oCommandBarPopup.Controls.Add
    m_mnuDimAll.Caption = CAPTION_MENU_DIM_ALL
    LoadButtonPicture RES_ID_DIM_ALL, m_mnuDimAll
    m_mnuDimAll.Tag = MENU_TAG_DIM_ALL
    If bShortcut9 Then
        m_mnuDimAll.ShortcutText = "Ctrl+Alt+V"
    End If
    
    Set m_mnuIndent = oCommandBarPopup.Controls.Add
    m_mnuIndent.Caption = CAPTION_MENU_INDENT
    LoadButtonPicture RES_ID_INDENT, m_mnuIndent
    m_mnuIndent.Tag = MENU_TAG_INDENT
    m_mnuIndent.BeginGroup = True
    If bShortcut11 Then
        m_mnuIndent.ShortcutText = "Ctrl+Shift+I"
    End If
    
    Set m_mnuIndentAll = oCommandBarPopup.Controls.Add
    m_mnuIndentAll.Caption = CAPTION_MENU_INDENT_ALL
    LoadButtonPicture RES_ID_INDENT_ALL, m_mnuIndentAll
    m_mnuIndentAll.Tag = MENU_TAG_INDENT_ALL
    If bShortcut12 Then
        m_mnuIndentAll.ShortcutText = "Ctrl+Alt+I"
    End If
    
    Set m_mnuTab = oCommandBarPopup.Controls.Add
    m_mnuTab.Caption = CAPTION_MENU_TAB
    LoadButtonPicture RES_ID_TAB, m_mnuTab
    m_mnuTab.Tag = MENU_TAG_TAB
    m_mnuTab.BeginGroup = True
    If bShortcut13 Then
        m_mnuTab.ShortcutText = "Ctrl+Alt+T"
    End If
    
    Set m_mnuZOrder = oCommandBarPopup.Controls.Add
    m_mnuZOrder.Caption = CAPTION_MENU_ZORDER
    LoadButtonPicture RES_ID_ZORDER, m_mnuZOrder
    m_mnuZOrder.Tag = MENU_TAG_ZORDER
    If bShortcut14 Then
        m_mnuZOrder.ShortcutText = "Ctrl+Shift+Z"
    End If
    
    Set m_mnuStats = oCommandBarPopup.Controls.Add
    m_mnuStats.Caption = CAPTION_MENU_STATS
    LoadButtonPicture RES_ID_STATS, m_mnuStats
    m_mnuStats.Tag = MENU_TAG_STATS
    m_mnuStats.BeginGroup = True
    If bShortcut15 Then
        m_mnuStats.ShortcutText = "Ctrl+Alt+S"
    End If
    
    Set m_mnuWhiteSpace = oCommandBarPopup.Controls.Add
    m_mnuWhiteSpace.Caption = CAPTION_MENU_WHITE_SPACE
    LoadButtonPicture RES_ID_WHITE_SPACE, m_mnuWhiteSpace
    m_mnuWhiteSpace.Tag = MENU_TAG_WHITE_SPACE
    m_mnuWhiteSpace.BeginGroup = True
    If bShortcut16 Then
        m_mnuWhiteSpace.ShortcutText = "Ctrl+Shift+W"
    End If
    
    Set m_mnuWhiteSpaceAll = oCommandBarPopup.Controls.Add
    m_mnuWhiteSpaceAll.Caption = CAPTION_MENU_WHITE_SPACE_ALL
    LoadButtonPicture RES_ID_WHITE_SPACE_ALL, m_mnuWhiteSpaceAll
    m_mnuWhiteSpaceAll.Tag = MENU_TAG_WHITE_SPACE_ALL
    If bShortcut17 Then
        m_mnuWhiteSpaceAll.ShortcutText = "Ctrl+Alt+W"
    End If
    
    Set m_mnuProcedureList = oCommandBarPopup.Controls.Add
    m_mnuProcedureList.Caption = CAPTION_MENU_PROCEDURE_LIST
    LoadButtonPicture RES_ID_PROCEDURE_LIST, m_mnuProcedureList
    m_mnuProcedureList.Tag = MENU_TAG_PROCEDURE_LIST
    m_mnuProcedureList.BeginGroup = True
    If bShortcut18 Then
        m_mnuProcedureList.ShortcutText = "Ctrl+Alt+L"
    End If
    
    Set m_mnuConfigure = oCommandBarPopup.Controls.Add
    m_mnuConfigure.Caption = CAPTION_CONFIGURE
    LoadButtonPicture RES_ID_CONFIGURE, m_mnuConfigure
    m_mnuConfigure.Tag = MENU_TAG_CONFIGURE
    m_mnuConfigure.BeginGroup = True
    
    If App.HelpFile <> "" Then
        Set m_mnuHelp = oCommandBarPopup.Controls.Add
        m_mnuHelp.Caption = CAPTION_HELP
        LoadButtonPicture RES_ID_HELP, m_mnuHelp
        m_mnuHelp.Tag = MENU_TAG_HELP
        m_mnuHelp.BeginGroup = True
    End If
    
    Set m_mnuAbout = oCommandBarPopup.Controls.Add
    m_mnuAbout.Caption = CAPTION_ABOUT
    m_mnuAbout.Tag = MENU_TAG_ABOUT
    m_mnuAbout.BeginGroup = True
    
    sMessage = "Subscribing to events"
    ' Subscribe to the events
    With m_VBInst.Events
        sMessage = "Subscribe the module event"
        Set m_evtButtonModule = .CommandBarEvents(m_cmdModule)
        Set m_evtButtonProcedure = .CommandBarEvents(m_cmdProcedure)
        Set m_evtButtonTimeStamp = .CommandBarEvents(m_cmdTimeStamp)
        Set m_evtComboTemplate = .CommandBarEvents(m_cboTemplate)
        Set m_evtButtonClose = .CommandBarEvents(m_cmdClose)
        Set m_evtButtonClear = .CommandBarEvents(m_cmdClear)
        Set m_evtButtonError1 = .CommandBarEvents(m_cmdError1)
        Set m_evtButtonError2 = .CommandBarEvents(m_cmdError2)
        Set m_evtConfigure = .CommandBarEvents(m_mnuConfigure)
        Set m_evtDim = .CommandBarEvents(m_cmdDim)
        Set m_evtDimAll = .CommandBarEvents(m_cmdDimAll)
        Set m_evtErrorAll = .CommandBarEvents(m_cmdErrorAll)
        Set m_evtIndent = .CommandBarEvents(m_cmdIndent)
        Set m_evtIndentAll = .CommandBarEvents(m_cmdIndentAll)
        Set m_evtTab = .CommandBarEvents(m_cmdTab)
        Set m_evtZOrder = .CommandBarEvents(m_cmdZOrder)
        Set m_evtStats = .CommandBarEvents(m_cmdStats)
        Set m_evtWhiteSpace = .CommandBarEvents(m_cmdWhiteSpace)
        Set m_evtWhiteSpaceAll = .CommandBarEvents(m_cmdWhiteSpaceAll)
        Set m_evtProcedureList = .CommandBarEvents(m_cmdProcedureList)
        
        Set m_evtMenuClose = .CommandBarEvents(m_mnuClose)
        Set m_evtMenuClear = .CommandBarEvents(m_mnuClear)
        Set m_evtMenuModule = .CommandBarEvents(m_mnuModule)
        Set m_evtMenuProcedure = .CommandBarEvents(m_mnuProcedure)
        Set m_evtMenuTimeStamp = .CommandBarEvents(m_mnuTimeStamp)
        Set m_evtMenuTemplate = .CommandBarEvents(m_mnuTemplate)
        Set m_evtMenuError1 = .CommandBarEvents(m_mnuError1)
        Set m_evtMenuError2 = .CommandBarEvents(m_mnuError2)
        Set m_evtMenuDim = .CommandBarEvents(m_mnuDim)
        Set m_evtMenuDimAll = .CommandBarEvents(m_mnuDimAll)
        Set m_evtMenuErrorAll = .CommandBarEvents(m_mnuErrorAll)
        Set m_evtMenuIndent = .CommandBarEvents(m_mnuIndent)
        Set m_evtMenuIndentAll = .CommandBarEvents(m_mnuIndentAll)
        Set m_evtMenuTab = .CommandBarEvents(m_mnuTab)
        Set m_evtMenuZOrder = .CommandBarEvents(m_mnuZOrder)
        Set m_evtMenuStats = .CommandBarEvents(m_mnuStats)
        Set m_evtMenuWhiteSpace = .CommandBarEvents(m_mnuWhiteSpace)
        Set m_evtMenuWhiteSpaceAll = .CommandBarEvents(m_mnuWhiteSpaceAll)
        Set m_evtMenuProcedureList = .CommandBarEvents(m_mnuProcedureList)
        Set m_evtMenuAbout = .CommandBarEvents(m_mnuAbout)
        If App.HelpFile <> "" Then
            Set m_evtMenuHelp = .CommandBarEvents(m_mnuHelp)
        End If
    End With
    
    If m_sTabDisplay = "1" Then
        sMessage = "Show the tab window"
        ShowTabWindow
    End If
    
    If m_sZOrderDisplay = "1" Then
        sMessage = "Show the zorder window"
        ShowZOrderWindow
    End If

    sMessage = ""
TIDY_UP:
    On Error Resume Next

    If lErrNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Sub
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description & vbCrLf & sMessage
    sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCEDURE_NAME)
    Resume TIDY_UP
End Sub

'*******************************************************************************
' DestroyMenuItems (SUB)
'*******************************************************************************
Private Sub DestroyMenuItems()
    Const TO_FIND As Long = 42
    
    Dim oCommandBar     As CommandBar
    Dim oCommandControl As CommandBarControl
    Dim sTag            As String
    Dim oObject         As Object
    Dim bFound(TO_FIND) As Boolean
    Dim bFoundAll       As Boolean
    Dim lCount          As Long
    Dim oMenuBar        As CommandBar
    
    On Error Resume Next
    
    Set m_oKeyHandler = Nothing
    
    Set m_oConfig = Nothing
    
    ' Remember toolbar setting for next time
    With m_VBInst.CommandBars(TOOLBAR_NAME)
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TOOLBARPOS, CStr(.Position)
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TOOLBARINDEX, CStr(.RowIndex)
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TOOLBARLEFT, CStr(.Left)
        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TOOLBARTOP, CStr(.Top)
        If m_TabWindow.Visible Then
            SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TAB_DISPLAY, "1"
        Else
            SaveSetting REG_APP_NAME, REG_SETTINGS, REG_TAB_DISPLAY, "0"
        End If
        If m_ZOrderWindow.Visible Then
            SaveSetting REG_APP_NAME, REG_SETTINGS, REG_ZORDER_DISPLAY, "1"
        Else
            SaveSetting REG_APP_NAME, REG_SETTINGS, REG_ZORDER_DISPLAY, "0"
        End If
        .Visible = False
    End With
        
    ' Unsubscribe events
    Set m_evtButtonClear = Nothing
    Set m_evtButtonClose = Nothing
    Set m_evtButtonModule = Nothing
    Set m_evtButtonProcedure = Nothing
    Set m_evtButtonTimeStamp = Nothing
    Set m_evtComboTemplate = Nothing
    Set m_evtButtonError1 = Nothing
    Set m_evtButtonError2 = Nothing
    Set m_evtConfigure = Nothing
    Set m_evtDim = Nothing
    Set m_evtDimAll = Nothing
    Set m_evtErrorAll = Nothing
    Set m_evtIndent = Nothing
    Set m_evtIndentAll = Nothing
    Set m_evtTab = Nothing
    Set m_evtZOrder = Nothing
    Set m_evtStats = Nothing
    Set m_evtWhiteSpace = Nothing
    Set m_evtWhiteSpaceAll = Nothing
    Set m_evtProcedureList = Nothing
    
    Set m_evtMenuClose = Nothing
    Set m_evtMenuClear = Nothing
    Set m_evtMenuModule = Nothing
    Set m_evtMenuProcedure = Nothing
    Set m_evtMenuTimeStamp = Nothing
    Set m_evtMenuTemplate = Nothing
    Set m_evtMenuError1 = Nothing
    Set m_evtMenuError2 = Nothing
    Set m_evtMenuAbout = Nothing
    Set m_evtMenuHelp = Nothing
    Set m_evtMenuDim = Nothing
    Set m_evtMenuDimAll = Nothing
    Set m_evtMenuErrorAll = Nothing
    Set m_evtMenuIndent = Nothing
    Set m_evtMenuIndentAll = Nothing
    Set m_evtMenuTab = Nothing
    Set m_evtMenuZOrder = Nothing
    Set m_evtMenuStats = Nothing
    Set m_evtMenuWhiteSpace = Nothing
    Set m_evtMenuWhiteSpaceAll = Nothing
    Set m_evtMenuProcedureList = Nothing
    
    For lCount = 1 To TO_FIND
        bFound(lCount) = False
    Next
    
    ' Remove cleanly, in case buttons were moved
    ' When moved to another toolbar the reference remains, but does not appear
    ' to reference the button anymore
    For Each oCommandControl In m_VBInst.CommandBars(TOOLBAR_NAME).Controls
        sTag = ""
        sTag = oCommandControl.Tag
        Select Case sTag
            Case BUTTON_TAG_MODULE
                oCommandControl.Delete
                bFound(1) = True
        
            Case BUTTON_TAG_PROCEDURE
                oCommandControl.Delete
                bFound(2) = True
        
            Case BUTTON_TAG_TIME_STAMP
                oCommandControl.Delete
                bFound(3) = True
        
            Case BUTTON_TAG_TEMPLATE
                SaveSetting REG_APP_NAME, REG_SETTINGS, REG_COMBO_WIDTH, oCommandControl.Width
                oCommandControl.Delete
                bFound(4) = True
        
            Case BUTTON_TAG_CLOSE
                oCommandControl.Delete
                bFound(5) = True
        
            Case BUTTON_TAG_CLEAR
                oCommandControl.Delete
                bFound(6) = True
        
            Case BUTTON_TAG_ERROR1
                oCommandControl.Delete
                bFound(7) = True
        
            Case BUTTON_TAG_ERROR2
                oCommandControl.Delete
                bFound(8) = True
        
            Case MENU_TAG_CONFIGURE
                oCommandControl.Delete
                bFound(9) = True
        
            Case MENU_TAG_MODULE
                oCommandControl.Delete
                bFound(10) = True
        
            Case MENU_TAG_PROCEDURE
                oCommandControl.Delete
                bFound(11) = True
        
            Case MENU_TAG_TIME_STAMP
                oCommandControl.Delete
                bFound(12) = True
        
            Case MENU_TAG_TEMPLATE
                oCommandControl.Delete
                bFound(13) = True
        
            Case MENU_TAG_CLOSE
                oCommandControl.Delete
                bFound(14) = True
        
            Case MENU_TAG_CLEAR
                oCommandControl.Delete
                bFound(15) = True
        
            Case MENU_TAG_ERROR1
                oCommandControl.Delete
                bFound(16) = True
        
            Case MENU_TAG_ERROR2
                oCommandControl.Delete
                bFound(17) = True
        
            Case MENU_TAG_ABOUT
                oCommandControl.Delete
                bFound(18) = True
        
            Case MENU_TAG_MAIN
                oCommandControl.Delete
                bFound(19) = True
            
            Case MENU_TAG_HELP
                oCommandControl.Delete
                bFound(20) = True
            
            Case BUTTON_TAG_DIM
                oCommandControl.Delete
                bFound(21) = True
            
            Case MENU_TAG_DIM
                oCommandControl.Delete
                bFound(22) = True
            
            Case BUTTON_TAG_DIM_ALL
                oCommandControl.Delete
                bFound(23) = True
            
            Case MENU_TAG_DIM_ALL
                oCommandControl.Delete
                bFound(24) = True
            
            Case BUTTON_TAG_ERROR_ALL
                oCommandControl.Delete
                bFound(25) = True
            
            Case MENU_TAG_ERROR_ALL
                oCommandControl.Delete
                bFound(26) = True
            
            Case BUTTON_TAG_INDENT
                oCommandControl.Delete
                bFound(27) = True
            
            Case MENU_TAG_INDENT
                oCommandControl.Delete
                bFound(28) = True
            
            Case BUTTON_TAG_INDENT_ALL
                oCommandControl.Delete
                bFound(29) = True
            
            Case MENU_TAG_INDENT_ALL
                oCommandControl.Delete
                bFound(30) = True
            
            Case BUTTON_TAG_TAB
                oCommandControl.Delete
                bFound(31) = True
            
            Case MENU_TAG_TAB
                oCommandControl.Delete
                bFound(32) = True
            
            Case BUTTON_TAG_ZORDER
                oCommandControl.Delete
                bFound(33) = True
            
            Case MENU_TAG_ZORDER
                oCommandControl.Delete
                bFound(34) = True
            
            Case BUTTON_TAG_STATS
                oCommandControl.Delete
                bFound(35) = True
            
            Case MENU_TAG_STATS
                oCommandControl.Delete
                bFound(36) = True
            
            Case BUTTON_TAG_WHITE_SPACE
                oCommandControl.Delete
                bFound(37) = True
            
            Case BUTTON_TAG_WHITE_SPACE_ALL
                oCommandControl.Delete
                bFound(38) = True
            
            Case BUTTON_TAG_PROCEDURE_LIST
                oCommandControl.Delete
                bFound(39) = True
            
            Case MENU_TAG_WHITE_SPACE
                oCommandControl.Delete
                bFound(40) = True
            
            Case MENU_TAG_WHITE_SPACE_ALL
                oCommandControl.Delete
                bFound(41) = True
            
            Case MENU_TAG_PROCEDURE_LIST
                oCommandControl.Delete
                bFound(42) = True
        End Select
        
        bFoundAll = True
        For lCount = 1 To TO_FIND
            If Not bFound(lCount) Then
                bFoundAll = False
                Exit For
            End If
        Next
        
        If bFoundAll Then
            Exit For
        End If
    Next
    
    On Error Resume Next
    Err.Clear
    Set oMenuBar = m_VBInst.CommandBars("Menu Bar")
    
    If Err.Number <> 0 Then
        Set oMenuBar = m_VBInst.CommandBars(1)
        
        If oMenuBar.Type <> msoBarTypeMenuBar Then
            For Each oMenuBar In m_VBInst.CommandBars
                If oMenuBar.Type = msoBarTypeMenuBar Then
                    Exit For
                End If
            Next
        End If
    End If
    
    If Not bFoundAll Then
        For Each oObject In oMenuBar.Controls
            For Each oCommandControl In oObject.Controls
                sTag = ""
                sTag = oCommandControl.Tag
                Select Case sTag
                    Case MENU_TAG_CONFIGURE
                        oCommandControl.Delete
                        bFound(9) = True
                    
                    Case BUTTON_TAG_MODULE
                        oCommandControl.Delete
                        bFound(1) = True
                
                    Case BUTTON_TAG_PROCEDURE
                        oCommandControl.Delete
                        bFound(2) = True
                
                    Case BUTTON_TAG_TIME_STAMP
                        oCommandControl.Delete
                        bFound(3) = True
                
                    Case BUTTON_TAG_TEMPLATE
                        SaveSetting REG_APP_NAME, REG_SETTINGS, REG_COMBO_WIDTH, oCommandControl.Width
                        oCommandControl.Delete
                        bFound(4) = True
                
                    Case BUTTON_TAG_CLOSE
                        oCommandControl.Delete
                        bFound(5) = True
                
                    Case BUTTON_TAG_CLEAR
                        oCommandControl.Delete
                        bFound(6) = True
                
                    Case BUTTON_TAG_ERROR1
                        oCommandControl.Delete
                        bFound(7) = True
                
                    Case BUTTON_TAG_ERROR2
                        oCommandControl.Delete
                        bFound(8) = True
                
                    Case MENU_TAG_MODULE
                        oCommandControl.Delete
                        bFound(10) = True
                
                    Case MENU_TAG_PROCEDURE
                        oCommandControl.Delete
                        bFound(11) = True
                
                    Case MENU_TAG_TIME_STAMP
                        oCommandControl.Delete
                        bFound(12) = True
                
                    Case MENU_TAG_TEMPLATE
                        oCommandControl.Delete
                        bFound(13) = True
                
                    Case MENU_TAG_CLOSE
                        oCommandControl.Delete
                        bFound(14) = True
                
                    Case MENU_TAG_CLEAR
                        oCommandControl.Delete
                        bFound(15) = True
                
                    Case MENU_TAG_ERROR1
                        oCommandControl.Delete
                        bFound(16) = True
                
                    Case MENU_TAG_ERROR2
                        oCommandControl.Delete
                        bFound(17) = True
                
                    Case MENU_TAG_ABOUT
                        oCommandControl.Delete
                        bFound(18) = True
                
                    Case MENU_TAG_HELP
                        oCommandControl.Delete
                        bFound(20) = True
                
                    Case BUTTON_TAG_DIM
                        oCommandControl.Delete
                        bFound(21) = True
                    
                    Case MENU_TAG_DIM
                        oCommandControl.Delete
                        bFound(22) = True

                    Case BUTTON_TAG_DIM_ALL
                        oCommandControl.Delete
                        bFound(23) = True
                    
                    Case MENU_TAG_DIM_ALL
                        oCommandControl.Delete
                        bFound(24) = True

                    Case BUTTON_TAG_ERROR_ALL
                        oCommandControl.Delete
                        bFound(25) = True
                    
                    Case MENU_TAG_ERROR_ALL
                        oCommandControl.Delete
                        bFound(26) = True
            
                    Case BUTTON_TAG_INDENT
                        oCommandControl.Delete
                        bFound(27) = True
                    
                    Case MENU_TAG_INDENT
                        oCommandControl.Delete
                        bFound(28) = True
                    
                    Case BUTTON_TAG_INDENT_ALL
                        oCommandControl.Delete
                        bFound(29) = True
                    
                    Case MENU_TAG_INDENT_ALL
                        oCommandControl.Delete
                        bFound(30) = True
            
                    Case BUTTON_TAG_TAB
                        oCommandControl.Delete
                        bFound(31) = True
                    
                    Case MENU_TAG_TAB
                        oCommandControl.Delete
                        bFound(32) = True
            
                    Case BUTTON_TAG_ZORDER
                        oCommandControl.Delete
                        bFound(33) = True
                    
                    Case MENU_TAG_ZORDER
                        oCommandControl.Delete
                        bFound(34) = True
            
                    Case BUTTON_TAG_STATS
                        oCommandControl.Delete
                        bFound(35) = True
                    
                    Case MENU_TAG_STATS
                        oCommandControl.Delete
                        bFound(36) = True
            
                    Case BUTTON_TAG_WHITE_SPACE
                        oCommandControl.Delete
                        bFound(37) = True
                    
                    Case BUTTON_TAG_WHITE_SPACE_ALL
                        oCommandControl.Delete
                        bFound(38) = True
                    
                    Case BUTTON_TAG_PROCEDURE_LIST
                        oCommandControl.Delete
                        bFound(39) = True
                    
                    Case MENU_TAG_WHITE_SPACE
                        oCommandControl.Delete
                        bFound(40) = True
                    
                    Case MENU_TAG_WHITE_SPACE_ALL
                        oCommandControl.Delete
                        bFound(41) = True
                    
                    Case MENU_TAG_PROCEDURE_LIST
                        oCommandControl.Delete
                        bFound(42) = True
                End Select

                bFoundAll = True
                For lCount = 1 To TO_FIND
                    If Not bFound(lCount) Then
                        bFoundAll = False
                        Exit For
                    End If
                Next
                
                If bFoundAll Then
                    Exit For
                End If
            Next
            
            If bFoundAll Then
                Exit For
            End If
        Next
    End If
    
    If Not bFoundAll Then
        For Each oCommandBar In m_VBInst.CommandBars
            ' Assume they wont be dropped on these
            If oCommandBar.Visible And oCommandBar.Name <> "Menu Bar" And oCommandBar.Name <> "Shortcut Menus" And oCommandBar.Name <> TOOLBAR_NAME Then
                For Each oCommandControl In oCommandBar.Controls
                    sTag = ""
                    sTag = oCommandControl.Tag
                    Select Case sTag
                        Case BUTTON_TAG_MODULE
                            oCommandControl.Delete
                            bFound(1) = True
                    
                        Case BUTTON_TAG_PROCEDURE
                            oCommandControl.Delete
                            bFound(2) = True
                    
                        Case BUTTON_TAG_TIME_STAMP
                            oCommandControl.Delete
                            bFound(3) = True
                    
                        Case BUTTON_TAG_TEMPLATE
                            SaveSetting REG_APP_NAME, REG_SETTINGS, REG_COMBO_WIDTH, oCommandControl.Width
                            oCommandControl.Delete
                            bFound(4) = True
                    
                        Case BUTTON_TAG_CLOSE
                            oCommandControl.Delete
                            bFound(5) = True
                    
                        Case BUTTON_TAG_CLEAR
                            oCommandControl.Delete
                            bFound(6) = True
                    
                        Case BUTTON_TAG_ERROR1
                            oCommandControl.Delete
                            bFound(7) = True
                    
                        Case BUTTON_TAG_ERROR2
                            oCommandControl.Delete
                            bFound(8) = True
                    
                        Case MENU_TAG_CONFIGURE
                            oCommandControl.Delete
                            bFound(9) = True
                    
                        Case MENU_TAG_MODULE
                            oCommandControl.Delete
                            bFound(10) = True
                    
                        Case MENU_TAG_PROCEDURE
                            oCommandControl.Delete
                            bFound(11) = True
                    
                        Case MENU_TAG_TIME_STAMP
                            oCommandControl.Delete
                            bFound(12) = True
                    
                        Case MENU_TAG_TEMPLATE
                            oCommandControl.Delete
                            bFound(13) = True
                    
                        Case MENU_TAG_CLOSE
                            oCommandControl.Delete
                            bFound(14) = True
                    
                        Case MENU_TAG_CLEAR
                            oCommandControl.Delete
                            bFound(15) = True
                    
                        Case MENU_TAG_ERROR1
                            oCommandControl.Delete
                            bFound(16) = True
                    
                        Case MENU_TAG_ERROR2
                            oCommandControl.Delete
                            bFound(17) = True
                    
                        Case MENU_TAG_ABOUT
                            oCommandControl.Delete
                            bFound(18) = True
                    
                        Case MENU_TAG_MAIN
                            oCommandControl.Delete
                            bFound(19) = True
                    
                        Case MENU_TAG_HELP
                            oCommandControl.Delete
                            bFound(20) = True
                    
                        Case BUTTON_TAG_DIM
                            oCommandControl.Delete
                            bFound(21) = True
                        
                        Case MENU_TAG_DIM
                            oCommandControl.Delete
                            bFound(22) = True
                    
                        Case BUTTON_TAG_DIM_ALL
                            oCommandControl.Delete
                            bFound(23) = True
                        
                        Case MENU_TAG_DIM_ALL
                            oCommandControl.Delete
                            bFound(24) = True
    
                        Case BUTTON_TAG_ERROR_ALL
                            oCommandControl.Delete
                            bFound(25) = True
                        
                        Case MENU_TAG_ERROR_ALL
                            oCommandControl.Delete
                            bFound(26) = True
            
                        Case BUTTON_TAG_INDENT
                            oCommandControl.Delete
                            bFound(27) = True
                        
                        Case MENU_TAG_INDENT
                            oCommandControl.Delete
                            bFound(28) = True
                        
                        Case BUTTON_TAG_INDENT_ALL
                            oCommandControl.Delete
                            bFound(29) = True
                        
                        Case MENU_TAG_INDENT_ALL
                            oCommandControl.Delete
                            bFound(30) = True
            
                        Case BUTTON_TAG_TAB
                            oCommandControl.Delete
                            bFound(31) = True
                        
                        Case MENU_TAG_TAB
                            oCommandControl.Delete
                            bFound(32) = True
            
                        Case BUTTON_TAG_ZORDER
                            oCommandControl.Delete
                            bFound(33) = True
                        
                        Case MENU_TAG_ZORDER
                            oCommandControl.Delete
                            bFound(34) = True
            
                        Case BUTTON_TAG_STATS
                            oCommandControl.Delete
                            bFound(35) = True
                        
                        Case MENU_TAG_STATS
                            oCommandControl.Delete
                            bFound(36) = True
            
                        Case BUTTON_TAG_WHITE_SPACE
                            oCommandControl.Delete
                            bFound(37) = True
                        
                        Case BUTTON_TAG_WHITE_SPACE_ALL
                            oCommandControl.Delete
                            bFound(38) = True
                        
                        Case BUTTON_TAG_PROCEDURE_LIST
                            oCommandControl.Delete
                            bFound(39) = True
                        
                        Case MENU_TAG_WHITE_SPACE
                            oCommandControl.Delete
                            bFound(40) = True
                        
                        Case MENU_TAG_WHITE_SPACE_ALL
                            oCommandControl.Delete
                            bFound(41) = True
                        
                        Case MENU_TAG_PROCEDURE_LIST
                            oCommandControl.Delete
                            bFound(42) = True
                    End Select
                    
                    bFoundAll = True
                    For lCount = 1 To TO_FIND
                        If Not bFound(lCount) Then
                            bFoundAll = False
                            Exit For
                        End If
                    Next
                    
                    If bFoundAll Then
                        Exit For
                    End If
                Next
            End If
            
            If bFoundAll Then
                Exit For
            End If
        Next
    End If
    
    If Not bFound(19) Then
        For Each oObject In oMenuBar.Controls
            If oObject.Tag = MENU_TAG_MAIN Then
                oObject.Delete
                bFound(19) = True
            End If
            If bFound(19) Then
                Exit For
            End If
        Next
    End If
    
    Set m_VBInst = Nothing
End Sub

'*******************************************************************************
' DimAllEntry (SUB)
'*******************************************************************************
Private Sub DimAllEntry()
    Dim lReply          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim sDeclaration    As String
    Dim lLineNumber     As Long
    Dim bProcedure      As Boolean
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim bDirty          As Boolean
    Dim oProject        As VBProject
    Dim lEndLine        As Long
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "DimAllEntry"
    
    On Error GoTo ERROR_HANDLER

    ' Make sure we have an active code window
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
    Else
        bDirty = False
        For Each oProject In m_VBInst.VBProjects
            If oProject.IsDirty Then
                bDirty = True
                Exit For
            End If
        Next
        
        If bDirty Then
            lReply = MsgBox("It is recommended that you save the project(s) before" & _
                " cleaning variables in all procedures, do you want to continue anyway?", _
                vbQuestion + vbYesNo, App.ProductName)
            If lReply = vbNo Then
                Exit Sub
            End If
        End If
        
        ' Get confirmation from user
        lReply = MsgBox("Clean variables in all procedures?", vbYesNo, App.ProductName)
        If lReply = vbYes Then
            With m_VBInst.ActiveCodePane.CodeModule
                ' Loop through procedures and properties
                For Each oMember In .Members
                
                    ' Document the method
                    If oMember.Type = vbext_mt_Method Then
                    
                        ' Make sure it is not a 'Declare'
                        bProcedure = False
                        If .ProcBodyLine(oMember.Name, vbext_pk_Proc) = 1 And _
                            .ProcStartLine(oMember.Name, vbext_pk_Proc) = 1 Then
                            
                            sDeclaration = " " & GetDeclarationLine(m_VBInst, 1, lEndLine)
                            
                            If (InStr(sDeclaration, " Sub " & oMember.Name) > 0 _
                                Or InStr(sDeclaration, " Function " & oMember.Name) > 0) _
                                And InStr(sDeclaration, " Declare ") = 0 Then
                                    bProcedure = True
                            End If
                        Else
                            bProcedure = True
                        End If
                        
                        If bProcedure Then
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Proc)
                            GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
            
                            If lTopLine <> 0 Then
                                RemoveUnusedVariablesInProcedure lTopLine, oMember, eptType
                            End If
                        End If
                        
                    ' Document Let, Set and Get part of property if they exist
                    ElseIf oMember.Type = vbext_mt_Property Then
'                        If m_sPropertyErrors = "1" Then
                            On Error Resume Next
                            Err.Clear
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Get)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If lTopLine <> 0 Then
                                    RemoveUnusedVariablesInProcedure lTopLine, oMember, eptType
                                End If
                            Else
                                Err.Clear
                            End If
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Set)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If lTopLine <> 0 Then
                                    RemoveUnusedVariablesInProcedure lTopLine, oMember, eptType
                                End If
                            Else
                                Err.Clear
                            End If
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Let)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If lTopLine <> 0 Then
                                    RemoveUnusedVariablesInProcedure lTopLine, oMember, eptType
                                End If
                            Else
                                Err.Clear
                            End If
                            On Error GoTo ERROR_HANDLER
'                        End If
                    End If
                Next
            End With
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
' DocumentModuleEntry (SUB)
'*******************************************************************************
Private Sub DocumentModuleEntry()
    Dim lReply          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim sTestForComment As String
    Dim sDeclaration    As String
    Dim bProcedure      As Boolean
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndLine        As Long
    
    Const PROCEDURE_NAME As String = "DocumentModuleEntry"
    
    On Error GoTo ERROR_HANDLER
        
    ' Make sure we have an active code window
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
    Else
        ' Get confirmation from user
        lReply = MsgBox("Document entire code module?", vbYesNo, App.ProductName)
        If lReply = vbYes Then
            
            ' Add module header
            sTestForComment = LTrim(m_VBInst.ActiveCodePane.CodeModule.Lines(1, 1))
            If Left(sTestForComment, 1) <> "'" And Left(sTestForComment, 3) <> "Rem" Then
                InsertHeader 1, PTUnknown, oMember, "", 0
            End If
            
            With m_VBInst.ActiveCodePane.CodeModule
                ' Loop through procedures and properties
                For Each oMember In .Members
                
                    ' Document the method
                    If oMember.Type = vbext_mt_Method Then
                    
                        ' Make sure it is not a 'Declare'
                        bProcedure = False
                        If .ProcBodyLine(oMember.Name, vbext_pk_Proc) = 1 And _
                            .ProcStartLine(oMember.Name, vbext_pk_Proc) = 1 Then
                            
                            sDeclaration = " " & GetDeclarationLine(m_VBInst, 1, lEndLine)
                            
                            If (InStr(sDeclaration, " Sub " & oMember.Name) > 0 _
                                Or InStr(sDeclaration, " Function " & oMember.Name) > 0) _
                                And InStr(sDeclaration, " Declare ") = 0 Then
                                    bProcedure = True
                            End If
                        Else
                            bProcedure = True
                        End If
                        
                        If bProcedure Then
                            lTopLine = .ProcBodyLine(oMember.Name, vbext_pk_Proc)
                            eptType = GetMethodType(m_VBInst, lTopLine, oMember.Name)
                            DocumentProcedure lTopLine, oMember, eptType, False
                        End If
                        
                    ' Document Let, Set and Get part of property if they exist
                    ElseIf oMember.Type = vbext_mt_Property Then
                        On Error Resume Next
                        Err.Clear
                        
                        lTopLine = .ProcBodyLine(oMember.Name, vbext_pk_Get)
                        If Err.Number = 0 Then
                            DocumentProcedure lTopLine, oMember, PTGet, False
                        Else
                            Err.Clear
                        End If
                        
                        lTopLine = .ProcBodyLine(oMember.Name, vbext_pk_Set)
                        If Err.Number = 0 Then
                            DocumentProcedure lTopLine, oMember, PTSet, False
                        Else
                            Err.Clear
                        End If
                        
                        lTopLine = .ProcBodyLine(oMember.Name, vbext_pk_Let)
                        If Err.Number = 0 Then
                            DocumentProcedure lTopLine, oMember, PTLet, False
                        Else
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                Next
            End With
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
' DocumentProcedure (SUB)
'*******************************************************************************
Private Sub DocumentProcedure(ByVal lLineNumber As Long, _
                              ByVal oMember As Member, _
                              ByVal eptType As enumProcedureType, _
                              ByVal bNoisy As Boolean)
    Dim sTestForComment As String
    Dim sText           As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndLine        As Long
    
    Const PROCEDURE_NAME As String = "DocumentProcedure"
    
    On Error GoTo ERROR_HANDLER
        
    With m_VBInst.ActiveCodePane.CodeModule
        ' Put together full line of declaration
        sText = GetDeclarationLine(m_VBInst, lLineNumber, lEndLine)
        
        If m_bDocumentBeforeProcedure Then
            ' Check it is not already documented (previous line a comment)
            If lLineNumber > 1 Then
                sTestForComment = LTrim(.Lines(lLineNumber - 1, 1))
                If Left(sTestForComment, 1) = "'" Or Left(sTestForComment, 3) = "Rem" Then
                    If bNoisy Then
                        MsgBox "This procedure is already documented", _
                            vbCritical, App.ProductName
                    End If
                    Exit Sub
                End If
            End If
        Else
            ' Check it is not already documented (next line a comment)
            sTestForComment = LTrim(.Lines(lEndLine + 1, 1))
            If Left(sTestForComment, 1) = "'" Or Left(sTestForComment, 3) = "Rem" Then
                If bNoisy Then
                    MsgBox "This procedure is already documented", _
                        vbCritical, App.ProductName
                End If
                Exit Sub
            End If
        End If
    End With
    
    ' Sub or Function?
    If eptType = PTUnknown Then
        eptType = GetMethodType(m_VBInst, lLineNumber, oMember.Name)
    End If
    
    InsertHeader lLineNumber, eptType, oMember, sText, lEndLine
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
' DocumentProcedureEntry (SUB)
'*******************************************************************************
Private Sub DocumentProcedureEntry()
    Dim lStartLine      As Long
    Dim lDummy          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim sTestForComment As String
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "DocumentProcedureEntry"
    
    On Error GoTo ERROR_HANDLER
    
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "Select a code pane first", _
            vbCritical, App.ProductName
    Else
        ' Find out where the cursor is at the moment
        m_VBInst.ActiveCodePane.GetSelection lStartLine, lDummy, lDummy, lDummy
        
        With m_VBInst.ActiveCodePane.CodeModule
            GetProcedureDetails m_VBInst, lStartLine, oMember, eptType, lTopLine, lEndDecLine
            
            If lTopLine <> 0 Then
                DocumentProcedure lTopLine, oMember, eptType, True
            Else
                ' We are in declaration area
                sTestForComment = LTrim(.Lines(1, 1))
                If Left(sTestForComment, 1) = "'" Or Left(sTestForComment, 3) = "Rem" Then
                    MsgBox "This declaration area is already documented", _
                        vbCritical, App.ProductName
                    Exit Sub
                End If
                InsertHeader 1, PTUnknown, oMember, "", 0
            End If
        End With
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
' ErrorAllEntry (SUB)
'*******************************************************************************
Private Sub ErrorAllEntry()
    Dim lReply          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim sDeclaration    As String
    Dim lLineNumber     As Long
    Dim bProcedure      As Boolean
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim oWindow         As VBIDE.Window
    Dim oProject        As VBProject
    Dim bDirty          As Boolean
    Dim lEndLine        As Long
    Dim lEndDecLine     As Long
    
    Const CTRL_END          As String = "^{End}"
    Const IMMEDIATE_WINDOW  As String = "Immediate"
    
    Const PROCEDURE_NAME As String = "ErrorAllEntry"
    
    On Error GoTo ERROR_HANDLER

    m_sErrorText = ""
     
    ' Make sure we have an active code window
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
    Else
        bDirty = False
        For Each oProject In m_VBInst.VBProjects
            If oProject.IsDirty Then
                bDirty = True
                Exit For
            End If
        Next
        
        If bDirty Then
            lReply = MsgBox("It is recommended that you save the project(s) before" & _
                " adding error handlers to all procedures, do you want to continue anyway?", _
                vbQuestion + vbYesNo, App.ProductName)
            If lReply = vbNo Then
                Exit Sub
            End If
        End If
        
        ' Get confirmation from user
        lReply = MsgBox("Error handle entire code module?", vbYesNo, App.ProductName)
        If lReply = vbYes Then
            With m_VBInst.ActiveCodePane.CodeModule
                ' Loop through procedures and properties
                For Each oMember In .Members
                
                    ' Document the method
                    If oMember.Type = vbext_mt_Method Then
                    
                        ' Make sure it is not a 'Declare'
                        bProcedure = False
                        If .ProcBodyLine(oMember.Name, vbext_pk_Proc) = 1 And _
                            .ProcStartLine(oMember.Name, vbext_pk_Proc) = 1 Then
                            
                            sDeclaration = " " & GetDeclarationLine(m_VBInst, 1, lEndLine)
                            
                            If (InStr(sDeclaration, " Sub " & oMember.Name) > 0 _
                                Or InStr(sDeclaration, " Function " & oMember.Name) > 0) _
                                And InStr(sDeclaration, " Declare ") = 0 Then
                                    bProcedure = True
                            End If
                        Else
                            bProcedure = True
                        End If
                        
                        If bProcedure Then
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Proc)
                            GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
            
                            If lTopLine <> 0 Then
                                If InStr(oMember.Name, "_") <> 0 Then
                                    If m_sEventErrors <> "2" Then
                                        ErrorHandleProcedure ERROR_NUMBER_1, lTopLine, oMember, eptType, False
                                    Else
                                        ErrorHandleProcedure ERROR_NUMBER_2, lTopLine, oMember, eptType, False
                                    End If
                                Else
                                    If m_sNormalErrors <> "2" Then
                                        ErrorHandleProcedure ERROR_NUMBER_1, lTopLine, oMember, eptType, False
                                    Else
                                        ErrorHandleProcedure ERROR_NUMBER_2, lTopLine, oMember, eptType, False
                                    End If
                                End If
                            End If
                        End If
                        
                    ' Document Let, Set and Get part of property if they exist
                    ElseIf oMember.Type = vbext_mt_Property Then
                        If m_sPropertyErrors = "1" Then
                            On Error Resume Next
                            Err.Clear
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Get)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If m_sNormalErrors <> "2" Then
                                    ErrorHandleProcedure ERROR_NUMBER_1, lTopLine, oMember, eptType, False
                                Else
                                    ErrorHandleProcedure ERROR_NUMBER_2, lTopLine, oMember, eptType, False
                                End If
                            Else
                                Err.Clear
                            End If
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Set)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If m_sNormalErrors <> "2" Then
                                    ErrorHandleProcedure ERROR_NUMBER_1, lTopLine, oMember, eptType, False
                                Else
                                    ErrorHandleProcedure ERROR_NUMBER_2, lTopLine, oMember, eptType, False
                                End If
                            Else
                                Err.Clear
                            End If
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Let)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If m_sNormalErrors <> "2" Then
                                    ErrorHandleProcedure ERROR_NUMBER_1, lTopLine, oMember, eptType, False
                                Else
                                    ErrorHandleProcedure ERROR_NUMBER_2, lTopLine, oMember, eptType, False
                                End If
                            Else
                                Err.Clear
                            End If
                            On Error GoTo ERROR_HANDLER
                        End If
                    End If
                Next
            End With
        End If
    End If
    
    If m_sErrorText <> "" Then
        On Error Resume Next
        
        Set oWindow = m_VBInst.ActiveWindow
        m_VBInst.Windows(IMMEDIATE_WINDOW).Visible = True
        m_VBInst.Windows(IMMEDIATE_WINDOW).SetFocus
        
        SendKeys CTRL_END, True
        SendKeys m_sErrorText, True
        
        oWindow.SetFocus
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
' ErrorHandleProcedure (SUB)
'*******************************************************************************
Private Sub ErrorHandleProcedure(ByVal lNumber As Long, _
                                 ByVal lTopLine As Long, _
                                 ByVal oMember As Member, _
                                 ByVal eptType As enumProcedureType, _
                                 ByVal bNoisy As Boolean)
    Dim lTop            As Long
    Dim lBottom         As Long
    Dim sText           As String
    Dim sProcName       As String
    Dim sEndTest        As String
    Dim sFileStart      As String
    Dim sFileEnd        As String
    Dim sTemplate       As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lStart          As Long
    Dim lLinesCount     As Long
    Dim lLastDeclare    As Long
    Dim bFinished       As Boolean
    Dim bLastLineCont   As Boolean
    Dim bFoundCode      As Boolean
    Dim eProcKind       As vbext_ProcKind
    Dim bFoundEnd       As Boolean
    Dim sTextNext       As String
    Dim sProcText       As String
    Dim lReply          As Long
    Dim sErrorName      As String
    Dim sOriginalText   As String
    Dim bComment        As Boolean
    
    Const PROCEDURE_NAME As String = "ErrorHandleProcedure"
    
    On Error GoTo ERROR_HANDLER
    
    Select Case eptType
        Case PTFunction
            sEndTest = "End Function"
            eProcKind = vbext_pk_Proc
            sErrorName = "Function " & oMember.Name & " has an existing error handler"
            
        Case PTGet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Get
            sErrorName = "Property Get " & oMember.Name & " has an existing error handler"
        
        Case PTLet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Let
            sErrorName = "Property Let " & oMember.Name & " has an existing error handler"
        
        Case PTSet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Set
            sErrorName = "Property Set " & oMember.Name & " has an existing error handler"
        
        Case PTSub
            sEndTest = "End Sub"
            eProcKind = vbext_pk_Proc
            sErrorName = "Sub " & oMember.Name & " has an existing error handler"
        
        Case Else
            Exit Sub
    End Select
    
    lTop = lTopLine
    With m_VBInst.ActiveCodePane.CodeModule
        sProcName = .ProcOfLine(lTopLine, eProcKind)
        lLinesCount = .CountOfLines
        
        ' Find first line after declaration by checking continuation characters
        bFoundEnd = False
        Do
            sOriginalText = .Lines(lTop, 1)
            sText = CleanUpLine(sOriginalText)
            lTop = lTop + 1
            If sText = sEndTest Then
                bFoundEnd = True
            ElseIf sText = "End _" Then
                sTextNext = CleanUpLine(.Lines(lTop, 1))
                If sTextNext = Mid(sEndTest, 5) Then
                    bFoundEnd = True
                End If
            End If
        Loop Until Right(sText, 2) <> " _" Or lTop > lLinesCount Or bFoundEnd
        
        If Not bFoundEnd Then
            Do
                sOriginalText = LTrim(.Lines(lTop, 1))
                If Left(sOriginalText, 1) = "'" Or Left(sOriginalText, 3) = "Rem" Then
                    bComment = True
                    lTop = lTop + 1
                    
                    sText = CleanUpLine(sOriginalText)
                    If sText = sEndTest Then
                        bFoundEnd = True
                    ElseIf sText = "End _" Then
                        sTextNext = CleanUpLine(.Lines(lTop, 1))
                        If sTextNext = Mid(sEndTest, 5) Then
                            bFoundEnd = True
                        End If
                    End If
                Else
                    bComment = False
                End If
            Loop Until Not bComment Or lTop > lLinesCount Or bFoundEnd
        End If
        
        If lTop > lLinesCount And Not bFoundEnd Then
            Exit Sub
        End If

        ' Find first line that is non blank,
        ' does not follow a line with a continuation character,
        ' and does not start Dim, Static or Const
        lStart = lTop
        bFinished = False
        bLastLineCont = False
        lLastDeclare = 0
        bFoundCode = False
        
        Do
            sText = CleanUpLine(.Lines(lStart, 1))
            lStart = lStart + 1
            If Left(sText, 3) = "Dim" Or Left(sText, 6) = "Static" Or Left(sText, 5) = "Const" Then
                lLastDeclare = lStart - 1
            ElseIf sText <> "" Then
                If sText = sEndTest Then
                    bFoundEnd = True
                ElseIf sText = "End _" Then
                    sTextNext = CleanUpLine(.Lines(lStart, 1))
                    If sTextNext = Mid(sEndTest, 5) Then
                        bFoundEnd = True
                    End If
                End If
                
                If Not bFoundEnd Then
                    bFoundCode = True
                End If
                If Not bLastLineCont Then
                    bFinished = True
                End If
            End If
            
            If Right(sText, 2) = " _" Then
                bLastLineCont = True
            Else
                bLastLineCont = False
            End If
        Loop Until bFinished Or lStart > lLinesCount Or bFoundEnd
        
        If lTop > lLinesCount And Not bFoundEnd Then
            Exit Sub
        End If
        
        ' If there were declaration lines we want to start straight after them
        ' otherwise we want to start just before the first line of code
        If lLastDeclare <> 0 Then
            lStart = lLastDeclare + 1
        Else
            lStart = lStart - 1
        End If
        
        If bFoundCode Or lLastDeclare <> 0 Then
            lTop = lStart
        End If
        
        ' Find bottom line
        lBottom = lStart
        bFoundEnd = False
        Do
            sText = CleanUpLine(.Lines(lBottom, 1))
            lBottom = lBottom + 1
            If sText = sEndTest Then
                bFoundEnd = True
            ElseIf sText = "End _" Then
                sTextNext = CleanUpLine(.Lines(lBottom, 1))
                If sTextNext = Mid(sEndTest, 5) Then
                    bFoundEnd = True
                End If
            End If
        Loop Until bFoundEnd Or lBottom > lLinesCount
        
        If lBottom > lLinesCount And Not bFoundEnd Then
            Exit Sub
        End If
        
        If bFoundEnd Then
            lBottom = lBottom - 1
        End If
    
        ' Make sure we are still in the same procedure
        If sProcName <> .ProcOfLine(lBottom, eProcKind) Then
            Exit Sub
        End If
    End With
    
    sProcText = m_VBInst.ActiveCodePane.CodeModule.Lines(lTop, lBottom - lTop + 1)
    If InStr(sProcText, "On Error") > 0 Then
        If bNoisy Then
            lReply = MsgBox("This procedure appears to already have error handling, are you sure you wish to insert an error handling template?", vbYesNo + vbQuestion, App.ProductName)
            If lReply = vbNo Then
                Exit Sub
            End If
        Else
            If m_sErrorText = "" Then
                m_sErrorText = sErrorName
            Else
                m_sErrorText = m_sErrorText & vbCrLf & sErrorName
            End If
            Exit Sub
        End If
    End If
    
    If lNumber = ERROR_NUMBER_1 Then
        sFileStart = m_sBoilerplatePath & ERROR1_START
        sFileEnd = m_sBoilerplatePath & ERROR1_END
    Else
        sFileStart = m_sBoilerplatePath & ERROR2_START
        sFileEnd = m_sBoilerplatePath & ERROR2_END
    End If
    
    sTemplate = GetFileContents(sFileEnd)
    If sTemplate <> "" Then
        ReplaceProcedureTags sTemplate, eptType, oMember
        ReplaceTags m_VBInst, sTemplate, m_sAuthor, m_sCompany, m_sDateFormat, m_sTimeFormat, m_sInitials, m_colUserTags
        
        ' Insert the template
        m_VBInst.ActiveCodePane.CodeModule.InsertLines lBottom, sTemplate
    End If
    
    sTemplate = GetFileContents(sFileStart)
    If sTemplate <> "" Then
        ReplaceProcedureTags sTemplate, eptType, oMember
        ReplaceTags m_VBInst, sTemplate, m_sAuthor, m_sCompany, m_sDateFormat, m_sTimeFormat, m_sInitials, m_colUserTags
        
        ' Insert the template
        m_VBInst.ActiveCodePane.CodeModule.InsertLines lTop, sTemplate
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
' FillTemplateCombo (SUB)
'*******************************************************************************
Private Sub FillTemplateCombo()
    Dim sWildcard       As String
    Dim lFileCount      As Long
    Dim sFile           As String
    Dim sFileArray()    As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "FillTemplateCombo"
    
    On Error GoTo ERROR_HANDLER
    
    ' Add list of templates to combo
    sWildcard = m_sTemplatePath & TEMPLATE_WILDCARD
    lFileCount = -1
    sFile = Dir(sWildcard)
    Do While sFile <> ""
        lFileCount = lFileCount + 1
        ReDim Preserve sFileArray(lFileCount)
        sFileArray(lFileCount) = sFile
        sFile = Dir()
    Loop
    
    If lFileCount > 0 Then
        QuickSortStringsStart sFileArray
    End If
    
    m_cboTemplate.Clear
    If lFileCount >= 0 Then
        For lFileCount = LBound(sFileArray) To UBound(sFileArray)
            m_cboTemplate.AddItem sFileArray(lFileCount)
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

'*******************************************************************************
' GetParams (SUB)
'*******************************************************************************
Private Sub GetParams(ByVal sLine As String, _
                      ByVal sName As String, _
                      ByVal eptType As enumProcedureType, _
                      ByRef udtParamsArray() As udtParameters, _
                      ByRef lParamCount As Long)
                      
    Dim sParams         As String
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lCount          As Long
    Dim lParamEnd       As Long
    Dim sParam          As String
    Dim udtParam        As udtParameters
    Dim lLocn           As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "GetParams"
    
    On Error GoTo ERROR_HANDLER
    
    lParamCount = 0
    
    sLine = Trim(sLine)
    
    'Find start of parameter declarations
    lStart = InStr(sLine, sName & "(")

    If lStart > 0 Then
        lStart = lStart + Len(sName) + 1
        
        ' Potential for parameter declarations
        If lStart <> Len(sLine) Then
        
            ' If we have an array return param
            If Right(sLine, 2) = "()" Then
                ' Look for second to last ")"
                lEnd = InStrRev(sLine, ")", Len(sLine) - 1)
            Else
                ' Look for last ")"
                lEnd = InStrRev(sLine, ")")
            End If
    
            ' Found ")"
            If lEnd > 1 Then
        
                ' Do we have something between the brackets?
                If lEnd > lStart Then
                    sParams = " " & Mid(sLine, lStart, lEnd - lStart) & ","
                    
                    ' How many parameters do we have (commas)
                    lParamCount = 0
                    For lCount = 1 To Len(sParams)
                        If Mid(sParams, lCount, 1) = "," Then
                            lParamCount = lParamCount + 1
                        End If
                    Next
                    
                    ' Don't document last param for Lets or Sets
                    If (eptType = PTLet Or eptType = PTSet) Then
                        lParamCount = lParamCount - 1
                    End If
                    
                    ReDim udtParamsArray(lParamCount)
                    
                    ' Remember each parameter
                    For lCount = 1 To lParamCount
                        lParamEnd = InStr(sParams, ",")
                        sParam = Left(sParams, lParamEnd - 1)
                        sParams = Mid(sParams, lParamEnd + 1)
                        
                        lLocn = InStr(sParam, " Optional ")
                        If lLocn > 0 Then
                            udtParam.IsOptional = True
                            sParam = Mid(sParam, lLocn + 9)
                        Else
                            udtParam.IsOptional = False
                        End If
                        
                        lLocn = InStr(sParam, " ByVal ")
                        If lLocn > 0 Then
                            udtParam.InOutBoth = "(In)"
                            sParam = Mid(sParam, lLocn + 6)
                        Else
                            udtParam.InOutBoth = "(In/Out)"
                            lLocn = InStr(sParam, " ByRef ")
                            If lLocn > 0 Then
                                sParam = Mid(sParam, lLocn + 6)
                            End If
                        End If
                        
                        lLocn = InStr(sParam, " ParamArray ")
                        If lLocn > 0 Then
                            sParam = Mid(sParam, lLocn + 11)
                            udtParam.IsParamArray = True
                        Else
                            udtParam.IsParamArray = False
                        End If
                        
                        lLocn = InStr(sParam, " As ")
                        If lLocn > 0 Then
                            udtParam.ParameterName = Trim(Left(sParam, lLocn))
                            sParam = Mid(sParam, lLocn + 3)
                        Else
                            udtParam.ParameterType = "Variant"
                        End If
                        
                        lLocn = InStr(sParam, " = ")
                        If lLocn > 0 Then
                            If udtParam.ParameterName = "" Then
                                udtParam.ParameterName = Trim(Left(sParam, lLocn))
                            Else
                                udtParam.ParameterType = Trim(Left(sParam, lLocn))
                            End If
                            udtParam.OptionalValue = Trim(Mid(sParam, lLocn + 3))
                        Else
                            If udtParam.ParameterName = "" Then
                                udtParam.ParameterName = Trim(sParam)
                            Else
                                udtParam.ParameterType = Trim(sParam)
                            End If
                        End If
                        
                        udtParamsArray(lCount - 1) = udtParam
                        
                        udtParam.InOutBoth = ""
                        udtParam.IsOptional = False
                        udtParam.IsParamArray = False
                        udtParam.OptionalValue = ""
                        udtParam.ParameterName = ""
                        udtParam.ParameterType = ""
                    Next
                End If
            End If
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
' GetTimeStampString (FUNCTION)
'*******************************************************************************
Private Function GetTimeStampString(ByVal sFile As String, _
                                    ByVal sIndent As String, _
                                    ByRef lLines As Long, _
                                    ByRef lWidth As Long) As String
    Dim sFileContents           As String
    Dim sFileContentsArray()    As String
    Dim lLineCount              As Long
    Dim sLine                   As String
    Dim lErrNumber              As Long
    Dim sErrSource              As String
    Dim sErrDescription         As String
    
    Const PROCEDURE_NAME As String = "GetTimeStampString"
    
    On Error GoTo ERROR_HANDLER
        
    lLines = 0
    sFileContents = GetFileContents(sFile)
    
    ReplaceTags m_VBInst, sFileContents, m_sAuthor, m_sCompany, m_sDateFormat, m_sTimeFormat, m_sInitials, m_colUserTags
    
    If sFileContents <> "" Then
        sFileContentsArray = Split(sFileContents, vbCrLf)
        sFileContents = ""
    
        ' Replace timestamp tags
        For lLineCount = LBound(sFileContentsArray) To UBound(sFileContentsArray)
            sLine = sIndent & LTrim(sFileContentsArray(lLineCount))
            lLines = lLines + 1
            
            If sFileContents = "" Then
                sFileContents = sLine
            Else
                sFileContents = sFileContents & vbCrLf & sLine
            End If
        Next
        
        lWidth = Len(sLine) + 1
        
        GetTimeStampString = sFileContents
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
' IDTExtensibility_OnAddInsUpdate (SUB)
'*******************************************************************************
Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

'*******************************************************************************
' IDTExtensibility_OnConnection (SUB)
'*******************************************************************************
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, _
                                          ByVal ConnectMode As VBIDE.vbext_ConnectMode, _
                                          ByVal AddInInst As VBIDE.AddIn, _
                                          custom() As Variant)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim sHelpfile       As String
    Dim oAddin          As AddIn
    
    Const PROCEDURE_NAME As String = "IDTExtensibility_OnConnection"
    
    On Error GoTo ERROR_HANDLER
    
    m_bDocumentBeforeProcedure = True
    
    Set m_VBInst = VBInst
    
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
    
    If ConnectMode <> vbext_cm_Startup Then
        ' if no AfterStartup method will be called, call it manually
        IDTExtensibility_OnStartupComplete custom()
    End If
    
    ' Already running?
    If Not m_TabWindow Is Nothing Then
        ShowTabWindow
        ShowZOrderWindow
        Exit Sub
    End If
    
    If ConnectMode = vbext_cm_External Then
        ' the add-in has been activated from an external program
        ' (e.g. the Addin Toolbar or after a click on a VBZ file)
        On Error Resume Next
        Set oAddin = m_VBInst.Addins("VBCodeHelper.CCodeHelper")
        On Error GoTo ERROR_HANDLER
        
        If oAddin Is Nothing Then
            On Error Resume Next
            Err.Clear
            Set m_TabWindow = m_VBInst.Windows.CreateToolWindow(m_VBInst.Addins(1), "VBCodeHelper.docTabOrder", "Smart Tab Order", GUID_TAB, m_docTab)
            If Err.Number <> 0 Then
                m_bTabLoaded = False
            End If
            Err.Clear
            Set m_ZOrderWindow = m_VBInst.Windows.CreateToolWindow(m_VBInst.Addins(1), "VBCodeHelper.docZOrder", "ZOrder Management", GUID_ZORDER, m_docZOrder)
            If Err.Number <> 0 Then
                m_bZOrderLoaded = False
            End If
            Err.Clear
            Set m_StatsWindow = m_VBInst.Windows.CreateToolWindow(m_VBInst.Addins(1), "VBCodeHelper.docStatistics", "Project Statistics", GUID_STATS, m_docStats)
            If Err.Number <> 0 Then
                m_bStatsLoaded = False
            End If
            Err.Clear
            Set m_TemplateWindow = m_VBInst.Windows.CreateToolWindow(m_VBInst.Addins(1), "VBCodeHelper.docTemplates", "Insert Template", GUID_TEMPLATES, m_docTemplates)
            If Err.Number <> 0 Then
                m_bTemplatesLoaded = False
            End If
            Err.Clear
            Set m_ProcedureListWindow = m_VBInst.Windows.CreateToolWindow(m_VBInst.Addins(1), "VBCodeHelper.docProcedureList", "Procedure List", GUID_PROCEDURE_LIST, m_docProcedureList)
            If Err.Number <> 0 Then
                m_bProcedureListLoaded = False
            End If
            Err.Clear
            On Error GoTo ERROR_HANDLER
        Else
            If oAddin.Connect = False Then
                Set m_TabWindow = m_VBInst.Windows.CreateToolWindow(oAddin, "VBCodeHelper.docTabOrder", "Smart Tab Order", GUID_TAB, m_docTab)
                Set m_ZOrderWindow = m_VBInst.Windows.CreateToolWindow(oAddin, "VBCodeHelper.docZOrder", "ZOrder Management", GUID_ZORDER, m_docZOrder)
                Set m_StatsWindow = m_VBInst.Windows.CreateToolWindow(oAddin, "VBCodeHelper.docStatistics", "Project Statistics", GUID_STATS, m_docStats)
                Set m_TemplateWindow = m_VBInst.Windows.CreateToolWindow(oAddin, "VBCodeHelper.docTemplates", "Insert Template", GUID_TEMPLATES, m_docTemplates)
                Set m_ProcedureListWindow = m_VBInst.Windows.CreateToolWindow(oAddin, "VBCodeHelper.docProcedureList", "Procedure List", GUID_PROCEDURE_LIST, m_docProcedureList)
            End If
        End If
    Else
        Set m_TabWindow = m_VBInst.Windows.CreateToolWindow(AddInInst, "VBCodeHelper.docTabOrder", "Smart Tab Order", GUID_TAB, m_docTab)
        Set m_ZOrderWindow = m_VBInst.Windows.CreateToolWindow(AddInInst, "VBCodeHelper.docZOrder", "ZOrder Management", GUID_ZORDER, m_docZOrder)
        Set m_StatsWindow = m_VBInst.Windows.CreateToolWindow(AddInInst, "VBCodeHelper.docStatistics", "Project Statistics", GUID_STATS, m_docStats)
        Set m_TemplateWindow = m_VBInst.Windows.CreateToolWindow(AddInInst, "VBCodeHelper.docTemplates", "Insert Template", GUID_TEMPLATES, m_docTemplates)
        Set m_ProcedureListWindow = m_VBInst.Windows.CreateToolWindow(AddInInst, "VBCodeHelper.docProcedureList", "Procedure List", GUID_PROCEDURE_LIST, m_docProcedureList)
    End If
    
    Set m_docTab.VBInst = m_VBInst
    Set m_docZOrder.VBInst = m_VBInst
    Set m_docStats.VBInst = m_VBInst
    Set m_docTemplates.VBInst = m_VBInst
    Set m_docProcedureList.VBInst = m_VBInst
    
    Set m_evtProjects = m_VBInst.Events.VBProjectsEvents
    Set m_evtComponents = m_VBInst.Events.VBComponentsEvents(Nothing)
    Set m_evtControls = m_VBInst.Events.VBControlsEvents(Nothing, Nothing)

    If ConnectMode = vbext_cm_External Then
        ShowTabWindow
        ShowZOrderWindow
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
' IDTExtensibility_OnDisconnection (SUB)
'*******************************************************************************
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, _
                                             custom() As Variant)
    On Error Resume Next
    
    If Not m_docProcedureList Is Nothing Then
        m_docProcedureList.TidyUp
    End If
    
    DestroyMenuItems
    
    Set m_docTab.VBInst = Nothing
    Set m_docZOrder.VBInst = Nothing
    Set m_docStats.VBInst = Nothing
    Set m_docTemplates.VBInst = Nothing
    Set m_docProcedureList.VBInst = Nothing
    
    Set m_docStats = Nothing
    Set m_StatsWindow = Nothing
    Set m_docTab = Nothing
    Set m_TabWindow = Nothing
    Set m_docZOrder = Nothing
    Set m_ZOrderWindow = Nothing
    Set m_docTemplates = Nothing
    Set m_TemplateWindow = Nothing
    Set m_docProcedureList = Nothing
    Set m_ProcedureListWindow = Nothing
End Sub

'*******************************************************************************
' IDTExtensibility_OnStartupComplete (SUB)
'*******************************************************************************
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "IDTExtensibility_OnStartupComplete"
    
    On Error GoTo ERROR_HANDLER
    
    CreateMenuItems
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
' InsertErrorHandler1 (SUB)
'*******************************************************************************
Private Sub InsertErrorHandler1()
    Dim lStartLine      As Long
    Dim lDummy          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "InsertErrorHandler1"
    
    On Error GoTo ERROR_HANDLER
    
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "Select a code pane first", _
            vbCritical, App.ProductName
    Else
        ' Find out where the cursor is at the moment
        m_VBInst.ActiveCodePane.GetSelection lStartLine, lDummy, lDummy, lDummy
        
        With m_VBInst.ActiveCodePane.CodeModule
            GetProcedureDetails m_VBInst, lStartLine, oMember, eptType, lTopLine, lEndDecLine
            
            If lTopLine <> 0 Then
                ErrorHandleProcedure ERROR_NUMBER_1, lTopLine, oMember, eptType, True
            Else
                MsgBox "You must be within a procedure in order to add error handling", _
                    vbCritical, App.ProductName
            End If
        End With
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
' InsertErrorHandler2 (SUB)
'*******************************************************************************
Private Sub InsertErrorHandler2()
    Dim lStartLine      As Long
    Dim lDummy          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "InsertErrorHandler2"
    
    On Error GoTo ERROR_HANDLER
    
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "Select a code pane first", _
            vbCritical, App.ProductName
    Else
        ' Find out where the cursor is at the moment
        m_VBInst.ActiveCodePane.GetSelection lStartLine, lDummy, lDummy, lDummy
        
        With m_VBInst.ActiveCodePane.CodeModule
            GetProcedureDetails m_VBInst, lStartLine, oMember, eptType, lTopLine, lEndDecLine
            
            If lTopLine <> 0 Then
                ErrorHandleProcedure ERROR_NUMBER_2, lTopLine, oMember, eptType, True
            Else
                MsgBox "You must be within a procedure in order to add error handling", _
                    vbCritical, App.ProductName
            End If
        End With
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
' InsertHeader (SUB)
'*******************************************************************************
Private Sub InsertHeader(ByVal lTopLine As Long, _
                         ByVal eptType As enumProcedureType, _
                         ByVal oMember As Member, _
                         ByVal sDeclaration As String, _
                         ByVal lEndLine As Long)
    Dim sFile                   As String
    Dim sTemplate               As String
    Dim sLine                   As String
    Dim sFileContents           As String
    Dim sFileContentsArray()    As String
    Dim lLineCount              As Long
    Dim lErrNumber              As Long
    Dim sErrSource              As String
    Dim sErrDescription         As String
    
    Const PROCEDURE_NAME As String = "InsertHeader"
    
    On Error GoTo ERROR_HANDLER
        
    ' Put together full filename for header template
    Select Case eptType
        Case PTSub
            sFile = m_sBoilerplatePath & HEADER_SUB
        Case PTFunction
            sFile = m_sBoilerplatePath & HEADER_FUNCTION
        Case PTGet
            sFile = m_sBoilerplatePath & HEADER_GET
        Case PTLet
            sFile = m_sBoilerplatePath & HEADER_LET
        Case PTSet
            sFile = m_sBoilerplatePath & HEADER_SET
        Case Else
            sFile = m_sBoilerplatePath & HEADER_MODULE
    End Select
    
    sFileContents = GetFileContents(sFile)
    
    ReplaceProcedureTags sFileContents, eptType, oMember
    ReplaceTags m_VBInst, sFileContents, m_sAuthor, m_sCompany, m_sDateFormat, m_sTimeFormat, _
        m_sInitials, m_colUserTags
        
    If sFileContents <> "" Then
        sFileContentsArray = Split(sFileContents, vbCrLf)
        sFileContents = ""
        
        For lLineCount = LBound(sFileContentsArray) To UBound(sFileContentsArray)
            sLine = sFileContentsArray(lLineCount)
            
            ReplaceHeaderTags sLine, eptType, sDeclaration, oMember
                
            If sTemplate <> "" Then
                sTemplate = sTemplate & vbCrLf & sLine
            Else
                sTemplate = sLine
            End If
        Next lLineCount
        
        ' lEndLine is 0 when documenting the declaration area
        If lEndLine = 0 Then
            m_VBInst.ActiveCodePane.CodeModule.InsertLines lTopLine, sTemplate
        Else
            ' Test whether we want the comment before or after the procedure declaraion
            If m_bDocumentBeforeProcedure Then
                m_VBInst.ActiveCodePane.CodeModule.InsertLines lTopLine, sTemplate
            Else
                m_VBInst.ActiveCodePane.CodeModule.InsertLines lEndLine + 1, sTemplate
            End If
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
' InsertTimeStampEntry (SUB)
'*******************************************************************************
Private Sub InsertTimeStampEntry()
    Dim lStartLine      As Long
    Dim lStartCol       As Long
    Dim lEndLine        As Long
    Dim lEndCol         As Long
    Dim sText           As String
    Dim sTimeStamp      As String
    Dim lChar           As Long
    Dim sSpace          As String
    Dim lLine           As Long
    Dim lLines          As Long
    Dim sFile           As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim bFoundCont      As Boolean
    Dim lLineCount      As Long
    
    Const PROCEDURE_NAME As String = "InsertTimeStampEntry"
    
    On Error GoTo ERROR_HANDLER
    
    ' Make sure user has an active code window
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
    Else
        With m_VBInst.ActiveCodePane
            ' Find out where the cursor is at the moment
            .GetSelection lStartLine, lStartCol, lEndLine, lEndCol
            
            ' Make sure the start line is not in the middle of a continuated line
            bFoundCont = False
            If lStartLine <> 1 Then
                lLine = lStartLine
                Do
                    lLine = lLine - 1
                    sText = Trim(.CodeModule.Lines(lLine, 1))
                    sText = CleanUpLine(sText)
                    If sText <> "" Then
                        If Right(sText, 2) = " _" Then
                            bFoundCont = True
                        End If
                    End If
                Loop Until sText = "" Or Right(sText, 2) <> " _" Or lLine = 1
            End If
            If bFoundCont Then
                If lStartLine = lEndLine Then
                    If sText = "" Or Right(sText, 2) <> " _" Then
                        lStartLine = lLine + 1
                        lEndLine = lStartLine
                    Else
                        lStartLine = lLine
                        lEndLine = lStartLine
                    End If
                Else
                    If sText = "" Or Right(sText, 2) <> " _" Then
                        lStartLine = lLine + 1
                    Else
                        lStartLine = lLine
                    End If
                End If
            End If
            
            ' Make sure the end line is not in the middle of a continuated line
            If lStartLine <> lEndLine Then
                lLineCount = .CodeModule.CountOfLines
                
                bFoundCont = False
                If lEndLine <> lLineCount Then
                    If lEndCol = 1 Then
                        lLine = lEndLine - 1
                    Else
                        lLine = lEndLine
                    End If
                    sText = Trim(.CodeModule.Lines(lLine, 1))
                    sText = CleanUpLine(sText)
                    Do Until sText = "" Or Right(sText, 2) <> " _" Or lLine = lLineCount
                        bFoundCont = True
                        lLine = lLine + 1
                        sText = Trim(.CodeModule.Lines(lLine, 1))
                        sText = CleanUpLine(sText)
                    Loop
                End If
                If bFoundCont Then
                    lEndLine = lLine + 1
                End If
            End If
            
            ' Find first non empty line (in order to judge indent)
            lLine = lStartLine - 1
            Do
                lLine = lLine + 1
                sText = .CodeModule.Lines(lLine, 1)
            Loop Until Trim(sText) <> "" Or lLine >= .CodeModule.CountOfLines
                        
            ' Put together the indent
            sSpace = ""
            If Trim(sText) <> "" Then
                lChar = Len(sText) - Len(LTrim(sText))
                sSpace = Space(lChar)
            End If
            
            ' Time stamp one line or a range?
            If lStartLine = lEndLine Then
                sFile = m_sBoilerplatePath & TIMESTAMP
                sTimeStamp = GetTimeStampString(sFile, sSpace, lLines, lStartCol)
                If sTimeStamp <> "" Then
                    .CodeModule.InsertLines lStartLine, sTimeStamp
                    lStartLine = lStartLine + lLines - 1
                    lEndLine = lStartLine
                    lEndCol = lStartCol
                    .SetSelection lStartLine, lStartCol, lEndLine, lEndCol
                End If
            Else
                ' Time stamp end part first so line numbering not altered
                sFile = m_sBoilerplatePath & TIMESTAMP_END
                sTimeStamp = GetTimeStampString(sFile, sSpace, lLines, lStartCol)
                If sTimeStamp <> "" Then
                    If lEndCol <> 1 Then
                        lEndLine = lEndLine + 1
                    End If
                    .CodeModule.InsertLines lEndLine, sTimeStamp
                End If
                
                sFile = m_sBoilerplatePath & TIMESTAMP_START
                sTimeStamp = GetTimeStampString(sFile, sSpace, lLines, lStartCol)
                If sTimeStamp <> "" Then
                    .CodeModule.InsertLines lStartLine, sTimeStamp
                    lStartLine = lStartLine + lLines - 1
                    lEndLine = lStartLine
                    lEndCol = lStartCol
                    .SetSelection lStartLine, lStartCol, lEndLine, lEndCol
                End If
            End If
        End With
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
' m_evtButtonClear_Click (SUB)
'*******************************************************************************
Private Sub m_evtButtonClear_Click(ByVal CommandBarControl As Object, _
                                   handled As Boolean, _
                                   CancelDefault As Boolean)
    On Error Resume Next
    
    ClearImmediate
End Sub

'*******************************************************************************
' m_evtButtonClose_Click (SUB)
'*******************************************************************************
Private Sub m_evtButtonClose_Click(ByVal CommandBarControl As Object, _
                                   handled As Boolean, _
                                   CancelDefault As Boolean)
    
    On Error Resume Next
    
    CloseWindows
End Sub

'*******************************************************************************
' m_evtButtonError1_Click (SUB)
'*******************************************************************************
Private Sub m_evtButtonError1_Click(ByVal CommandBarControl As Object, _
                                    handled As Boolean, _
                                    CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtButtonError1_Click"
    
    On Error GoTo ERROR_HANDLER
    
    InsertErrorHandler1
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
' m_evtButtonError2_Click (SUB)
'*******************************************************************************
Private Sub m_evtButtonError2_Click(ByVal CommandBarControl As Object, _
                                    handled As Boolean, _
                                    CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtButtonError2_Click"
    
    On Error GoTo ERROR_HANDLER
    
    InsertErrorHandler2
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
' m_evtButtonModule_Click (SUB)
'*******************************************************************************
Private Sub m_evtButtonModule_Click(ByVal CommandBarControl As Object, _
                                    handled As Boolean, _
                                    CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtButtonModule_Click"
    
    On Error GoTo ERROR_HANDLER
    
    DocumentModuleEntry
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
' m_evtButtonProcedure_Click (SUB)
'*******************************************************************************
Private Sub m_evtButtonProcedure_Click(ByVal CommandBarControl As Object, _
                                       handled As Boolean, _
                                       CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtButtonProcedure_Click"
    
    On Error GoTo ERROR_HANDLER
    
    DocumentProcedureEntry
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
' m_evtButtonTimeStamp_Click (SUB)
'*******************************************************************************
Private Sub m_evtButtonTimeStamp_Click(ByVal CommandBarControl As Object, _
                                       handled As Boolean, _
                                       CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtButtonTimeStamp_Click"
    
    On Error GoTo ERROR_HANDLER
    
    InsertTimeStampEntry
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
' m_evtComboTemplate_Click (SUB)
'*******************************************************************************
Private Sub m_evtComboTemplate_Click(ByVal CommandBarControl As Object, _
                                     handled As Boolean, _
                                     CancelDefault As Boolean)
    Dim sFile           As String
    Dim lCurrentLine    As Long
    Dim lDummy          As Long
    Dim sTemplate       As String
    Dim oMember         As Member
    Dim eptType         As enumProcedureType
    Dim lTopLine        As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "m_evtComboTemplate_Click"
    
    On Error GoTo ERROR_HANDLER

    ' Check user has a code window
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
    Else
        sFile = m_sTemplatePath & m_cboTemplate.Text
        sTemplate = GetFileContents(sFile)
        
        If sTemplate <> "" Then
            With m_VBInst.ActiveCodePane
                ' Find out where the cursor is at the moment
                .GetSelection lCurrentLine, lDummy, lDummy, lDummy
                
                GetProcedureDetails m_VBInst, lCurrentLine, oMember, eptType, lTopLine, lEndDecLine
                ReplaceProcedureTags sTemplate, eptType, oMember
                ReplaceTags m_VBInst, sTemplate, m_sAuthor, m_sCompany, m_sDateFormat, m_sTimeFormat, m_sInitials, m_colUserTags
                
                ReplaceObjectStartEndSelection sTemplate
            
                ' Insert the template
                If Len(Trim(sTemplate)) > 0 Then
                    .CodeModule.InsertLines lCurrentLine, sTemplate
                End If
            End With
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
' ReplaceObjectStartEndSelection (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' New procedure to implement functionality provided by end user
'*******************************************************************************
Private Sub ReplaceObjectStartEndSelection(sTemplate As String)
    If InStr(sTemplate, TAG_MAXOBJERROR) > 0 Then
    End If

    If InStr(sTemplate, TAG_NEXTOBJERROR) > 0 Then
    End If

    If InStr(sTemplate, TAG_SELECTION) > 0 Or InStr(sTemplate, TAG_FIRSTCODELINE) > 0 Or InStr(sTemplate, TAG_LASTCODELINE) Then
        ' Get first code line
        
        ' Check for selection tag and selected text
    End If
End Sub

'*******************************************************************************
' m_evtComponents_ItemActivated (SUB)
'*******************************************************************************
Private Sub m_evtComponents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtComponents_ItemActivated"
    
    On Error GoTo ERROR_HANDLER
    
    If m_TabWindow.Visible Then
        m_docTab.RefreshList 0
    End If
    
    If m_ZOrderWindow.Visible Then
        m_docZOrder.RefreshList 0
    End If
    
    If m_ProcedureListWindow.Visible Then
        m_ProcedureListWindow.Visible = False
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
' m_evtComponents_ItemSelected (SUB)
'*******************************************************************************
Private Sub m_evtComponents_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtComponents_ItemSelected"
    
    On Error GoTo ERROR_HANDLER
    
    m_evtComponents_ItemActivated VBComponent
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
' m_evtButtonClear_Click (SUB)
'*******************************************************************************
Private Sub m_evtConfigure_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtConfigure_Click"
    
    On Error GoTo ERROR_HANDLER
    If m_oConfig Is Nothing Then
        Set m_oConfig = New VBCodeHelperConfig.CEntryPoint
    End If
    
    m_oConfig.ShowConfigForm False, g_sLicenseKey, g_sLicensedTo
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
' m_evtControls_ItemAdded (SUB)
'*******************************************************************************
Private Sub m_evtControls_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtControls_ItemAdded"
    
    On Error GoTo ERROR_HANDLER
    
    If m_TabWindow.Visible Then
        m_docTab.ControlAdded VBControl
    End If
    
    If m_ZOrderWindow.Visible Then
        m_docZOrder.ControlAdded VBControl
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
' m_evtControls_ItemRemoved (SUB)
'*******************************************************************************
Private Sub m_evtControls_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtControls_ItemRemoved"
    
    On Error GoTo ERROR_HANDLER
    
    If m_TabWindow.Visible Then
        m_docTab.ControlRemoved VBControl
    End If
    
    If m_ZOrderWindow.Visible Then
        m_docZOrder.ControlRemoved VBControl
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
' m_evtControls_ItemRenamed (SUB)
'*******************************************************************************
Private Sub m_evtControls_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtControls_ItemRenamed"
    
    On Error GoTo ERROR_HANDLER
    
    If m_TabWindow.Visible Then
        m_docTab.ControlRenamed VBControl, OldName, OldIndex
    End If
    
    If m_ZOrderWindow.Visible Then
        m_docZOrder.ControlRenamed VBControl, OldName, OldIndex
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
' m_evtDim_Click (SUB)
'*******************************************************************************
Private Sub m_evtDim_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtDim_Click"
    
    On Error GoTo ERROR_HANDLER
    
    CleanUnusedVariablesProcedure
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
' m_evtDimAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtDimAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtDimAll_Click"
    
    On Error GoTo ERROR_HANDLER
    
    DimAllEntry
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
' m_evtErrorAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtErrorAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtErrorAll_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ErrorAllEntry
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
' m_evtIndent_Click (SUB)
'*******************************************************************************
Private Sub m_evtIndent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtIndent_Click"
    
    On Error GoTo ERROR_HANDLER
    
    SmartIndentProcedure
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
' m_evtIndentAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtIndentAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtIndentAll_Click"
    
    On Error GoTo ERROR_HANDLER
    
    SmartIndentAllProcedures
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
' m_evtMenuAbout_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuAbout_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuAbout_Click"
    
    On Error GoTo ERROR_HANDLER
    
    Set m_frmAbout = New FAbout
    
    m_frmAbout.Show
            
    SetWindowPos m_frmAbout.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    
    SetForegroundWindow m_frmAbout.hwnd
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
' m_evtMenuClear_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuClear_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    
    ClearImmediate
End Sub

'*******************************************************************************
' m_evtMenuClose_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuClose_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    
    CloseWindows
End Sub

'*******************************************************************************
' m_evtMenuDim_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuDim_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuDim_Click"
    
    On Error GoTo ERROR_HANDLER
    
    CleanUnusedVariablesProcedure
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
' m_evtMenuDimAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuDimAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuDimAll_Click"
    
    On Error GoTo ERROR_HANDLER
    
    DimAllEntry
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
' m_evtMenuError1_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuError1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuError1_Click"
    
    On Error GoTo ERROR_HANDLER
    
    InsertErrorHandler1
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
' m_evtMenuError2_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuError2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuError2_Click"
    
    On Error GoTo ERROR_HANDLER
    
    InsertErrorHandler2
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
' m_evtMenuErrorAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuErrorAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuErrorAll_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ErrorAllEntry
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
' m_evtMenuHelp_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuHelp_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuHelp_Click"
    
    On Error GoTo ERROR_HANDLER
    
    If App.HelpFile <> "" Then
        WinHelp FDummy.hwnd, App.HelpFile, HELP_FINDER, ""
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
' m_evtMenuIndent_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuIndent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuIndent_Click"
    
    On Error GoTo ERROR_HANDLER
    
    SmartIndentProcedure
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
' m_evtMenuIndentAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuIndentAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuIndentAll_Click"
    
    On Error GoTo ERROR_HANDLER
    
    SmartIndentAllProcedures
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
' m_evtMenuModule_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuModule_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuModule_Click"
    
    On Error GoTo ERROR_HANDLER
    
    DocumentModuleEntry
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
' m_evtMenuProcedure_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuProcedure_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuProcedure_Click"
    
    On Error GoTo ERROR_HANDLER
    
    DocumentProcedureEntry
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
' m_evtMenuProcedureList_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuProcedureList_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuProcedureList_Click"
    
    On Error GoTo ERROR_HANDLER
    ProcedureList
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
' m_evtMenuStats_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuStats_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuStats_Click"
    
    On Error GoTo ERROR_HANDLER
    
    CalculateStats
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
' m_evtMenuTab_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuTab_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuTab_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ShowTabWindow
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
' m_evtMenuTemplate_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuTemplate_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuTemplate_Click"
    
    On Error GoTo ERROR_HANDLER
    
    TemplateMenuEntry
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
' m_evtMenuTimeStamp_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuTimeStamp_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuTimeStamp_Click"
    
    On Error GoTo ERROR_HANDLER
    
    InsertTimeStampEntry
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
' m_evtMenuWhiteSpace_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuWhiteSpace_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuWhiteSpace_Click"
    
    On Error GoTo ERROR_HANDLER
    ClearCurrentModuleWhiteSpace
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
' m_evtMenuWhiteSpaceAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuWhiteSpaceAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuWhiteSpaceAll_Click"
    
    On Error GoTo ERROR_HANDLER
    ClearAllWhiteSpace
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
' m_evtMenuZOrder_Click (SUB)
'*******************************************************************************
Private Sub m_evtMenuZOrder_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtMenuZOrder_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ShowZOrderWindow
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
' m_evtProcedureList_Click (SUB)
'*******************************************************************************
Private Sub m_evtProcedureList_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtProcedureList_Click"
    
    On Error GoTo ERROR_HANDLER
    ProcedureList
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
' m_evtProjects_ItemRemoved (SUB)
'*******************************************************************************
Private Sub m_evtProjects_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtProjects_ItemRemoved"
    
    On Error GoTo ERROR_HANDLER
    
    If m_TabWindow.Visible Then
        m_docTab.RefreshList 0
    End If
    
    If m_ZOrderWindow.Visible Then
        m_docZOrder.RefreshList 0
    End If

    m_ProcedureListWindow.Visible = False
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
' m_evtStats_Click (SUB)
'*******************************************************************************
Private Sub m_evtStats_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtStats_Click"
    
    On Error GoTo ERROR_HANDLER
    
    CalculateStats
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
' m_evtTab_Click (SUB)
'*******************************************************************************
Private Sub m_evtTab_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtTab_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ShowTabWindow
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
' m_evtWhiteSpace_Click (SUB)
'*******************************************************************************
Private Sub m_evtWhiteSpace_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtWhiteSpace_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ClearCurrentModuleWhiteSpace
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
' m_evtWhiteSpaceAll_Click (SUB)
'*******************************************************************************
Private Sub m_evtWhiteSpaceAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtWhiteSpaceAll_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ClearAllWhiteSpace
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
' m_evtZOrder_Click (SUB)
'*******************************************************************************
Private Sub m_evtZOrder_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_evtZOrder_Click"
    
    On Error GoTo ERROR_HANDLER
    
    ShowZOrderWindow
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
' m_oConfig_ConfigFinished (SUB)
'*******************************************************************************
Private Sub m_oConfig_ConfigFinished(ByVal bUpdateOptions As Boolean)
    Dim bToolbar            As Boolean
    Dim bButton0            As Boolean
    Dim bButton1            As Boolean
    Dim bButton2            As Boolean
    Dim bButton3            As Boolean
    Dim bButton4            As Boolean
    Dim bButton5            As Boolean
    Dim bButton6            As Boolean
    Dim bButton7            As Boolean
    Dim bButton8            As Boolean
    Dim bButton9            As Boolean
    Dim bButton10           As Boolean
    Dim bButton11           As Boolean
    Dim bButton12           As Boolean
    Dim bButton13           As Boolean
    Dim bButton14           As Boolean
    Dim bButton15           As Boolean
    Dim bButton16           As Boolean
    Dim bButton17           As Boolean
    Dim bButton18           As Boolean
    Dim lErrNumber          As Long
    Dim sErrSource          As String
    Dim sErrDescription     As String
    Dim bShortcuts          As Boolean
    Dim bShortcut0          As Boolean
    Dim bShortcut1          As Boolean
    Dim bShortcut2          As Boolean
    Dim bShortcut3          As Boolean
    Dim bShortcut4          As Boolean
    Dim bShortcut5          As Boolean
    Dim bShortcut6          As Boolean
    Dim bShortcut7          As Boolean
    Dim bShortcut8          As Boolean
    Dim bShortcut9          As Boolean
    Dim bShortcut10         As Boolean
    Dim bShortcut11         As Boolean
    Dim bShortcut12         As Boolean
    Dim bShortcut13         As Boolean
    Dim bShortcut14         As Boolean
    Dim bShortcut15         As Boolean
    Dim bShortcut16         As Boolean
    Dim bShortcut17         As Boolean
    Dim bShortcut18         As Boolean
    
    Const PROCEDURE_NAME As String = "m_oConfig_ConfigFinished"
    
    On Error GoTo ERROR_HANDLER
    If bUpdateOptions Then
        ResetOptions
        FillTemplateCombo
        
        On Error Resume Next
        bToolbar = (CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TOOLBARSHOW, vbChecked)) = vbChecked)
        bButton0 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHOW, True))
        bButton1 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHOW, True))
        bButton2 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHOW, True))
        bButton3 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHOW, True))
        bButton4 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHOW, True))
        bButton5 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHOW, True))
        bButton6 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHOW, True))
        bButton7 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHOW, True))
        bButton8 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHOW, True))
        bButton9 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHOW, True))
        bButton10 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHOW, True))
        bButton11 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHOW, True))
        bButton12 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHOW, True))
        bButton13 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHOW, True))
        bButton14 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHOW, True))
        bButton15 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHOW, True))
        bButton16 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHOW, True))
        bButton17 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHOW, True))
        bButton18 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHOW, True))
        
        m_VBInst.CommandBars(TOOLBAR_NAME).Visible = bToolbar
        m_cmdModule.Visible = bButton0
        m_cmdProcedure.Visible = bButton1
        m_cmdTimeStamp.Visible = bButton2
        m_cboTemplate.Visible = bButton3
        m_cmdClose.Visible = bButton4
        m_cmdClear.Visible = bButton5
        m_cmdError1.Visible = bButton6
        m_cmdError2.Visible = bButton7
        m_cmdErrorAll.Visible = bButton10
        m_cmdDim.Visible = bButton8
        m_cmdDimAll.Visible = bButton9
        m_cmdIndent.Visible = bButton11
        m_cmdIndentAll.Visible = bButton12
        m_cmdTab.Visible = bButton13
        m_cmdZOrder.Visible = bButton14
        m_cmdStats.Visible = bButton15
        m_cmdWhiteSpace.Visible = bButton16
        m_cmdWhiteSpaceAll.Visible = bButton17
        m_cmdProcedureList.Visible = bButton18
    
        bShortcuts = (CLng(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_SHORTCUTS, vbChecked)) = vbChecked)
        bShortcut0 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON0_SHORTCUT, True))
        bShortcut1 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON1_SHORTCUT, True))
        bShortcut2 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON2_SHORTCUT, True))
        bShortcut3 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON3_SHORTCUT, True))
        bShortcut4 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON4_SHORTCUT, True))
        bShortcut5 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON5_SHORTCUT, True))
        bShortcut6 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON6_SHORTCUT, True))
        bShortcut7 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON7_SHORTCUT, True))
        bShortcut8 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON8_SHORTCUT, True))
        bShortcut9 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON9_SHORTCUT, True))
        bShortcut10 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON10_SHORTCUT, True))
        bShortcut11 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON11_SHORTCUT, True))
        bShortcut12 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON12_SHORTCUT, True))
        bShortcut13 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON13_SHORTCUT, True))
        bShortcut14 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON14_SHORTCUT, True))
        bShortcut15 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON15_SHORTCUT, True))
        bShortcut16 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON16_SHORTCUT, True))
        bShortcut17 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON17_SHORTCUT, True))
        bShortcut18 = CBool(GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BUTTON18_SHORTCUT, True))
        
        If bShortcuts Then
            ' Ctrl+Shift+M
            If bShortcut0 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyM, vbCtrlMask + vbShiftMask
                m_mnuModule.ShortcutText = "Ctrl+Shift+M"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyM, vbCtrlMask + vbShiftMask
                m_mnuModule.ShortcutText = ""
            End If
            ' Ctrl+Shift+P
            If bShortcut1 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyP, vbCtrlMask + vbShiftMask
                m_mnuProcedure.ShortcutText = "Ctrl+Shift+P"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyP, vbCtrlMask + vbShiftMask
                m_mnuProcedure.ShortcutText = ""
            End If
            ' Ctrl+Shift+S
            If bShortcut2 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyS, vbCtrlMask + vbShiftMask
                m_mnuTimeStamp.ShortcutText = "Ctrl+Shift+S"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyS, vbCtrlMask + vbShiftMask
                m_mnuTimeStamp.ShortcutText = ""
            End If
            ' Ctrl+Shift+T
            If bShortcut3 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyT, vbCtrlMask + vbShiftMask
                m_mnuTemplate.ShortcutText = "Ctrl+Shift+T"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyT, vbCtrlMask + vbShiftMask
                m_mnuTemplate.ShortcutText = ""
            End If
            ' Ctrl+Shift+F4
            If bShortcut4 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyF4, vbCtrlMask + vbShiftMask
                m_mnuClose.ShortcutText = "Ctrl+Shift+F4"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyF4, vbCtrlMask + vbShiftMask
                m_mnuClose.ShortcutText = ""
            End If
            ' Ctrl+Shift+Bksp
            If bShortcut5 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyBack, vbCtrlMask + vbShiftMask
                m_mnuClear.ShortcutText = "Ctrl+Shift+Bksp"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyBack, vbCtrlMask + vbShiftMask
                m_mnuClear.ShortcutText = ""
            End If
            ' Ctrl+Alt+F1
            If bShortcut6 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyF1, vbCtrlMask + vbAltMask
                m_mnuError1.ShortcutText = "Ctrl+Alt+F1"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyF1, vbCtrlMask + vbAltMask
                m_mnuError1.ShortcutText = ""
            End If
            ' Ctrl+Alt+F2
            If bShortcut7 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyF2, vbCtrlMask + vbAltMask
                m_mnuError2.ShortcutText = "Ctrl+Alt+F2"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyF2, vbCtrlMask + vbAltMask
                m_mnuError2.ShortcutText = ""
            End If
            ' Ctrl+Shift+V
            If bShortcut8 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyV, vbCtrlMask + vbShiftMask
                m_mnuDim.ShortcutText = "Ctrl+Shift+V"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyV, vbCtrlMask + vbShiftMask
                m_mnuDim.ShortcutText = ""
            End If
            ' Ctrl+Alt+V
            If bShortcut9 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyV, vbCtrlMask + vbAltMask
                m_mnuDimAll.ShortcutText = "Ctrl+Alt+V"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyV, vbCtrlMask + vbAltMask
                m_mnuDimAll.ShortcutText = ""
            End If
            ' Ctrl+Alt+H
            If bShortcut10 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyH, vbCtrlMask + vbAltMask
                m_mnuErrorAll.ShortcutText = "Ctrl+Alt+H"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyH, vbCtrlMask + vbAltMask
                m_mnuErrorAll.ShortcutText = ""
            End If
            ' Ctrl+Shift+I
            If bShortcut11 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyI, vbCtrlMask + vbShiftMask
                m_mnuIndent.ShortcutText = "Ctrl+Shift+I"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyI, vbCtrlMask + vbShiftMask
                m_mnuIndent.ShortcutText = ""
            End If
            ' Ctrl+Alt+I
            If bShortcut12 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyI, vbCtrlMask + vbAltMask
                m_mnuIndentAll.ShortcutText = "Ctrl+Alt+I"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyI, vbCtrlMask + vbAltMask
                m_mnuIndentAll.ShortcutText = ""
            End If
            ' Ctrl+Alt+T
            If bShortcut13 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyT, vbCtrlMask + vbAltMask
                m_mnuTab.ShortcutText = "Ctrl+Alt+T"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyT, vbCtrlMask + vbAltMask
                m_mnuTab.ShortcutText = ""
            End If
            ' Ctrl+Shift+Z
            If bShortcut14 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyZ, vbCtrlMask + vbShiftMask
                m_mnuZOrder.ShortcutText = "Ctrl+Shift+Z"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyZ, vbCtrlMask + vbShiftMask
                m_mnuZOrder.ShortcutText = ""
            End If
            ' Ctrl+Alt+S
            If bShortcut15 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyS, vbCtrlMask + vbAltMask
                m_mnuStats.ShortcutText = "Ctrl+Alt+S"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyS, vbCtrlMask + vbAltMask
                m_mnuStats.ShortcutText = ""
            End If
            ' Ctrl+Shift+W
            If bShortcut16 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyW, vbCtrlMask + vbShiftMask
                m_mnuWhiteSpace.ShortcutText = "Ctrl+Shift+W"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyW, vbCtrlMask + vbShiftMask
                m_mnuWhiteSpace.ShortcutText = ""
            End If
            ' Ctrl+Alt+W
            If bShortcut17 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyW, vbCtrlMask + vbAltMask
                m_mnuWhiteSpaceAll.ShortcutText = "Ctrl+Alt+W"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyW, vbCtrlMask + vbAltMask
                m_mnuWhiteSpaceAll.ShortcutText = ""
            End If
            ' Ctrl+Alt+L
            If bShortcut18 Then
                m_oKeyHandler.RegisterHotKey FDummy.hwnd, vbKeyL, vbCtrlMask + vbAltMask
                m_mnuProcedureList.ShortcutText = "Ctrl+Alt+L"
            Else
                m_oKeyHandler.UnregisterHotKey vbKeyL, vbCtrlMask + vbAltMask
                m_mnuProcedureList.ShortcutText = ""
            End If
        Else
            m_oKeyHandler.UnregisterHotKey vbKeyM, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyP, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyS, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyT, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyF4, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyBack, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyF1, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyF2, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyV, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyV, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyH, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyI, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyI, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyT, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyZ, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyS, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyW, vbCtrlMask + vbShiftMask
            m_oKeyHandler.UnregisterHotKey vbKeyW, vbCtrlMask + vbAltMask
            m_oKeyHandler.UnregisterHotKey vbKeyL, vbCtrlMask + vbAltMask
            
            m_mnuModule.ShortcutText = ""
            m_mnuProcedure.ShortcutText = ""
            m_mnuTimeStamp.ShortcutText = ""
            m_mnuTemplate.ShortcutText = ""
            m_mnuClose.ShortcutText = ""
            m_mnuClear.ShortcutText = ""
            m_mnuError1.ShortcutText = ""
            m_mnuError2.ShortcutText = ""
            m_mnuDim.ShortcutText = ""
            m_mnuDimAll.ShortcutText = ""
            m_mnuErrorAll.ShortcutText = ""
            m_mnuIndent.ShortcutText = ""
            m_mnuIndentAll.ShortcutText = ""
            m_mnuTab.ShortcutText = ""
            m_mnuZOrder.ShortcutText = ""
            m_mnuStats.ShortcutText = ""
            m_mnuWhiteSpace.ShortcutText = ""
            m_mnuWhiteSpaceAll.ShortcutText = ""
            m_mnuProcedureList.ShortcutText = ""
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
' m_oKeyHandler_HotKey (SUB)
'*******************************************************************************
Private Sub m_oKeyHandler_HotKey(iKeyCode As Integer, iShift As Integer)
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "m_oKeyHandler_HotKey"
    
    On Error GoTo ERROR_HANDLER
    
    ' Ctrl - Shift
    If (iShift And vbCtrlMask) <> 0 And (iShift And vbShiftMask) <> 0 And (iShift And vbAltMask) = 0 Then
        Select Case iKeyCode
            Case vbKeyM
                DocumentModuleEntry
                
            Case vbKeyP
                DocumentProcedureEntry
                
            Case vbKeyS
                InsertTimeStampEntry
                
            Case vbKeyT
                TemplateMenuEntry
                
            Case vbKeyF4
                CloseWindows
                
            Case vbKeyBack
                ClearImmediate
                
            Case vbKeyV
                CleanUnusedVariablesProcedure
                
            Case vbKeyI
                SmartIndentProcedure
                
            Case vbKeyZ
                ShowZOrderWindow
        End Select
    End If
    
    ' Ctrl - Alt
    If (iShift And vbCtrlMask) <> 0 And (iShift And vbShiftMask) = 0 And (iShift And vbAltMask) <> 0 Then
        Select Case iKeyCode
            Case vbKeyF1
                InsertErrorHandler1
                
            Case vbKeyF2
                InsertErrorHandler2
                
            Case vbKeyH
                ErrorAllEntry
                
            Case vbKeyV
                DimAllEntry
                
            Case vbKeyT
                ShowTabWindow
                
            Case vbKeyI
                SmartIndentAllProcedures
                
            Case vbKeyS
                CalculateStats
        End Select
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
' NonModalApp (PROPERTY GET)
'*******************************************************************************
Public Property Get NonModalApp() As Boolean
    NonModalApp = True
End Property

'*******************************************************************************
' ProcedureList (SUB)
'*******************************************************************************
Private Sub ProcedureList()
    On Error Resume Next
    
    If InRunMode(m_VBInst) Then
        Exit Sub
    End If

    If m_VBInst.ActiveVBProject Is Nothing Then
        Exit Sub
    End If
    
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
        Exit Sub
    End If
    
    m_ProcedureListWindow.Visible = True
    
    m_docProcedureList.DocumentBeforeProcedure = m_bDocumentBeforeProcedure
    
    m_docProcedureList.RefreshList
End Sub

'*******************************************************************************
' GetVarNames (FUNCTION)
'*******************************************************************************
Private Function GetVarNames(sInput() As String) As String()
    Dim sVariables()    As String
    Dim lVarCount       As Long
    Dim sTemp           As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "GetVarNames"
    
    On Error GoTo ERROR_HANDLER
    
    ReDim sVariables(LBound(sInput) To UBound(sInput))
    For lVarCount = LBound(sInput) To UBound(sInput)
        sTemp = Trim(sInput(lVarCount))
        If InStr(sTemp, "(") > 0 Then
            sTemp = Left(sTemp, InStr(sTemp, "(") - 1)
        ElseIf InStr(sTemp, " ") > 0 Then
            sTemp = Left(sTemp, InStr(sTemp, " ") - 1)
        End If
        sVariables(lVarCount) = sTemp
    Next
    
    GetVarNames = sVariables
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
' SplitDeclaration (FUNCTION)
'*******************************************************************************
Private Function SplitDeclaration(ByVal sInput As String) As String()
    Dim lCount          As Long
    Dim sOutput()       As String
    Dim lCharCount      As Long
    Dim sVar            As String
    Dim sTestChar       As String
    Dim lLength         As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim bFinished       As Boolean
    
    Const PROCEDURE_NAME As String = "SplitDeclaration"
    
    On Error GoTo ERROR_HANDLER
    
    lCount = 0
    lCharCount = 0
    sInput = Trim(sInput)
    ReDim Preserve sOutput(lCount)
    lLength = Len(sInput)
    bFinished = False
    
    If sInput <> "" Then
        sVar = ""
        
        Do
            lCharCount = lCharCount + 1
            sTestChar = Mid(sInput, lCharCount, 1)
            
            ' Array
            If sTestChar = "(" Then
                Do
                    sVar = sVar & sTestChar
                    lCharCount = lCharCount + 1
                    sTestChar = Mid(sInput, lCharCount, 1)
                Loop Until lCharCount = lLength Or sTestChar = ")"
                sVar = sVar & sTestChar
                
            ' Var separated with commas
            ElseIf sTestChar = "," Then
                ReDim Preserve sOutput(lCount)
                sOutput(lCount) = Trim(sVar)
                sVar = ""
                lCount = lCount + 1
                
            Else
                sVar = sVar & sTestChar
            End If
        Loop Until lCharCount = lLength Or bFinished
        
        ReDim Preserve sOutput(lCount)
        sOutput(lCount) = Trim(sVar)
    End If
    
    SplitDeclaration = sOutput
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
' ReplaceHeaderTags (SUB)
'*******************************************************************************
Private Sub ReplaceHeaderTags(sSource As String, _
                              ByVal eptType As enumProcedureType, _
                              ByVal sDeclaration As String, _
                              ByVal oMember As Member)
    Dim sTemp                   As String
    Dim lLocn                   As Long
    Dim udtParamsArray()        As udtParameters
    Dim lParamCount             As Long
    Dim lCount                  As Long
    Dim lMaxInOut               As Long
    Dim lMaxName                As Long
    Dim lMaxType                As Long
    Dim lErrNumber              As Long
    Dim sErrSource              As String
    Dim sErrDescription         As String
    
    Const PROCEDURE_NAME As String = "ReplaceHeaderTags"
    
    On Error GoTo ERROR_HANDLER
                                  
    ' Make sure we have at least one % otherwise it is not worth searching for tags
    If InStr(sSource, TAG_QUOTE) > 0 Then
    
        ' Need to find out what the return type is from the declaration
        If InStr(sSource, TAG_RETVAL) > 0 Then
            If eptType <> PTFunction Then
                ReplaceTag sSource, TAG_RETVAL, "N\A"
            Else
                ' Beware of functions that return arrays
                sDeclaration = Trim(sDeclaration)
                If Right(sDeclaration, 2) = "()" Then
                    If InStr(sDeclaration, " As ") Then
                        ' Find second last ")"
                        lLocn = InStrRev(sDeclaration, ")", Len(sDeclaration) - 1)
                    Else
                        lLocn = Len(sDeclaration)
                    End If
                Else
                    ' Find last ")"
                    lLocn = InStrRev(sDeclaration, ")")
                End If
                
                sTemp = ""
                If lLocn > 1 Then
                    ' Check for " As " and determine type
                    sTemp = Mid(sDeclaration, lLocn + 1)
                    If InStr(sTemp, " As ") > 0 Then
                        sTemp = Trim(Mid(sTemp, InStr(sTemp, " As ") + 3))
                    Else
                        sTemp = "Variant"
                    End If
                End If
                
                If sTemp = "" Then
                    ReplaceTag sSource, TAG_RETVAL, "N\A"
                Else
                    ReplaceTag sSource, TAG_RETVAL, sTemp & " - "
                End If
            End If
        End If
    
        If InStr(sSource, TAG_PARAMS) > 0 Then
            If eptType = PTUnknown Then
                ReplaceTag sSource, TAG_PARAMS, "Unknown"
            Else
                ' Get output params
                GetParams sDeclaration, oMember.Name, eptType, udtParamsArray, lParamCount
                
                ' Loop through parameters, writing output
                If lParamCount > 0 Then
                    sTemp = ""
                    
                    ' Get max lengths for alignment
                    lMaxInOut = 0
                    lMaxName = 0
                    lMaxType = 0
                    For lCount = 0 To lParamCount - 1
                        If Len(udtParamsArray(lCount).InOutBoth) > lMaxInOut Then
                            lMaxInOut = Len(udtParamsArray(lCount).InOutBoth)
                        End If
                        If Len(udtParamsArray(lCount).ParameterName) > lMaxName Then
                            lMaxName = Len(udtParamsArray(lCount).ParameterName)
                        End If
                        If Len(udtParamsArray(lCount).ParameterType) > lMaxType Then
                            lMaxType = Len(udtParamsArray(lCount).ParameterType)
                        End If
                    Next lCount
    
                    For lCount = 0 To lParamCount - 1
                        If sTemp <> "" Then
                            sTemp = sTemp & vbCrLf
                        End If
                        
                        Select Case m_sParamFormat
                            Case "1"
                                sTemp = sTemp & Replace(sSource, TAG_PARAMS, _
                                    udtParamsArray(lCount).InOutBoth & _
                                    String(lMaxInOut - Len(udtParamsArray(lCount).InOutBoth), " ") & _
                                    " - " & udtParamsArray(lCount).ParameterName & _
                                    String(lMaxName - Len(udtParamsArray(lCount).ParameterName), " ") & _
                                    " - " & udtParamsArray(lCount).ParameterType) & _
                                    String(lMaxType - Len(udtParamsArray(lCount).ParameterType), " ") & " - "
                            
                            Case Else
                                sTemp = sTemp & Replace(sSource, TAG_PARAMS, _
                                    udtParamsArray(lCount).InOutBoth & _
                                    String(lMaxInOut - Len(udtParamsArray(lCount).InOutBoth), " ") & _
                                    " - " & udtParamsArray(lCount).ParameterName & _
                                    String(lMaxName - Len(udtParamsArray(lCount).ParameterName), " ") & _
                                    " - " & udtParamsArray(lCount).ParameterType) & _
                                    String(lMaxType - Len(udtParamsArray(lCount).ParameterType), " ") & " - "
                        End Select
                    Next
                    sSource = sTemp
                Else
                    ReplaceTag sSource, TAG_PARAMS, "None"
                End If
            End If
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
' ResetOptions (SUB)
'*******************************************************************************
Private Sub ResetOptions()
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim sBlankLines     As String
    
    Const PROCEDURE_NAME As String = "ResetOptions"
    
    On Error GoTo ERROR_HANDLER
    
    GetPaths m_sTemplatePath, m_sBoilerplatePath
    
    GetUserTags m_colUserTags
    
    GetOtherTags m_sAuthor, m_sCompany, m_sInitials, m_sTimeFormat, m_sDateFormat, m_bDocumentBeforeProcedure, m_bCloseActive
    
    m_sNormalErrors = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_NORMAL_ERRORS, DEF_REG_NORMAL_ERRORS)
    m_sEventErrors = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_EVENT_ERRORS, DEF_REG_EVENT_ERRORS)
    m_sPropertyErrors = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_PROPERTY_ERRORS, DEF_REG_PROPERTY_ERRORS)
    
    m_sIndent = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_INDENT, DEF_REG_INDENT)
    
    m_sIndentDim = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_INDENT_DIM, DEF_REG_INDENT_DIM)
    
    m_sTabDisplay = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_TAB_DISPLAY, DEF_REG_TAB_DISPLAY)
    
    m_sZOrderDisplay = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_ZORDER_DISPLAY, DEF_REG_ZORDER_DISPLAY)
    
    sBlankLines = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_BLANK_LINES, DEF_REG_BLANK_LINES)
    
    m_lBlankLines = CLng(Val(sBlankLines))
    If m_lBlankLines < 1 Then
        m_lBlankLines = 1
    End If

    m_sParamFormat = GetSetting(REG_APP_NAME, REG_SETTINGS, REG_PARAMFORMAT, DEF_REG_PARAMFORMAT)
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
' ShowTabWindow (SUB)
'*******************************************************************************
Private Sub ShowTabWindow()
    On Error Resume Next
    
    m_TabWindow.Visible = True
    m_docTab.RefreshList 3
End Sub

'*******************************************************************************
' ShowZOrderWindow (SUB)
'*******************************************************************************
Private Sub ShowZOrderWindow()
    On Error Resume Next
    
    m_ZOrderWindow.Visible = True
    m_docZOrder.RefreshList 3
End Sub

'*******************************************************************************
' SmartIndentAllProcedures (SUB)
'*******************************************************************************
Private Sub SmartIndentAllProcedures()
    Dim lReply          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim sDeclaration    As String
    Dim lLineNumber     As Long
    Dim bProcedure      As Boolean
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim bDirty          As Boolean
    Dim oProject        As VBProject
    Dim lEndLine        As Long
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "SmartIndentAllProcedures"
    
    On Error GoTo ERROR_HANDLER

    ' Make sure we have an active code window
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "A code pane must be selected first", _
            vbCritical, App.ProductName
    Else
        bDirty = False
        For Each oProject In m_VBInst.VBProjects
            If oProject.IsDirty Then
                bDirty = True
                Exit For
            End If
        Next
        
        If bDirty Then
            lReply = MsgBox("It is recommended that you save the project(s) before" & _
                " smart indenting all procedures, do you want to continue anyway?", _
                vbQuestion + vbYesNo, App.ProductName)
            If lReply = vbNo Then
                Exit Sub
            End If
        End If
        
        ' Get confirmation from user
        lReply = MsgBox("Smart indent all procedures?", vbYesNo, App.ProductName)
        If lReply = vbYes Then
            With m_VBInst.ActiveCodePane.CodeModule
                ' Loop through procedures and properties
                For Each oMember In .Members
                
                    ' Document the method
                    If oMember.Type = vbext_mt_Method Then
                    
                        ' Make sure it is not a 'Declare'
                        bProcedure = False
                        If .ProcBodyLine(oMember.Name, vbext_pk_Proc) = 1 And _
                            .ProcStartLine(oMember.Name, vbext_pk_Proc) = 1 Then
                            
                            sDeclaration = " " & GetDeclarationLine(m_VBInst, 1, lEndLine)
                            
                            If (InStr(sDeclaration, " Sub " & oMember.Name) > 0 _
                                Or InStr(sDeclaration, " Function " & oMember.Name) > 0) _
                                And InStr(sDeclaration, " Declare ") = 0 Then
                                    bProcedure = True
                            End If
                        Else
                            bProcedure = True
                        End If
                        
                        If bProcedure Then
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Proc)
                            GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
            
                            If lTopLine <> 0 Then
                                SmartIndentThisProcedure lTopLine, oMember, eptType
                            End If
                        End If
                        
                    ' Document Let, Set and Get part of property if they exist
                    ElseIf oMember.Type = vbext_mt_Property Then
'                        If m_sPropertyErrors = "1" Then
                            On Error Resume Next
                            Err.Clear
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Get)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If lTopLine <> 0 Then
                                    SmartIndentThisProcedure lTopLine, oMember, eptType
                                End If
                            Else
                                Err.Clear
                            End If
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Set)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If lTopLine <> 0 Then
                                    SmartIndentThisProcedure lTopLine, oMember, eptType
                                End If
                            Else
                                Err.Clear
                            End If
                            
                            lLineNumber = .ProcBodyLine(oMember.Name, vbext_pk_Let)
                            If Err.Number = 0 Then
                                GetProcedureDetails m_VBInst, lLineNumber, oMember, eptType, lTopLine, lEndDecLine
                                If lTopLine <> 0 Then
                                    SmartIndentThisProcedure lTopLine, oMember, eptType
                                End If
                            Else
                                Err.Clear
                            End If
                            On Error GoTo ERROR_HANDLER
'                        End If
                    End If
                Next
                
                SmartIndentDeclaration
            End With
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
' SmartIndentDeclaration (SUB)
'*******************************************************************************
Private Sub SmartIndentDeclaration()
    Dim lTop            As Long
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lStart          As Long
    Dim lLinesCount     As Long
    Dim bFoundEnd       As Boolean
    Dim dIndentChars    As Double
    Dim lIndentChars    As Long
    Dim lIndent         As Long
    Dim sNewLine        As String
    Dim sOldLine        As String
    Dim sCleaned        As String
    Dim sNextLine       As String
    Dim bCont           As Boolean
    Dim lCurrent        As Long
    Dim sWords()        As String
    Dim bDimCont        As Boolean
    
    Const PROCEDURE_NAME As String = "SmartIndentDeclaration"
    
    Const PREFIX_WORDS  As String = " Private Public Friend "
    Const END_WORDS     As String = " End #Else #ElseIf #End "
    Const START_WORDS   As String = " #Else Enum Type "
    Const IF_WORDS      As String = " #If #ElseIf "
    Const DIM_WORDS     As String = " Dim Static Const "

    On Error GoTo ERROR_HANDLER
    
    dIndentChars = Val(m_sIndent)
    If dIndentChars < 1 Or dIndentChars > 8 Then
        lIndentChars = 4
    Else
        lIndentChars = dIndentChars
    End If
    
    lTop = 1
    lLinesCount = m_VBInst.ActiveCodePane.CodeModule.CountOfDeclarationLines
    
    With m_VBInst.ActiveCodePane.CodeModule
        lStart = lTop
        bFoundEnd = False
        lIndent = 0
        bCont = False
        Do
            lCurrent = lStart
            sOldLine = .Lines(lStart, 1)
            sNewLine = Trim(sOldLine)
            sCleaned = CleanUpLine(sOldLine)
            lStart = lStart + 1
                
            If sNewLine = "" Then
                .ReplaceLine lCurrent, ""
            Else
                sNextLine = ""
                
                If bCont Then
                    If bDimCont Then
                        .ReplaceLine lCurrent, String(lIndentChars, " ") & sNewLine
                        If Right(sCleaned, 2) = " _" Then
                            bDimCont = True
                        Else
                            bDimCont = False
                        End If
                    Else
                        .ReplaceLine lCurrent, String((lIndent + 1) * lIndentChars, " ") & sNewLine
                    End If
                    
                ElseIf sCleaned = "" Then
                    .ReplaceLine lCurrent, String(lIndent * lIndentChars, " ") & sNewLine
                    
                Else
                    sWords = Split(sCleaned, " ")
                    
                    If InStr(END_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                        lIndent = lIndent - 1
                        
                    ElseIf Trim(sWords(0)) = "End" Then
                        If Len(sCleaned) > 3 Then
                            If Trim(sWords(1)) = "Select" Then
                                lIndent = lIndent - 2
                            Else
                                lIndent = lIndent - 1
                            End If
                        End If
                    End If
                    
                    If lIndent < 0 Then
                        lIndent = 0
                    End If
                    
                    .ReplaceLine lCurrent, String(lIndent * lIndentChars, " ") & sNewLine
                    
                    If InStr(START_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                        lIndent = lIndent + 1
                        
                    ElseIf InStr(PREFIX_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                        If InStr(START_WORDS, " " & Trim(sWords(1)) & " ") > 0 Then
                            lIndent = lIndent + 1
                        End If
                        
                    ElseIf InStr(IF_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                        If Right(sCleaned, 5) = " Then" Or Right(sCleaned, 2) = " _" Then
                            lIndent = lIndent + 1
                        End If
                    End If
                End If
                
                If Right(sCleaned, 2) = " _" Then
                    bCont = True
                Else
                    bCont = False
                End If
            End If
        Loop Until lStart > lLinesCount
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
' SmartIndentProcedure (SUB)
'*******************************************************************************
Private Sub SmartIndentProcedure()
    Dim lStartLine      As Long
    Dim lDummy          As Long
    Dim oMember         As Member
    Dim lTopLine        As Long
    Dim eptType         As enumProcedureType
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lEndDecLine     As Long
    
    Const PROCEDURE_NAME As String = "SmartIndentProcedure"
    
    On Error GoTo ERROR_HANDLER
    
    If m_VBInst.ActiveCodePane Is Nothing Then
        MsgBox "Select a code pane first", _
            vbCritical, App.ProductName
    Else
        ' Find out where the cursor is at the moment
        m_VBInst.ActiveCodePane.GetSelection lStartLine, lDummy, lDummy, lDummy
        
        With m_VBInst.ActiveCodePane.CodeModule
            GetProcedureDetails m_VBInst, lStartLine, oMember, eptType, lTopLine, lEndDecLine
            
            If lTopLine <> 0 Then
                SmartIndentThisProcedure lTopLine, oMember, eptType
            Else
                SmartIndentDeclaration
'                MsgBox "You must be within a procedure in order to perform the smart indent", _
'                    vbCritical, App.ProductName
            End If
        End With
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
' SmartIndentThisProcedure (SUB)
'*******************************************************************************
Private Sub SmartIndentThisProcedure(ByVal lLineNumber As Long, _
                                     ByVal oMember As Member, _
                                     ByVal eptType As enumProcedureType)
    Dim lTop            As Long
    Dim sProcName       As String
    Dim sEndTest        As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim lStart          As Long
    Dim lLinesCount     As Long
    Dim eProcKind       As vbext_ProcKind
    Dim bFoundEnd       As Boolean
    Dim dIndentChars    As Double
    Dim lIndentChars    As Long
    Dim lIndent         As Long
    Dim sNewLine        As String
    Dim sOldLine        As String
    Dim sCleaned        As String
    Dim sNextLine       As String
    Dim bCont           As Boolean
    Dim lCurrent        As Long
    Dim sWords()        As String
    Dim bDimCont        As Boolean
    Dim sNumber         As String
    Dim lCheck          As Long
    Dim sWhite          As String
    
    Const PROCEDURE_NAME As String = "SmartIndentThisProcedure"
    
    Const END_WORDS     As String = " Else ElseIf Case Next Loop Wend #Else #ElseIf #End "
    Const START_WORDS   As String = " Else With For Do Case While #Else Enum Type "
    Const IF_WORDS      As String = " If ElseIf #If #ElseIf "
    Const DIM_WORDS     As String = " Dim Static Const "

    On Error GoTo ERROR_HANDLER
    
    dIndentChars = Val(m_sIndent)
    If dIndentChars < 1 Or dIndentChars > 8 Then
        lIndentChars = 4
    Else
        lIndentChars = dIndentChars
    End If
    
    Select Case eptType
        Case PTFunction
            sEndTest = "End Function"
            eProcKind = vbext_pk_Proc
            
        Case PTGet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Get
        
        Case PTLet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Let
        
        Case PTSet
            sEndTest = "End Property"
            eProcKind = vbext_pk_Set
        
        Case PTSub
            sEndTest = "End Sub"
            eProcKind = vbext_pk_Proc
        
        Case Else
            Exit Sub
    End Select
    
    lTop = lLineNumber
    With m_VBInst.ActiveCodePane.CodeModule
        sProcName = .ProcOfLine(lLineNumber, eProcKind)
        lLinesCount = .CountOfLines
        
        ' Find first line after declaration by checking continuation characters
        bFoundEnd = False
        bCont = False
        Do
            sOldLine = .Lines(lTop, 1)
            If bCont Then
                sNumber = ""
                sNewLine = Trim(sOldLine)
            Else
                sNumber = GetLineNumber(Trim(sOldLine))
                sNewLine = TrimLine(sOldLine)
            End If
            sCleaned = CleanUpLine(sNewLine)
            
            If lTop = lLineNumber Then
                If sNumber <> "" Then
                    .ReplaceLine lTop, sNumber & " " & sNewLine
                Else
                    .ReplaceLine lTop, sNewLine
                End If
            Else
                If sNewLine <> "" Then
                    sWhite = String(lIndentChars, " ")
                    
                    If m_sIndentDim = "2" Then
                        If Left(sCleaned, 4) = "Dim " Or _
                            Left(sCleaned, 6) = "Const " Or _
                            Left(sCleaned, 5) = "Type " Or _
                            Left(sCleaned, 5) = "Enum " Or _
                            Left(sCleaned, 7) = "Static " Then
                            
                            If sNumber <> "" Then
                                .ReplaceLine lTop, sNumber & " " & sNewLine
                            Else
                                .ReplaceLine lTop, sNewLine
                            End If
                        Else
                            If Len(sNumber) >= Len(sWhite) Then
                                .ReplaceLine lTop, sNumber & " " & sNewLine
                            Else
                                .ReplaceLine lTop, sNumber & String(Len(sWhite) - Len(sNumber), " ") & sNewLine
                            End If
                        End If
                    Else
                        If Len(sNumber) >= Len(sWhite) Then
                            .ReplaceLine lTop, sNumber & " " & sNewLine
                        Else
                            .ReplaceLine lTop, sNumber & String(Len(sWhite) - Len(sNumber), " ") & sNewLine
                        End If
                    End If
                Else
                    If sNumber <> "" Then
                        .ReplaceLine lTop, sNumber
                    Else
                        .ReplaceLine lTop, ""
                    End If
                End If
            End If
            
            lTop = lTop + 1
            
            If sCleaned = sEndTest Then
                bFoundEnd = True
            ElseIf sCleaned = "End _" Then
                sNextLine = CleanUpLine(.Lines(lTop, 1))
                If sNextLine = Mid(sEndTest, 5) Then
                    bFoundEnd = True
                End If
            End If
            
            If Right(sCleaned, 2) = " _" Then
                bCont = True
            Else
                bCont = False
            End If
        Loop Until (Not bCont) Or lTop > lLinesCount Or bFoundEnd
        
        If lTop > lLinesCount And Not bFoundEnd Then
            Exit Sub
        End If

        ' Loop through the procedure
        lStart = lTop
        bFoundEnd = False
        lIndent = 1
        bCont = False
        Do
            lCurrent = lStart
            
            sOldLine = .Lines(lStart, 1)
            If bCont Then
                sNumber = ""
                sNewLine = Trim(sOldLine)
            Else
                sNumber = GetLineNumber(Trim(sOldLine))
                sNewLine = TrimLine(sOldLine)
            End If
            sCleaned = CleanUpLine(sNewLine)
            
            lStart = lStart + 1
            
            If sNewLine = "" Then
                If sNumber = "" Then
                    .ReplaceLine lCurrent, ""
                Else
                    .ReplaceLine lCurrent, sNumber
                End If
            Else
                sNextLine = ""
                
                If sCleaned = sEndTest Then
                    bFoundEnd = True
                ElseIf sCleaned = "End _" Then
                    sNextLine = CleanUpLine(.Lines(lCurrent + 1, 1))
                    If sNextLine = Mid(sEndTest, 5) Then
                        bFoundEnd = True
                        Exit Do
                    End If
                End If
                
                If bFoundEnd Then
                    If sNumber = "" Then
                        .ReplaceLine lCurrent, sNewLine
                    Else
                        .ReplaceLine lCurrent, sNumber & " " & sNewLine
                    End If
                    If sNextLine <> "" Then
                        sWhite = String(lIndentChars, " ")
                        .ReplaceLine lCurrent + 1, sWhite & LTrim(.Lines(lCurrent + 1, 1))
                    End If
                Else
                    If bCont Then
                        If bDimCont Then
                            sWhite = String(lIndentChars, " ")
                            .ReplaceLine lCurrent, sWhite & sNewLine
                            If Right(sCleaned, 2) = " _" Then
                                bDimCont = True
                            Else
                                bDimCont = False
                            End If
                        Else
                            sWhite = String((lIndent + 1) * lIndentChars, " ")
                            .ReplaceLine lCurrent, sWhite & sNewLine
                        End If
                    ElseIf sCleaned = "" Then
                        sWhite = String(lIndent * lIndentChars, " ")
                        If Len(sNumber) >= Len(sWhite) Then
                            .ReplaceLine lCurrent, sNumber & " " & sNewLine
                        Else
                            .ReplaceLine lCurrent, sNumber & String(Len(sWhite) - Len(sNumber), " ") & sNewLine
                        End If
                    Else
                        sWords = Split(sCleaned, " ")
                        
                        If InStr(END_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                            lIndent = lIndent - 1
                            
                        ElseIf Trim(sWords(0)) = "End" Then
                            If Len(sCleaned) > 3 Then
                                If Trim(sWords(1)) = "Select" Then
                                    lIndent = lIndent - 2
                                Else
                                    lIndent = lIndent - 1
                                End If
                            End If
                        End If
                        
                        If lIndent < 0 Then
                            lIndent = 0
                        End If
                        
                        If m_sIndentDim = "2" Then
                            If InStr(DIM_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                                If sNumber <> "" Then
                                    .ReplaceLine lCurrent, sNumber & " " & sNewLine
                                Else
                                    .ReplaceLine lCurrent, sNewLine
                                End If
                                If Right(sCleaned, 2) = " _" Then
                                    bDimCont = True
                                Else
                                    bDimCont = False
                                End If
                            Else
                                sWhite = String(lIndent * lIndentChars, " ")
                                If Len(sNumber) >= Len(sWhite) Then
                                    .ReplaceLine lCurrent, sNumber & " " & sNewLine
                                Else
                                    .ReplaceLine lCurrent, sNumber & String(Len(sWhite) - Len(sNumber), " ") & sNewLine
                                End If
                                bDimCont = False
                            End If
                        Else
                            sWhite = String(lIndent * lIndentChars, " ")
                            If Len(sNumber) >= Len(sWhite) Then
                                .ReplaceLine lCurrent, sNumber & " " & sNewLine
                            Else
                                .ReplaceLine lCurrent, sNumber & String(Len(sWhite) - Len(sNumber), " ") & sNewLine
                            End If
                        End If
                        
                        If InStr(START_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                            lIndent = lIndent + 1
                            
                        ElseIf Trim(sWords(0)) = "Select" Then
                            lIndent = lIndent + 2
                            
                        ElseIf InStr(IF_WORDS, " " & Trim(sWords(0)) & " ") > 0 Then
                            If Right(sCleaned, 5) = " Then" Or Right(sCleaned, 2) = " _" Then
                                lIndent = lIndent + 1
                            End If
                        End If
                    End If
                    
                    If Right(sCleaned, 2) = " _" Then
                        bCont = True
                    Else
                        bCont = False
                    End If
                End If
            End If
        Loop Until bFoundEnd Or lStart > lLinesCount
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
' TemplateMenuEntry (SUB)
'*******************************************************************************
Private Sub TemplateMenuEntry()
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    
    Const PROCEDURE_NAME As String = "TemplateMenuEntry"
    
    On Error GoTo ERROR_HANDLER

    m_TemplateWindow.Visible = True
    m_docTemplates.RefreshList

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
