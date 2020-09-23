Attribute VB_Name = "mIDE"
Option Explicit

Public Type typeCodeSelection
    StartLine As Long
    StartColumn As Long
    EndLine As Long
    EndColumn As Long
End Type

Public Type typeScreenArea
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    
End Type

Public Const sProcHeader = "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
Public Const sHardTab = "    "

Public Function vbComponentTypeText(ByRef vbComp As VBComponent)
    Dim sRet As String
    Select Case True
        Case vbComp.Type = vbext_ct_ClassModule
            sRet = "Class"
        Case vbComp.Type = vbext_ct_VBForm
            sRet = "Form"
        Case vbComp.Type = vbext_ct_VBMDIForm
            sRet = "MDIForm"
        Case Else
    End Select
    vbComponentTypeText = sRet
End Function


Public Function vbComponentType(ByVal eCompType As vbext_ComponentType) As String
    Dim sRet As String
    Select Case True
        Case eCompType = vbext_ct_ActiveXDesigner
            sRet = "ActiveX Designers"
        Case eCompType = vbext_ct_ClassModule
            sRet = "Class Modules"
        Case eCompType = vbext_ct_DocObject
            sRet = "User Documents"
        Case eCompType = vbext_ct_MSForm
            sRet = "MSForms"
        Case eCompType = vbext_ct_PropPage
            sRet = "Property Pages"
        Case eCompType = vbext_ct_RelatedDocument
            sRet = "Related Documents"
        Case eCompType = vbext_ct_ResFile
            sRet = "Resource File"
        Case eCompType = vbext_ct_StdModule
            sRet = "Modules"
        Case eCompType = vbext_ct_UserControl
            sRet = "User Controls"
        Case eCompType = vbext_ct_VBForm
            sRet = "Forms"
        Case eCompType = vbext_ct_VBMDIForm
            sRet = "MDI Forms"
        Case Else
            sRet = "UNKNOWN"
    End Select
    vbComponentType = sRet
End Function

Public Function GetTrueProcTopLine(ByRef vArray As Variant) As Long
    
    Dim l As Long
    Dim sLine As String
    For l = 0 To UBound(vArray)
        sLine = Trim(UCase(vArray(l)))
        If _
            ts.sLeftIs(sLine, "PUBLIC FUNCTION ") Or _
            ts.sLeftIs(sLine, "PUBLIC SUB ") Or _
            ts.sLeftIs(sLine, "PRIVATE FUNCTION ") Or _
            ts.sLeftIs(sLine, "PRIVATE SUB ") Or _
            ts.sLeftIs(sLine, "FUNCTION ") Or _
            ts.sLeftIs(sLine, "SUB ") Then
            
            GetTrueProcTopLine = l
            Exit Function
        End If
    Next l
        
End Function

Public Function GetCommentText(ByRef vArray As Variant, ByVal lProcTrueTopLine As Long, ByVal sProcName As String)
    
    Dim lCurrLine As Long
    Dim sComments As String
    Dim sCurrLine As String
    lCurrLine = 0
    Do While vArray(lCurrLine) = ""
        lCurrLine = lCurrLine + 1
    Loop
    Do While lCurrLine < lProcTrueTopLine
        sCurrLine = UncommentLine(vArray(lCurrLine))
        Select Case True
            Case sCurrLine = sProcHeader
            Case sCurrLine = " " & sProcName
            Case Else
                sComments = sComments & sCurrLine & vbCrLf
        End Select
        lCurrLine = lCurrLine + 1
    Loop
    GetCommentText = sComments
    
End Function


Public Function UncommentLine(ByVal sLine As String) As String
    Dim sBegin As String
    If Left(LTrim(sLine), 5) = "'" & sHardTab Then
        sBegin = Left(sLine, InStr(sLine, "'" & sHardTab) - 1)
        sLine = LTrim(sLine)
        Do While Left(sLine, 5) = "'" & sHardTab
            sLine = Mid(sLine, 6)
        Loop
        sLine = sBegin & sLine
    End If
    If Left(LTrim(sLine), 1) = "'" Then
        sBegin = Left(sLine, InStr(sLine, "'") - 1)
        sLine = LTrim(sLine)
        Do While Left(sLine, 1) = "'"
            sLine = Mid(sLine, 2)
        Loop
        UncommentLine = sBegin & sLine
    Else
        UncommentLine = sLine
    End If
End Function

Public Function ideGetVBArea(ByRef vb As VBide.VBE) As typeScreenArea
    Dim ret As typeScreenArea
    ret.Left = -1
    ret.Top = -1
    ret.Height = -1
    ret.Width = -1
    If vb.DisplayModel = vbext_dm_SDI Then
        ideGetVBArea = ret
        Exit Function
    End If
    With vb.MainWindow
        If .WindowState = vbext_ws_Normal Then
            ret.Left = .Left
            ret.Top = .Top
        Else
            ret.Left = 0
            ret.Top = 0
        End If
        ret.Width = .Width
        ret.Height = .Height
    End With
    ideGetVBArea = ret
End Function

Public Function ideGetWorkAreaTwips( _
                                                        ByRef vb As VBide.VBE) _
                        As typeScreenArea
    
    Dim ret As typeScreenArea
    ret = ideGetWorkAreaPixels(vb)
    ret.Height = twipsY(ret.Height)
    ret.Top = twipsY(ret.Top)
    ret.Left = twipsX(ret.Left)
    ret.Width = twipsX(ret.Width)
    ideGetWorkAreaTwips = ret
    
End Function

Public Function ideGetWorkAreaPixels(ByRef vb As VBide.VBE) As typeScreenArea
    
    Dim ret As typeScreenArea
    If vb.DisplayModel = vbext_dm_SDI Then
        Exit Function
    End If
    
    ' Provide for Window Caption
    Dim vOffset As Long
    vOffset = ts.sysMetrics(SM_CYBORDER) + ts.sysMetrics(SM_CYCAPTION)
    
    Dim cWindow As New clsWindow
    Set cWindow = New clsWindow
    cWindow.hwnd = vb.MainWindow.hwnd
    cWindow.RefreshChildren False
    
    Dim l As Long
    For l = 1 To cWindow.Children.Count
        Debug.Print cWindow.Children(l).sClassName
        If UCase(cWindow.Children(l).sClassName) = "MDICLIENT" Then
            ret.Left = cWindow.Children(l).Left
            ret.Top = cWindow.Children(l).Top
            ret.Height = cWindow.Children(l).Bottom - cWindow.Children(l).Top
            ret.Width = cWindow.Children(l).Right - cWindow.Children(l).Left
            Exit For
        End If
    Next l
    
    Set cWindow = Nothing
    ideGetWorkAreaPixels = ret
    
End Function

Public Function ideFontName() As String
    
    Dim cReg As New cRegistry
    cReg.ClassKey = HKEY_CURRENT_USER
    cReg.SectionKey = "Software\Microsoft\VBA\Microsoft Visual Basic"
    cReg.ValueKey = "FontFace"
    If cReg.Value <> "" Then
        ideFontName = cReg.Value
    Else
        ideFontName = "Courier New"
    End If
    Set cReg = Nothing
    
End Function

Public Function ideFontSize() As Integer
    
    Dim cReg As New cRegistry
    cReg.ClassKey = HKEY_CURRENT_USER
    cReg.SectionKey = "Software\Microsoft\VBA\Microsoft Visual Basic"
    cReg.ValueKey = "FontHeight"
    If cReg.Value <> 0 Then
        ideFontSize = cReg.Value
    Else
        ideFontSize = 10
    End If
    Set cReg = Nothing
    
End Function



Public Function formFillWorkArea( _
                                ByRef frm As Form) _
                As Boolean
            
    If gVB.DisplayModel = vbext_dm_MDI Then
        Dim area As typeScreenArea
        area = mIDE.ideGetWorkAreaTwips(gVB)
        frm.Move area.Left, area.Top, area.Width, area.Height
    End If
            
End Function

