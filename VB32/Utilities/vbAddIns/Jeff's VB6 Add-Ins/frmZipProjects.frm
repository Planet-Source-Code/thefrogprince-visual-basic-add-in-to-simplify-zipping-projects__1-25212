VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmZipProjects 
   Caption         =   "Zip Up Project(s)"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "frmZipProjects.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIncludeCompatible 
      Caption         =   "Include Compatible?"
      Height          =   225
      Left            =   6390
      TabIndex        =   7
      ToolTipText     =   "Necessary In Certain Cases In Windows 2000 To Avoid GPF"
      Top             =   300
      Width           =   1965
   End
   Begin VB.CheckBox chkIncludeEXE 
      Caption         =   "Include Compiled?"
      Height          =   225
      Left            =   4620
      TabIndex        =   6
      ToolTipText     =   "Necessary In Certain Cases In Windows 2000 To Avoid GPF"
      Top             =   300
      Width           =   1725
   End
   Begin VB.CheckBox chkUseWinZip 
      Caption         =   "Use External Winzip?"
      Height          =   225
      Left            =   2580
      TabIndex        =   5
      ToolTipText     =   "Necessary In Certain Cases In Windows 2000 To Avoid GPF"
      Top             =   300
      Width           =   1965
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H8000000F&
      Height          =   5115
      Left            =   4500
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   810
      Width           =   4065
   End
   Begin MSComctlLib.ProgressBar pbMain 
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   5940
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sbMain 
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      Top             =   5940
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   7050
      Top             =   1410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "ZIP"
      DialogTitle     =   "Create A Zip Of Files"
   End
   Begin VB.CommandButton cmdZip 
      Caption         =   "&Zip"
      Height          =   495
      Left            =   90
      TabIndex        =   1
      Top             =   210
      Width           =   2355
   End
   Begin MSComctlLib.TreeView tvProjectExplorer 
      Height          =   4995
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8811
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmZipProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oZip As cZip
Attribute oZip.VB_VarHelpID = -1
Dim colFiles As Collection
Public bLoading As Boolean

Private Sub chkUseWinZip_Click()
    If Me.chkUseWinZip.Value Then
        If Me.sWinZipEXE <> "" Then
            Shell Me.sWinZipEXE, vbNormalFocus
        Else
            MsgBox "You must have installed WinZip's command line interface for this to work since it is not installed with the standard WinZip eval package.", vbOKOnly + vbInformation, "Note"
            Me.chkUseWinZip.Value = 0
        End If
    End If
    'Me.chkZipIndividually.Enabled = bNot(Me.chkUseWinZip.Value)
    
End Sub

Public Function sWinZipEXE() As String
    
    Dim cReg As cRegistry
    Set cReg = New cRegistry
    cReg.ClassKey = HKEY_LOCAL_MACHINE
    cReg.SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winzip32.exe"
    sWinZipEXE = ts.shell16bitFileName(cReg.Value)
    Set cReg = Nothing
    
End Function

Private Sub cmdZip_Click()
    
    Dim bZipIndividually As Boolean
    Dim bEXE As Boolean
    Dim bCompatible As Boolean
    Dim bWinZip As Boolean
    bWinZip = Me.chkUseWinZip.Value
    bZipIndividually = False
    bEXE = Me.chkIncludeEXE.Value
    bCompatible = Me.chkIncludeCompatible.Value
    
    ' First Test To Make Sure The DLL Calls Will Work
    If Not bWinZip Then
        On Error Resume Next
        Set oZip = New cZip
        If Err.Number <> 0 Then
            MsgBox "Error #: " & Err.Number & vbCrLf & "Desc: " & Err.Description & vbCrLf & vbCrLf & "While this VB6 add-in should not require an install, to use the project ZIPping feature, the following dll: " & vbCrLf & vbCrLf & "vbzip10.dll" & vbCrLf & vbCrLf & "MUST be in your system path somewhere (preferably the Windows\System directory). Please locate this file and ensure that it exists in your path and then try again.", vbCritical + vbOKOnly, "Support DLL Not Found"
            Unload Me
            Exit Sub
        End If
        On Error GoTo 0
    Else
        Dim sWinZipEXE As String
        sWinZipEXE = Me.sWinZipEXE
        If sWinZipEXE = "" Then
            MsgBox "In order to use the external WinZip option, you must have WinZip installed on your computer."
            Exit Sub
        End If
        
    End If
    
    ' Get FileName Of New Zip
    If Not bWinZip Then
        Dim sNewZipFile As String
        Dim sNewZipPath As String
        On Error Resume Next
        sNewZipPath = ts.sFileName(gVB.VBProjects(1).Filename, efpFilePath)
        On Error GoTo 0
        Me.dlgCommon.DefaultExt = "ZIP"
        Me.dlgCommon.Filename = sNewZipPath & "BackUp" & Format(Now, "YYYY-MM-DD")
        Me.dlgCommon.Filter = "Zip Files (*.zip)|*.zip|All Files (*.*)|*.*"
        Me.dlgCommon.CancelError = True
        On Error Resume Next
        Me.dlgCommon.ShowSave
        If Err.Number <> 0 Then
            Exit Sub
        End If
        On Error GoTo 0
        sNewZipFile = Me.dlgCommon.Filename
    End If
        
    ' Start Creating Zip File
    Me.txtOutput.Text = ""
    Me.pbMain.Max = Me.tvProjectExplorer.Nodes.Count
    Me.pbMain.Min = 0
    
    ' Prep Zip Class
    If Not bWinZip Then
        oZip.StoreFolderNames = True
        oZip.RecurseSubDirs = False
        oZip.ZipFile = sNewZipFile
        oZip.ClearFileSpecs
    Else
        Set colFiles = New Collection
        
    End If
    
    Dim lCurrNode As Long
    Dim aKeys As Variant
    For lCurrNode = 1 To Me.tvProjectExplorer.Nodes.Count
        DoEvents
        Me.pbMain.Value = lCurrNode
        If Me.tvProjectExplorer.Nodes(lCurrNode).Checked Then
            aKeys = Split(Me.tvProjectExplorer.Nodes(lCurrNode).key, "~")
            If bZipIndividually Then
                oZip.ClearFileSpecs
            End If
            Select Case True
                Case UBound(aKeys) = 0
                    SaveObject gVB.VBProjects(aKeys(0))
                    Me.AddFileSpec gVB.VBProjects(aKeys(0)).Filename
                    If bEXE Then
                        Me.AddFileSpec gVB.VBProjects(aKeys(0)).BuildFileName
                                        
                    End If
                    If bCompatible Then
                        Me.AddFileSpec gVB.VBProjects(aKeys(0)).CompatibleOleServer
                    End If
                Case UBound(aKeys) = 2 And aKeys(1) = "Resource File"
                    SaveObject gVB.VBProjects(aKeys(0)).VBComponents(Val(aKeys(2)))
                    Me.AddFileSpec gVB.VBProjects(aKeys(0)).VBComponents(Val(aKeys(2))).FileNames(1)
                    DoEvents
                    DoEvents
                Case UBound(aKeys) = 2
                    Dim i As Integer
                    Dim vKey As Variant
                    vKey = aKeys(2)
                    If ts.sIs(vKey, esiOnlyNumbers) Then
                        vKey = Val(vKey)
                    End If
                    SaveObject gVB.VBProjects(aKeys(0)).VBComponents(vKey)
                    With gVB.VBProjects(aKeys(0)).VBComponents(vKey)
                        For i = 1 To .fileCount
                            Me.AddFileSpec gVB.VBProjects(aKeys(0)).VBComponents(vKey).FileNames(i)
                        Next i
                    End With
            End Select
            If bZipIndividually Then
                If oZip.FileSpecCount > 0 Then
                    DoEvents
                    DoEvents
                    DoEvents
                    DoEvents
                    oZip.Zip
                End If
                Dim q As Integer
                For q = 1 To 10
                    DoEvents
                Next q
            End If
        End If
    Next lCurrNode
    
    Me.pbMain.Value = 0
    If Not bWinZip Then
        Me.pbMain.Max = (oZip.FileSpecCount * 2) + 1
    Else
        Me.pbMain.Max = colFiles.Count
    End If
    
    Dim oErr As New clsError
    Select Case True
        Case Not bZipIndividually And Not bWinZip
            oZip.Zip oErr
        Case bWinZip

            Dim sParams As String
            Dim l As Long
            Dim lNextTop As Long
'            Do While l < colFiles.Count
            For l = 1 To colFiles.Count
                sParams = sParams & """" & colFiles(l) & """ "
            Next l
            Shell sWinZipEXE & " " & sParams, vbNormalFocus
            For l = 1 To 50
                DoEvents
            Next l
'            Loop
    End Select
    If Not bWinZip Then
        If oErr.Number <> 0 Then
            MsgBox "Bummer!" & vbCrLf & "It did not work for you.  =( This could be because you are missing some the required vbZip32.dll file.  Here's your error message: " & vbCrLf & vbCrLf & "Error #: " & oErr.Number & vbCrLf & "Desc: " & oErr.Description, vbCritical + vbOKOnly, "ZIP Failure"
        Else
            MsgBox "Zip file successfully created.  =)", vbOKOnly + vbInformation, "Success"
        End If
    End If
    Set oErr = Nothing
    
'    Unload Me
    
    Set oZip = Nothing
    
End Sub

Public Function AddFileSpec(ByVal sSpec As String)
    Select Case True
        Case Me.chkUseWinZip.Value
            colFiles.Add sSpec
        Case Else
            oZip.AddFileSpec sSpec
    End Select
End Function

Public Function SaveObject(ByRef oToSave As Object)
    
    ' Disable Auto-Save
    Exit Function
    
    Dim sFileName As String
    On Error Resume Next
    sFileName = oToSave.Filename
    On Error GoTo 0
    If sFileName = "" Then
        sFileName = oToSave.FileNames(1)
    End If
    
    On Error Resume Next
    oToSave.SaveAs sFileName
    On Error GoTo 0
    
End Function

Private Sub Form_Activate()
    If bLoading Then
        bLoading = False
        mIDE.formFillWorkArea Me
    End If
    
End Sub

Private Sub Form_Load()
    
    bLoading = True
    Me.LoadTree
    Me.txtOutput.Locked = True

    
End Sub

Public Function LoadTree()
    
    Dim iCurrProj As Integer
    Dim iCurrForm As Integer
    Dim iCurrCtl As Integer
    Dim sPrjKey As String
    Dim sDsgKey As String
    Dim sClsKey As String
    Dim sCmpKey As String
    Dim sCmpName As String
    Dim sDispName As String
    For iCurrProj = 1 To gVB.VBProjects.Count
        With gVB.VBProjects(iCurrProj)
            sPrjKey = gVB.VBProjects(iCurrProj).Name
            Me.tvProjectExplorer.Nodes.Add , tvwLast, sPrjKey, sPrjKey
            
            For iCurrForm = 1 To .VBComponents.Count
                sClsKey = sPrjKey & "~" & vbComponentType(.VBComponents(iCurrForm).Type)
                On Error Resume Next
                Me.tvProjectExplorer.Nodes.Add sPrjKey, tvwChild, sClsKey, vbComponentType(.VBComponents(iCurrForm).Type)
                On Error GoTo 0
                If .VBComponents(iCurrForm).Type = vbext_ct_ResFile Then
                    sCmpName = iCurrForm
                    sDispName = ts.sFileName(.VBComponents(iCurrForm).FileNames(1), efpFileName + efpFileExt)
                Else
                    sCmpName = .VBComponents(iCurrForm).Name
                    sDispName = sCmpName
                End If
                If sDispName = "" Then
                    sDispName = ts.sFileName(.VBComponents(iCurrForm).FileNames(1), efpFileNameAndExt)
                    sCmpName = iCurrForm
                End If
                sCmpKey = sClsKey & "~" & sCmpName
                Me.tvProjectExplorer.Nodes.Add sClsKey, tvwChild, sCmpKey, sDispName
                Me.tvProjectExplorer.Nodes(sCmpKey).EnsureVisible
            Next
            
        End With
        Me.tvProjectExplorer.Nodes(sPrjKey).Checked = True
        SetChildrenChecksTo Me.tvProjectExplorer.Nodes(sPrjKey), True
    Next
    
End Function

Public Function SetChildrenChecksTo(ByRef nodeParent As Node, ByVal bChecked As Boolean)
    
    If nodeParent.Children > 0 Then
        Dim nodeChild As Node
        Set nodeChild = nodeParent.Child
        Dim i As Integer
        For i = 1 To nodeParent.Children
            nodeChild.Checked = bChecked
            If nodeChild.Children > 0 Then
                SetChildrenChecksTo nodeChild, bChecked
            End If
            On Error Resume Next
            Set nodeChild = nodeChild.Next
            On Error GoTo 0
        Next
        Set nodeChild = Nothing
    End If
    
End Function



Private Sub Form_Resize()
    Static bResizing As Boolean
    If Not bResizing Then
        bResizing = True
        If Me.Width < Me.pbMain.Width * 1.5 Then
            Me.Width = Me.pbMain.Width * 1.5
        End If
        If Me.Height < Me.cmdZip.Height * 3 Then
            Me.Height = Me.cmdZip.Height * 3
        End If
'        With Me.chkZipIndividually
'            .Move 0, 0
'        End With
        With Me.chkUseWinZip
            .Move 0, 0
        End With
        With Me.chkIncludeEXE
            .Move Me.chkUseWinZip.Width, 0
        End With
        With Me.chkIncludeCompatible
            .Move Me.chkIncludeEXE.Left + Me.chkIncludeEXE.Width, 0
        End With
        With Me.cmdZip
            .Move 0, Me.chkUseWinZip.Height, Me.ScaleWidth, .Height
        End With
        With Me.tvProjectExplorer
            .Move 0, Me.cmdZip.Top + Me.cmdZip.Height + twipsY(1), Me.ScaleWidth / 3, Me.ScaleHeight - (Me.cmdZip.Top + Me.cmdZip.Height + twipsY(1)) - Me.sbMain.Height
        End With
        With Me.txtOutput
            .Move Me.tvProjectExplorer.Width + twipsX(1), Me.tvProjectExplorer.Top, Me.ScaleWidth - Me.tvProjectExplorer.Width, Me.tvProjectExplorer.Height
        End With
        With Me.sbMain
            .Move Me.pbMain.Width + twipsX(1), Me.ScaleHeight - .Height, Me.ScaleWidth - (Me.pbMain.Width + twipsX(1)), .Height
        End With
        With Me.pbMain
            .Move 0, Me.sbMain.Top, .Width, .Height
        End With
        bResizing = False
    End If
End Sub


Private Sub oZip_Progress(ByVal lCount As Long, ByVal sMsg As String)
'    If Me.chkZipIndividually.Value = 0 Then
        sMsg = ts.sNT(sMsg)
        sMsg = ts.sTrimChars(sMsg, Chr(10) & Chr(13))
        If Left(Trim(sMsg), 1) = "(" Then
            sMsg = vbTab & Trim(sMsg)
        End If
        Me.txtOutput = Me.txtOutput.Text & sMsg & vbCrLf
        Me.sbMain.SimpleText = sMsg
        Me.pbMain.Value = Me.pbMain.Value + 1
'    End If
End Sub

Private Sub tvProjectExplorer_NodeCheck(ByVal Node As MSComctlLib.Node)
    SetChildrenChecksTo Node, Node.Checked
End Sub
