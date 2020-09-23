VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7380
   ClientLeft      =   8730
   ClientTop       =   4425
   ClientWidth     =   10575
   _ExtentX        =   18653
   _ExtentY        =   13018
   _Version        =   393216
   Description     =   "Groovy little function to ZIP up a project (or projects). =)  "
   DisplayName     =   "Jeffs ZIP Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   SatName         =   "JeffsZIPAddIn.dll"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const guidMYTOOL$ = "_J_E_F_F_S_Z_I_P_V_B_A_D_D_I_N_"
Const sFriendlyID = "Jeff's ZIP VB6 Add-In"
    

Public WithEvents PrjHandler  As VBProjectsEvents          'projects event handler
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler  As VBComponentsEvents        'components event handler
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents CtlHandler  As VBControlsEvents          'controls event handler
Attribute CtlHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Dim mcbMenuCommandBar         As Office.CommandBarControl  'command bar object
Dim mcbMenuBar                As Office.CommandBar
'Dim mcbToolbox                As VBide.Window



Sub Show()
  'on error GoTo ShowErr

  gwinWindow.Visible = True
'  gdocJeffsAddIns.RefreshList 3

  Exit Sub
ShowErr:
  MsgBox Err.Description
End Sub

Public Property Get NonModalApp() As Boolean
  NonModalApp = True  'used by addin toolbar
End Property

'------------------------------------------------------
'this method adds the Add-In to the VB Tools menu
'it is called by the VB addin manager
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    'on error GoTo AddinInstance_OnConnectionErr
    
    Dim aiTmp As AddIn
    
    'save the vb instance
    Set gVB = Application
    
    If Not gwinWindow Is Nothing Then
        'already running so just show it
        Show
        If ConnectMode = ext_cm_AfterStartup Then
            'started from the addin manager
            AddToCommandBar
        End If
        Exit Sub
    End If
    
    'create the tool window
    If ConnectMode = ext_cm_External Then
        'need to see if it is already running
        'on error Resume Next
        Set aiTmp = gVB.Addins("JeffsZIPAddInForVB.Connect")
        'on error GoTo AddinInstance_OnConnectionErr
        If aiTmp Is Nothing Then
            'app is not in the VBADDIN.INI file so it is not in the collection
            'so lets attempt to use the 1st addin in the collection just
            'to get this app running and if there are none, an error
            'will occur and this app will not run
            Set gwinWindow = gVB.windows.CreateToolWindow(gVB.Addins("JeffsZIPAddInForVB.docJeffsAddIns"), "JeffsZIPAddInForVB.docJeffsAddIns", "JeffsZIPAddInForVB", guidMYTOOL$, gdocJeffsAddIns)
        Else
            If aiTmp.Connect = False Then
                Set gwinWindow = gVB.windows.CreateToolWindow(aiTmp, "JeffsZIPAddInForVB.docJeffsAddIns", "JeffsZIPAddInForVB", guidMYTOOL$, gdocJeffsAddIns)
            End If
        End If
    Else
        'must've been called from addin mgr
'        gVB.ReadOnlyMode = 2
        '                                                      This is the project Name+Designer             Project+Document
        Set gwinWindow = gVB.windows.CreateToolWindow(gVB.Addins("JeffsZIPAddInForVB.Connect"), "JeffsZIPAddInForVB.docJeffsAddIns", "JeffsZIPAddInForVB", guidMYTOOL$, gdocJeffsAddIns)  ' , "JeffsAddIns.docJeffsAddIns", "JeffsZIPAddInForVB", guidMYTOOL$, gdocJeffsAddIns)
    End If
    
    
    'sink the project, components and controls event handler
    Set Me.PrjHandler = gVB.Events.VBProjectsEvents
    Set Me.CmpHandler = gVB.Events.VBComponentsEvents(Nothing)
    Set Me.CtlHandler = gVB.Events.VBControlsEvents(Nothing, Nothing)
    
    If ConnectMode = vbext_cm_External Then
        'started from the addin toolbar
        Show
    ElseIf ConnectMode = vbext_cm_AfterStartup Then
        'started from the addin manager
        AddToCommandBar
    End If
    
    Exit Sub
    
AddinInstance_OnConnectionErr:
    MsgBox Err.Description
End Sub

'------------------------------------------------------
'this event removes the commandbar menu
'it is called by the VB addin manager
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
'  On Error GoTo IDTExtensibility_OnDisconnectionErr
  'delete the command bar entry
    SaveCommandBarSettings mcbMenuBar, "CommandBar"
    
    On Error Resume Next
  mcbMenuCommandBar.Delete
  mcbMenuBar.Delete
'  On Error GoTo 0
  Set mcbMenuCommandBar = Nothing
  Set mcbMenuBar = Nothing
  
  'save the form state for next time VB is loaded
  If gwinWindow.Visible Then
    jSaveSetting App.Title, "DisplayOnConnect", "1"
  Else
    jSaveSetting App.Title, "DisplayOnConnect", "0"
  End If
  
  Set gwinWindow = Nothing
  
IDTExtensibility_OnDisconnectionErr:
  
End Sub

'this event fires when the IDE is fully loaded
Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
  AddToCommandBar
End Sub

'this event fires when the command bar control is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Show
End Sub

#If Not bStandAloneControls Then
'this event fires when a control is added to the current form in the IDE
Private Sub CtlHandler_ItemAdded(ByVal vbControl As VBide.vbControl)
    
    If modMain.gdocJeffsAddIns.bAutoRename Then
        Load frmControlRenamer
        frmControlRenamer.LoadUpControl vbControl
        frmControlRenamer.Show vbModal
        
        If frmControlRenamer.bOK Then
            On Error Resume Next
            vbControl.ControlObject.Name = frmControlRenamer.sControlName
            vbControl.ControlObject.Caption = frmControlRenamer.sControlName
            vbControl.ControlObject.Text = frmControlRenamer.sControlName
            
        End If
        
        Set frmControlRenamer = Nothing
    End If
    
    If gwinWindow.Visible Then
        
    End If
End Sub
#End If

'this event fires when a control is renamed on the current form in the IDE
Private Sub CtlHandler_ItemRenamed(ByVal vbControl As VBide.vbControl, ByVal OldName As String, ByVal OldIndex As Long)
    If Not gwinWindow Is Nothing Then
  If gwinWindow.Visible Then
'    gdocJeffsAddIns.ControlRenamed VBControl, OldName, OldIndex
  End If
End If

End Sub

'this event fires when a control is removed from the current form in the IDE
Private Sub CtlHandler_ItemRemoved(ByVal vbControl As VBide.vbControl)
    If Not gwinWindow Is Nothing Then
  If gwinWindow.Visible Then
'    gdocJeffsAddIns.ControlRemoved VBControl
  End If
End If
End Sub

'this event fires when a form becomes activated in the IDE
Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBide.VBComponent)
  'on error GoTo CmpHandler_ItemActivatedErr
    If Not gwinWindow Is Nothing Then
  If gwinWindow.Visible Then
'    gdocJeffsAddIns.RefreshList 0
  End If
End If
CmpHandler_ItemActivatedErr:
End Sub

'this event fires when a form is selected in the project window
Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBide.VBComponent)
  CmpHandler_ItemActivated VBComponent
End Sub

Sub AddToCommandBar()
  'On Error GoTo AddToCommandBarErr
  On Error GoTo 0
  'make sure the standard toolbar is visible
  'gVB.CommandBars("Jeff's Cool VB6 Add-Ins").Visible = True
  
  'add it to the command bar
  'the following line will add the JeffsAddIns manager to the
  'Standard toolbar to the right of the ToolBox button
    Dim sTest As String
    On Error Resume Next
    sTest = gVB.CommandBars(sFriendlyID).Name
    If Err.Number <> 0 Then
        Set mcbMenuBar = AddCommandBar(gVB.CommandBars, "Jeff's ZIP Add-In 4VB6", msoBarFloating, False, False)
    Else
        Set mcbMenuBar = gVB.CommandBars(sFriendlyID)
    End If
    sTest = mcbMenuBar.Controls(sFriendlyID).Caption
    If Err.Number <> 0 Then
        Set mcbMenuCommandBar = AddCommandBarControl(mcbMenuBar, msoControlButton)
        mcbMenuCommandBar.DescriptionText = sFriendlyID
    End If
    On Error GoTo 0
  'set the caption
  mcbMenuCommandBar.Caption = "Jeff's ZIP Add-In"
  'copy the icon to the clipboard
  If modMain.LoadResPic2Control("icoJeff16-256", vbResIcon, mcbMenuCommandBar) Then
    MsgBox "ICO YEAH!"
    End If
  If modMain.LoadResPic2Control("bmpJeff-16b", vbResBitmap, mcbMenuCommandBar) Then
'    MsgBox "BMP YEAH!"
  End If
  
  
  
'  Clipboard.SetData LoadResPicture("icoJeff-16-256", vbResBitmap)
  'set the icon for the button
'  Clipboard.SetData LoadResPicture("ICOJEFF", 1), vbCFBitmap
  
  
'  Set mcbMenuCommandBar.Picture = LoadResPicture("GIFJEFF-16", 0)
    ' mcbMenuCommandBar.MaskColor = vbRed
    'Dim Jeff As PictureBox
    
    'Clipboard.GetData gdocJeffsAddIns.images16("JEFF").Image
    
    'Dim oImage As Object
'    Set oImage = CreateObject("image/gif")
'    Clipboard.SetData gdocJeffsAddIns.images16("JEFF")
    
''    Clipboard.SetData LoadResPicture("BMPJEFF-16", vbResBitmap)
'    Clipboard.SetData LoadResPicture("icoJEFF", vbResIcon)
    
    'modMain.LoadData2Control
    
'    mcbMenuCommandBar.PasteFace
'  mcbMenuBar.Position = GetSetting(APP_CATEGORY, App.Title, "CommandBarLocation", mcbMenuBar.Position)
'    mcbMenuBar.Visible = GetSetting(APP_CATEGORY, App.Title, "DisplayCommandBar", True)
  
  LoadCommandBarSettings mcbMenuBar, "CommandBar"
  'sink the event
  Set Me.MenuHandler = gVB.Events.CommandBarEvents(mcbMenuCommandBar)
  
  'restore the last state
  'If GetSetting(APP_CATEGORY, App.Title, "DisplayOnConnect", "0") = "1" Then
  If jGetSetting(App.Title, "DisplayOnConnect", "0") = "1" Then
    'set this to display the form on connect
    Me.Show
  End If
  
  Exit Sub
    
AddToCommandBarErr:
  MsgBox Err.Description
End Sub

Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBide.VBProject)
  'this takes care of the user removing the only project
      If Not gwinWindow Is Nothing Then
If gwinWindow.Visible Then
'    gdocJeffsAddIns.RefreshList 0
  End If
End If
End Sub

