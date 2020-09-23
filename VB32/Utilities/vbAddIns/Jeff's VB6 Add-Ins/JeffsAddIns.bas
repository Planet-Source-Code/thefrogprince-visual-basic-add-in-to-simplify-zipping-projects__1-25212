Attribute VB_Name = "modMain"
Option Explicit



Public Const COLOR_MENU = 4






Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long









Public Enum enumResDataTypes
    rdCursor = 1
    rdBitmap = 2
    rdIcon = 3
    rdMenu = 4
    rdDialogBox = 5
    rdString = 6
    rdFontDirectoryResource = 7
    rdFontResource = 8
    rdAcceleratorTable = 9
    rdUserDefined = 10
    rdGroupCursor = 12
    rdGroupIcon = 14
End Enum


Private Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal msg&, ByVal wp&, ByVal lp&)
Private Declare Sub SetFocus Lib "user32" (ByVal hwnd&)
Private Declare Function GetParent Lib "user32" (ByVal hwnd&) As Long
Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const WM_SYSCHAR = &H106
Const VK_F = 70  ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
Dim hwndMenu       As Long           'needed to pass the menu keystrokes to VB

Global gVB  As VBide.VBE       'instance of VB IDE
Global gwinWindow   As VBide.Window    'used to make sure we only run one instance
Global gdocJeffsAddIns As Object       'user doc object


Global Const APP_CATEGORY = "JeffsZIPAddInsForVB6"

#If bStandAloneControls Then
Const guidMYTOOL$ = "_J_E_F_F_S_Z_I_P_V_B_A_D_D_I_N_"
Const sFriendlyID = "Jeff's ZIP Add-In 4VB6"
    
#Else
Const guidMYTOOL$ = "_J_E_F_F_S_C_O_O_L_A_D_D_I_N_S_"
Const sFriendlyID = "Jeff's Cool VB6 Add-Ins"
#End If

Public Function jSaveSetting(ByVal sSection As String, ByVal sKey As String, ByVal sValue As Variant)
    SaveSetting sFriendlyID, sSection, sKey, sValue
    
End Function

Public Function jGetSetting(ByVal sSection As String, ByVal sKey As String, Optional ByVal sDefault As Variant) As String
    jGetSetting = GetSetting(sFriendlyID, sSection, sKey, sDefault)
End Function


Function InRunMode(VBInst As VBide.VBE) As Boolean
  InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function

Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
  If Shift <> 4 Then Exit Sub
  If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
  If gVB.DisplayModel = vbext_dm_SDI Then Exit Sub
  
  On Error Resume Next
  If hwndMenu = 0 Then hwndMenu = FindHwndMenu(ud.hwnd)
  On Error GoTo 0
  PostMessage hwndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000
  KeyCode = 0
  SetFocus hwndMenu
End Sub

Function FindHwndMenu&(ByVal hwnd&)
  Dim h As Long
  
Loop2:
  h = GetParent(hwnd)
  If h = 0 Then FindHwndMenu = hwnd: Exit Function
  hwnd = h
  GoTo Loop2
End Function


Public Function AddCommandBarControl(ByRef oCommandBar As CommandBar, ByVal eControlType As MsoControlType, Optional ByVal ControlID As Variant, Optional ByVal sParm As String, Optional ByVal InsertBeforeID As Variant, Optional ByVal bTemporary As Boolean) As CommandBarControl
    
    Set AddCommandBarControl = oCommandBar.Controls.Add(eControlType, ControlID, sParm, InsertBeforeID, bTemporary)
        
End Function



Public Function AddCommandBar(ByRef colCommandBars As CommandBars, ByVal sBarName As String, ByVal ePosition As MsoBarPosition, Optional ByVal bReplaceActiveMenuBar As Boolean = False, Optional ByVal bTemporary As Boolean = False) As CommandBar
    
    Set AddCommandBar = colCommandBars.Add(sBarName, ePosition, bReplaceActiveMenuBar, bTemporary)
    
End Function


Public Function LoadResPic2Control(ByVal sImageName As String, ByVal eImageType As LoadResConstants, ByRef ctlOnBar As CommandBarControl) As Boolean
    
    Dim eClipFormat As ClipBoardConstants
    
    Select Case True
        Case eImageType = vbResBitmap
            eClipFormat = vbCFBitmap
        Case eImageType = vbResCursor
            eClipFormat = vbCFDIB
        Case eImageType = vbResIcon
            eClipFormat = vbCFBitmap
        Case Else
            eClipFormat = vbCFEMetafile
    End Select
    
    LoadResPic2Control = False
    On Error Resume Next
    
    Err.Clear
    Dim Jeff  As Picture
    Set Jeff = LoadResPicture(sImageName, eImageType)
    
    Clipboard.SetData LoadResPicture(sImageName, eImageType)
    
    
    Dim cMap As COLORMAP
    cMap.From = &H80
    cMap.to = GetSysColor(COLOR_MENU)
    
    Clipboard.SetData BitmapToPicture(CreateMappedBitmap(App.hInstance, 101, 0, cMap, 1))
    
    If Err.Number = 0 Then
        ctlOnBar.PasteFace
        If Err.Number = 0 Then
            LoadResPic2Control = True
        End If
    End If
    
    On Error GoTo 0
    
End Function

Public Function LoadData2Control(ByVal sDataName As String, ByVal eDataType As enumResDataTypes, ByRef barCtl As CommandBarControl) As Boolean
    
    Dim eClipType As ClipBoardConstants
    
    eClipType = vbCFFiles
    Clipboard.Clear
    On Error Resume Next
    Clipboard.SetData LoadResData(sDataName, eDataType), eClipType
    
    Dim cMap As COLORMAP
    cMap.From = &H80
    cMap.to = GetSysColor(COLOR_MENU)
    
    Clipboard.SetData BitmapToPicture(CreateMappedBitmap(App.hInstance, 100, 0, cMap, 1))
        
    If Err = 0 Then
        barCtl.PasteFace
        If Err = 0 Then
            LoadData2Control = True
        End If
    End If
    
End Function


Public Function LoadImage2Control(ByRef imgCtl As Image, ByRef barCtl As CommandBarControl)
    
    MsgBox "Play Here..."
    
End Function



Public Function SaveCommandBarSettings(ByRef cbToSave As CommandBar, Optional ByVal sCommandBarTitle As String = "", Optional ByVal sApplicationTitle As String = "")
    
    On Error Resume Next
    CleanUpAppTitle sApplicationTitle
    CleanUpObjectTitle cbToSave, sCommandBarTitle
    sCommandBarTitle = Trim(sCommandBarTitle)
    jSaveSetting "Display Settings", sCommandBarTitle & ".Visible", cbToSave.Visible
    jSaveSetting "Display Settings", sCommandBarTitle & ".Protection", cbToSave.Protection
    jSaveSetting "Display Settings", sCommandBarTitle & ".Position", cbToSave.Position
    jSaveSetting "Display Settings", sCommandBarTitle & ".RowIndex", cbToSave.RowIndex
    SaveObjectCoords cbToSave, sCommandBarTitle, sApplicationTitle
    
End Function
Public Function LoadCommandBarSettings(ByRef cbToLoadInto As CommandBar, Optional ByVal sCommandBarTitle As String = "", Optional ByVal sApplicationTitle As String = "")
    
    On Error Resume Next
    CleanUpAppTitle sApplicationTitle
    CleanUpObjectTitle cbToLoadInto, sCommandBarTitle
    cbToLoadInto.Visible = jGetSetting("Display Settings", sCommandBarTitle & ".Visible", True)
    cbToLoadInto.Protection = jGetSetting("Display Settings", sCommandBarTitle & ".Protection", msoBarNoProtection)
    cbToLoadInto.Position = jGetSetting("Display Settings", sCommandBarTitle & ".Position", msoBarFloating)
    cbToLoadInto.RowIndex = jGetSetting("Display Settings", sCommandBarTitle & ".RowIndex", cbToLoadInto.RowIndex)
    LoadObjectCoords cbToLoadInto, sCommandBarTitle, sApplicationTitle

End Function

Public Function CleanUpAppTitle(ByRef sAppTitle As String) As String
    If Trim(sAppTitle) = "" Then
        If Trim(App.Title) = "" Then
            If Trim(App.EXEName) = "" Then
                sAppTitle = "ERROR in function CleanUpAppTitle"
            Else
                sAppTitle = App.EXEName
            End If
        Else
            sAppTitle = App.Title
        End If
    End If
    'CleanUpAppTitle = sAppTitle
End Function

Public Function CleanUpObjectTitle(ByRef oObject As Object, ByRef sObjectTitle As String) As String
    If Trim(sObjectTitle) = "" Then
        On Error Resume Next
        sObjectTitle = oObject.Name
        If Err <> 0 Or Trim(sObjectTitle) = "" Then
            sObjectTitle = oObject.Caption
            If Err <> 0 Or Trim(sObjectTitle) = "" Then
                sObjectTitle = oObject.Index
                If Err <> 0 Then
                    sObjectTitle = "ERROR in CLeanUpObjectTitle"
                End If
            End If
        End If
    End If
    sObjectTitle = Trim(sObjectTitle)
    'CleanUpObjectTitle = sObjectTitle
    
End Function
Public Function LoadObjectCoords(ByRef oLoadMe As Object, Optional ByRef sObjectTitle As String = "", Optional ByRef sApplicationTitle As String = "")
    CleanUpAppTitle sApplicationTitle
    CleanUpObjectTitle oLoadMe, sObjectTitle
    
    On Error Resume Next  ' Not all props may work
    oLoadMe.Top = jGetSetting("Display Settings", sObjectTitle & ".Top", oLoadMe.Top)
    oLoadMe.Left = jGetSetting("Display Settings", sObjectTitle & ".Left", oLoadMe.Left)
    oLoadMe.Width = jGetSetting("Display Settings", sObjectTitle & ".Width", oLoadMe.Width)
    oLoadMe.Height = jGetSetting("Display Settings", sObjectTitle & ".Height", oLoadMe.Height)
    On Error GoTo 0
    
End Function
Public Function SaveObjectCoords(ByRef oSaveMe As Object, Optional sObjectTitle As String = "", Optional ByVal sApplicationTitle As String = "")
    CleanUpAppTitle sApplicationTitle
    CleanUpObjectTitle oSaveMe, sObjectTitle
    
    jSaveSetting "Display Settings", sObjectTitle & ".Top", oSaveMe.Top
    jSaveSetting "Display Settings", sObjectTitle & ".Left", oSaveMe.Left
    jSaveSetting "Display Settings", sObjectTitle & ".Width", oSaveMe.Width
    jSaveSetting "Display Settings", sObjectTitle & ".Height", oSaveMe.Height
    
End Function




'"BMPJEFF-16"

Public Function BitmapToPicture(ByVal hBmp As Long, _
                         Optional ByVal hPal As Long = 0) As IPicture
    
    ' Code adapted from HardCore Visual Basic articles in MSDN
    ' Fill picture description
    Dim IPic As IPicture, picdes As PictDesc, iidIPicture As IID
    
    picdes.Size = Len(picdes)
    picdes.Type = vbPicTypeBitmap
    picdes.hBmp = hBmp
    picdes.hPal = hPal
    
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    iidIPicture.Data1 = &H7BF80980
    iidIPicture.Data2 = &HBF32
    iidIPicture.Data3 = &H101A
    iidIPicture.Data4(0) = &H8B
    iidIPicture.Data4(1) = &HBB
    iidIPicture.Data4(2) = &H0
    iidIPicture.Data4(3) = &HAA
    iidIPicture.Data4(4) = &H0
    iidIPicture.Data4(5) = &H30
    iidIPicture.Data4(6) = &HC
    iidIPicture.Data4(7) = &HAB
    
    ' Create picture from bitmap handle
    OleCreatePictureIndirect picdes, iidIPicture, True, IPic
    Set BitmapToPicture = IPic
    
End Function


