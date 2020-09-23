VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'» Module: clsWindow    File: clsWindow.cls    Author: TheFrogPrince@hotmail.com
'»
'» This class file extends MS Windows "window object" to VB.  All
'» you have to do to use it is set the .hWnd property to a handle
'» of a window or object you want to manipulate or query.  The
'» parent window is automatically calculated and loaded, and a
'» simple call to the .RefreshChildren method will load the
'» collection of child windows.
'»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'» CODE DEPENDECIES
'»==================
'»      colWindows.cls -    a collection oriented class file that
'»                          is a collection of this class file.
'»      clsWNDClassEx.cls - this class file extends the WNDCLASS
'»                          type. Needed because VB does not
'»                          support hanging a user defined type
'»                          off of a class module.
'»      tools6.bas -        general collection of routines and
'»                          API declarations.
'»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'» NOTES
'»=======
'»
'»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
Option Explicit         ' of course =o)


'############################################
'######            CONSTANTS           ######
'############################################




'##############################################
'                 DECLARATIONS
'##############################################


'=*******************************=
'          USER32.dll
'=*******************************=
Private Declare Function GetWindow _
                                    Lib "user32" ( _
                                ByVal hWnd As Long, _
                                ByVal wCmd As enumGetWindowOptions) _
                            As Long        'ByVal wCmd As Long) As Long
Private Declare Function GetWindowText _
                                    Lib "user32" _
                                    Alias "GetWindowTextA" ( _
                                ByVal hWnd As Long, _
                                ByVal lpString As String, _
                                ByVal cch As Long) _
                            As Long
Private Declare Function GetWindowTextLength _
                                    Lib "user32" _
                                    Alias "GetWindowTextLengthA" ( _
                                ByVal hWnd As Long) _
                            As Long
Private Declare Function GetDesktopWindow _
                                    Lib "user32" () _
                            As Long
Private Declare Function SetActiveWindow _
                                    Lib "user32" ( _
                                ByVal hWnd As Long) _
                            As Long
Private Declare Function GetActiveWindow _
                                    Lib "user32" () _
                            As Long
Private Declare Function GetParent _
                                    Lib "user32" ( _
                                ByVal hWnd As Long) _
                            As Long
Private Declare Function GetClassName _
                                    Lib "user32" _
                                    Alias "GetClassNameA" ( _
                                ByVal hWnd As Long, _
                                ByVal lpClassName As String, _
                                ByVal nClassNameBufferLen As Long) _
                            As Long
Private Declare Function GetWindowLong _
                                    Lib "user32" _
                                    Alias "GetWindowLongA" ( _
                                ByVal hWnd As Long, _
                                ByVal nIndex As enumGetWindowLongOptions) _
                            As Long
Private Declare Function SetWindowLong _
                                    Lib "user32" _
                                    Alias "SetWindowLongA" ( _
                                ByVal hWnd As Long, _
                                ByVal nIndex As enumGetWindowLongOptions, _
                                ByVal dwNewLong As Long) _
                            As Long
Private Declare Function SetWindowText _
                                    Lib "user32" _
                                    Alias "SetWindowTextA" ( _
                                ByVal hWnd As Long, _
                                ByVal lpString As String) _
                            As Long
Private Declare Function GetClassInfoEx _
                                    Lib "user32" _
                                    Alias "GetClassInfoExA" ( _
                                ByVal hInstance As Long, _
                                ByVal sClassName As String, _
                                ByRef tWNDClassEx As Any) _
                            As Boolean

'=*******************************=
'          KERNEL32.dll
'=*******************************=
Private Declare Function GetModuleFileName _
                                    Lib "kernel32" _
                                    Alias "GetModuleFileNameA" ( _
                                ByVal hModule As Long, _
                                ByVal lpFileName As String, _
                                ByVal nSize As Long) _
                            As Long


'##############################################
'                 VARIABLES
'##############################################


'local variable(s) to hold property value(s)
Private mvarParent As clsWindow 'local copy
'local variable(s) to hold property value(s)
Private mvarChildren As colWindows 'local copy

'local variable(s) to hold property value(s)
Private mvarhWnd As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarsClassName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsWindowText As String 'local copy
'local variable(s) to hold property value(s)
Private mvareWindowStyle As enumWindowStyles 'local copy
'local variable(s) to hold property value(s)
Private mvarhInstance As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarcWNDClassEx As clsWNDClassEx 'local copy
'local variable(s) to hold property value(s)
Private mvarsModuleFileName As String 'local copy




'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»



'############################################
'######            FUNCTIONS           ######
'############################################

Public Property Get sModuleFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sModuleFileName
    Dim lNameLen As Long
    Dim sReturn As String * 512
    
    lNameLen = GetModuleFileName(Me.hInstance, sReturn, Len(sReturn))
    mvarsModuleFileName = Left(sReturn, lNameLen)
    sModuleFileName = mvarsModuleFileName
    
End Property



Public Property Set cWNDClassEx(ByVal vData As clsWNDClassEx)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.cWNDClassEx = Form1
    Set mvarcWNDClassEx = vData
End Property


Public Property Get cWNDClassEx() As clsWNDClassEx
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cWNDClassEx
    Set cWNDClassEx = mvarcWNDClassEx
End Property




Public Property Let hInstance(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hInstance = 5
    mvarhInstance = vData
    SetWindowLong mvarhWnd, GWL_HINSTANCE, vData
End Property


Public Property Get hInstance() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hInstance
    mvarhInstance = GetWindowLong(mvarhWnd, GWL_HINSTANCE)
    hInstance = mvarhInstance
End Property


Public Property Let eWindowStyle(ByVal vData As enumWindowStyles)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.eWindowStyle = Form1
    mvareWindowStyle = vData
    Dim lSuccess As Long
    MsgBox "Stop Here"
    lSuccess = SetWindowLong(mvarhWnd, GWL_STYLE, vData)
        
End Property


Public Property Get eWindowStyle() As enumWindowStyles
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.eWindowStyle
    Dim lReturn As Long
    lReturn = GetWindowLong(mvarhWnd, GWL_STYLE)  ' Window Style
    eWindowStyle = lReturn
    
End Property



Friend Property Let sWindowText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sWindowText = 5
    
    mvarsWindowText = vData
    SetWindowText mvarhWnd, vData
        
End Property


Friend Property Get sWindowText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sWindowText
    Dim sReturn As String * 512
    Dim lLen As Long
    lLen = GetWindowText(mvarhWnd, sReturn, Len(sReturn))      ' Window text
    sWindowText = sNT(sReturn)
    'Set sWindowText = mvarsWindowText
End Property



Public Property Let sClassName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClassName = 5
    mvarsClassName = vData
End Property


Public Property Get sClassName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClassName
    Dim sReturn As String * 512
    GetClassName mvarhWnd, sReturn, Len(sReturn)
    sClassName = sNT(sReturn)
End Property



Public Property Let hWnd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hWnd = 5
    mvarhWnd = vData
    
    ' Attach Parent Object (if possible)
    Dim lParentHWnd As Long
    lParentHWnd = GetParent(vData)
    If lParentHWnd > 0 Then
        If mvarParent Is Nothing Then
            Set mvarParent = New clsWindow
            mvarParent.hWnd = lParentHWnd
        End If
    End If
    
    ' Load Public Types
    Set mvarcWNDClassEx = New clsWNDClassEx
    Set mvarcWNDClassEx.Parent = Me
    
End Property


Public Property Get hWnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hWnd
    hWnd = mvarhWnd
End Property



Public Property Set Children(ByVal vData As colWindows)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Children = Form1
    Set mvarChildren = vData
End Property


Public Property Get Children() As colWindows
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Children
    Set Children = mvarChildren
End Property


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' RefreshInfo
'   Call this function to load/refresh any of the public
'   variables that are not dynamic properties.
''Public Sub RefreshInfo()
''
''End Sub



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' RefreshChildren
'   Call this function to load/refresh the list of
'   child windows.
Public Function RefreshChildren(Optional ByVal bRecursiveRefresh As Boolean) As Long
    
    Set mvarChildren = New colWindows
    
    Set mvarChildren.Parent = Me
    
    Dim lCurrChildHwnd As Long
    lCurrChildHwnd = GetWindow(mvarhWnd, GW_CHILD)
    
    Do While lCurrChildHwnd <> 0
        Dim cNewWindow As clsWindow
        Set cNewWindow = New clsWindow
        Set cNewWindow = mvarChildren.Add(lCurrChildHwnd)
        If bRecursiveRefresh Then
            cNewWindow.RefreshChildren bRecursiveRefresh
        End If
        lCurrChildHwnd = GetWindow(lCurrChildHwnd, GW_HWNDNEXT)
    Loop
    
End Function

Public Property Set Parent(ByVal vData As clsWindow)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Parent = Form1
    Set mvarParent = vData
End Property


Public Property Get Parent() As clsWindow
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Set Parent = mvarParent
End Property



Private Sub Class_Initialize()
    Set mvarChildren = New colWindows
    Set mvarChildren.Parent = Me
    
End Sub

Private Sub Class_Terminate()
    Set mvarChildren = Nothing
    Set mvarParent = Nothing
    
End Sub






' Code in the process of being "assimilated"
'




''''''''Option Explicit
'''''''''API declarations for destroying child forms
''''''''Private Declare Function SendMessage Lib "User32" Alias _
''''''''   "SendMessageA" (ByVal hWnd As Long, ByVal wMsg _
''''''''   As Long, ByVal wParam As Long, lParam As Any) As Long
''''''''Private Const WM_CLOSE = &H10
''''''''
''''''''Private Declare Function GetWindowLong Lib "User32" _
''''''''   Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal _
''''''''   nIndex As Long) As Long
''''''''Private Const GWL_STYLE = (-16)
''''''''
'''''''''Collection of children
''''''''Private chChildren() As Long
''''''''
''''''''Private Function DestroyChildren() As Boolean
''''''''Dim i As Integer, iNumChildren As Integer
''''''''DestroyChildren = True
''''''''iNumChildren = -1
'''''''''Find the number of child forms
''''''''On Error Resume Next
''''''''iNumChildren = UBound(chChildren)
''''''''On Error GoTo 0
''''''''If iNumChildren = -1 Then Exit Function
'''''''''Rip through the child forms, send a close message
'''''''''to each, wait and then check for them again.
''''''''For i = 0 To iNumChildren
''''''''   If chChildren(i) = 0 Then
''''''''   Else
''''''''      Call SendMessage(chChildren(i), WM_CLOSE, 0, 0)
''''''''      DoEvents
''''''''      If GetWindowLong(chChildren(i), GWL_STYLE) = _
''''''''         0 Then
''''''''         chChildren(i) = 0
''''''''      Else
''''''''         DestroyChildren = False
''''''''         Exit For
''''''''      End If
''''''''   End If
''''''''Next i
''''''''End Function
''''''''


Public Property Get Left() As Long
    Left = ts.windowCoords(mvarhWnd).Left
End Property

Public Property Get Top() As Long
    Top = ts.windowCoords(mvarhWnd).Top
End Property

Public Property Get Right() As Long
    Right = ts.windowCoords(mvarhWnd).Right
End Property

Public Property Get Bottom() As Long
    Bottom = ts.windowCoords(mvarhWnd).Bottom
End Property

Public Property Get ShowCmd() As enumShowWindow
    ShowCmd = ts.windowPlacement(mvarhWnd).ShowCmd
End Property

