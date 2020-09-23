VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWNDClassEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'» Module: clsWNDClassEx   File: clsWNDClassEx.cls    Author: TheFrogPrince@hotmail.com
'»
'» This class file extends MS Windows WNDCLASS type.  In addition
'» to replicating all of members of the type, this class also provides
'» a .Parent property for linking to a loaded clsWindow object, as well
'» as method to Clear and Refresh the type data.  This class file is
'» necessary because VB does not allow you to hang a user defined type
'» from a class module, the clsWindow object exposes the class information
'» (via this class) of the window it is set to.
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'» CODE DEPENDECIES
'»==================
'»      clsWindow.cls -    a class file that extends the MS Windows
'»                          window object to VB.
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
Private Declare Function GetWindowLong _
                                    Lib "user32" _
                                    Alias "GetWindowLongA" ( _
                                ByVal hwnd As Long, _
                                ByVal nIndex As enumGetWindowLongOptions) _
                            As Long
Private Declare Function GetClassInfoEx _
                                    Lib "user32" _
                                    Alias "GetClassInfoExA" ( _
                                ByVal hInstance As Long, _
                                ByVal sClassName As String, _
                                ByRef tWNDClassEx As Any) _
                            As Boolean


'##############################################
'                 VARIABLES
'##############################################


' This Class Is An Exact Copy Of The WNDCLASSEX structure.
Public cbSize As Long
Public style As enumClassStyles
Public lpfnwndproc As Long
Public cbClsextra As Long
Public cbWndExtra As Long
Public hInstance As Long
Public hIcon As Long
Public hCursor As Long
Public hbrBackground As Long
Public lpszMenuName As String
Public lpszClassName As String
Public hIconSm As Long


'local variable(s) to hold property value(s)
Private mvarParent As clsWindow 'local copy

'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»



'############################################
'######            FUNCTIONS           ######
'############################################



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Clear
'   This function will clear the contents of the current
'   instance of the class.
Public Function Clear()
    Set mvarParent = Nothing
    
    Me.cbClsextra = 0
    Me.cbSize = 0
    Me.cbWndExtra = 0
    Me.hbrBackground = 0
    Me.hCursor = 0
    Me.hIcon = 0
    Me.hIconSm = 0
    Me.hInstance = 0
    Me.lpfnwndproc = 0
    Me.lpszClassName = 0
    Me.lpszMenuName = 0
    Me.style = 0

End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Refresh
'   This function will load the contents of the class based
'   on the attached parent window.
Public Function Refresh()
    
    If Not mvarParent Is Nothing Then
        
        Dim lhInstance As Long
        Dim lReturn As Long
        Dim tWNDClassEx As typeWNDClassEx
        On Error Resume Next
        lhInstance = GetWindowLong(mvarParent.hwnd, GWL_HINSTANCE)
        lReturn = GetClassInfoEx(lhInstance, mvarParent.sClassName, tWNDClassEx)
        On Error GoTo 0
        
        If lReturn <> 0 Then
            Me.cbClsextra = tWNDClassEx.cbClsextra
            Me.cbSize = tWNDClassEx.cbSize
            Me.cbWndExtra = tWNDClassEx.cbWndExtra
            Me.hbrBackground = tWNDClassEx.hbrBackground
            Me.hCursor = tWNDClassEx.hCursor
            Me.hIcon = tWNDClassEx.hIcon
            Me.hIconSm = tWNDClassEx.hIconSm
            Me.hInstance = tWNDClassEx.hInstance
            Me.lpfnwndproc = tWNDClassEx.lpfnwndproc
            Me.lpszClassName = tWNDClassEx.lpszClassName
            Me.lpszMenuName = tWNDClassEx.lpszMenuName
            Me.style = tWNDClassEx.style
        Else
            Me.Clear
        End If
    End If
    
End Function

Public Property Set Parent(ByVal vData As clsWindow)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Parent = Form1
    If Not vData Is Nothing Then
        Set mvarParent = vData
        
        Me.Refresh
    Else
        Me.Clear
    End If
End Property

Public Property Get Parent() As clsWindow
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Set Parent = mvarParent
End Property

