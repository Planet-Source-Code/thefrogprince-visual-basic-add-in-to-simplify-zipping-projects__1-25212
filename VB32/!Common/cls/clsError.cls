VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Public Description As String

Public HelpContext As Long

Public HelpFile As String

Public LastDllError As Long

Public Number As Long
Attribute Number.VB_VarUserMemId = 0

Public Source As String


Public Sub Clear()
    
    Me.Description = ""
    Me.HelpContext = 0
    Me.HelpFile = ""
    Me.LastDllError = 0
    Me.Number = 0
    Me.Source = ""
    
End Sub

Public Function LoadFromResource(ByVal lID As Long)
    Me.Clear
    On Error Resume Next
    Me.Description = LoadResString(lID)
    If Err.Number = 0 Then
        Me.Number = lID
    End If
    
End Function

Public Function CopyFrom(ByVal oError As VBA.ErrObject)
    
    Me.Description = oError.Description
    Me.HelpContext = oError.HelpContext
    Me.HelpFile = oError.HelpFile
    Me.LastDllError = oError.LastDllError
    Me.Number = oError.Number
    Me.Source = oError.Source
    
End Function

Public Function Raise(Optional ByVal sText2PrePend As String = "")
    
    Err.Raise Me.Number, Me.Source, sText2PrePend & Me.Description, Me.HelpFile, Me.HelpContext
    
End Function

Public Function SetErr(ByVal lNumber As Long, Optional ByVal sDescription As String = "No Description Available For This Error", Optional ByVal sSource As String = "", Optional ByVal sHelpFile As String = "", Optional ByVal lHelpContext As Long = 0, Optional ByVal lLastDLLError As Long = 0)
    
    Me.Clear
    Me.Number = lNumber
    Me.Description = sDescription
    Me.Source = sSource
    Me.HelpContext = lHelpContext
    Me.HelpFile = sHelpFile
    Me.LastDllError = lLastDLLError
    
End Function

