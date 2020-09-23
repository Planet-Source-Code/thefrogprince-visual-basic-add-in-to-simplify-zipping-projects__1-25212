Attribute VB_Name = "ts"
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'» Module: ts       File: tools6.bas        Author: TheFrogPrince@hotmail.com
'»
'» This module houses subs and functions that are common and
'» compatible with Windows 95/98/NT/2000.
'»
'»»»»»»»»»»»»»»»»»»»»»»»»»
'» CODE DEPENDECIES
'»==================
'»      mAPIconstants.bas   -   module to house all API or general VB constants
'=      clsError.cls        -   class file that supports extended error return
'=                              info from many of the functions that wrap errors
'=                              in On Error Resume Next
'=
'=   OPTIONAL DEPENDENCIES (see Notes below)
'=   ---------------------
'»      clsWindows.cls      -   class to allow window object manipulation
'»          - mdiTopForm()
'»      resTools6.res       -   an accompanying resource file containing
'»                              graphics and string tables.  If you are
'»                              already using a resource file, you will
'»                              need to merge this one in to yours,
'»                              otherwise, just use resTools6.res as a
'»                              starting template.
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'» NOTES
'»=======
'»  This module has a couple of compile time constants that it
'»  references.
'»      bShellLnkTLB -  This constant governs functions that make
'»      ------------    use of the ShellLnk.TLB.  Since this is
'»                      a fairly obscure reference that must be
'»                      manually included in all new VB6 projects
'»                      so... I deployed this constant so tools6
'»                      would remain easily includable in new
'»                      projects.  To turn on the functionality\
'»                      this type library provides, go to Project
'»                      References, and scroll down to:
'»                          VB 5 - IShellLinkA Interface(ANSI)
'»                      turn it on (nice how it's so far away from
'»                      the rest of the MS references)... and then
'»                      add to your compile time constants:
'»                          bShellLnkTLB = -1 :
'»
'»      bSHFolderDLL -  This constant governs functions that make
'»      ------------    use of the SHFolder.DLL.  There are no
'»                      references that need to be turned on with
'»                      this, but you will need to be sure that
'»                      you include the SHFolder.DLL file in your
'»                      install (since most packagers won't detect
'»                      the Declares in this module to it). To
'»                      turn on the related functions, add to your
'»                      compile time constants:
'»                          bSHFolderDLL = -1 :
'=      bWindowCls -    This constant indicates the presence of the
'=      ----------      windows Class files (clsWindow.cls, clsWNDClassEX.cls
'=                      and colWindows.cls).
'=      bRegistryCls -    This constant indicates the presence of the
'=      ----------      registry class file:  cRegistry.cls
'=      bTreeView -     This project contains a reference to the common
'=                      controls TreeView.
'=
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'= ERROR CODES
'= =-=-=-=-=-=
'=      This is the list of error codes that are used by this module.
'=      Everytime you create a new error in a routine, be sure you come
'=      here (to the top) and ADD IT.  =)  Will creatly assist future help
'=      files should we turn this into a DLL.
'=          -30000 = Path is valid but it is not a directory.
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
Option Explicit


'############################################
'######            VARIABLES           ######
'############################################



'############################################
'######            CONSTANTS           ######
'############################################




'############################################
'######          DECLARATIONS          ######
'############################################



'=-=-=-=-=-=-=-=-=
'  shFolder.dll
'=-=-=-=-=-=-=-=-=
#If bSHFolderDLL Then
Private Declare Function SHGetFolderPath _
                                   Lib "shfolder.dll" _
                                   Alias "SHGetFolderPathA" ( _
                               ByVal hWndOwner As Long, _
                               ByVal nFolder As enumShellFolders, _
                               ByVal hToken As Long, _
                               ByVal dwReserved As enumGetFolderPath, _
                               ByVal lpszPath As String) _
                            As Long     ' 0 if no error
#End If ' bSHFolderDLL

'=-=-=-=-=-=-=-=
'  User32.dll
'=-=-=-=-=-=-=-=
Public Declare Function menuHandle _
                                    Lib "user32" _
                                    Alias "GetMenu" ( _
                                ByVal hWnd As Long) _
                            As Long
Private Declare Function GetDesktopWindow _
                                    Lib "user32.dll" () _
                            As Long     ' The window hand (hWnd) of the Desktop
Private Declare Function GetWindow _
                                   Lib "user32.dll" ( _
                               ByVal hWnd As Long, _
                               ByVal wCmd As Long) _
                            As Long    ' The new window handle
Private Declare Function GetClassName _
                                   Lib "user32" _
                                   Alias "GetClassNameA" ( _
                               ByVal hWnd As Long, _
                               ByVal lpClassName As String, _
                               ByVal nClassNameBufferLen As Long) _
                            As Long
Private Declare Function SetActiveWindow _
                                    Lib "user32" ( _
                                ByVal hWnd As Long) _
                            As Long
Private Declare Function GetActiveWindow _
                                    Lib "user32" () _
                            As Long
Private Declare Function GetSystemMenu _
                                    Lib "user32" ( _
                                ByVal hWnd As Long, _
                                ByVal bRevert As Long) _
                            As Long
Private Declare Function RemoveMenu _
                                    Lib "user32" ( _
                                ByVal hMenu As Long, _
                                ByVal nPosition As Long, _
                                ByVal wFlags As Long) _
                            As Long
Private Declare Function LockWindowUpdate _
                                    Lib "user32" ( _
                                ByVal hwndLock As Long) _
                            As Long
Private Declare Function GetCursorPos _
                                    Lib "user32" ( _
                                lpPoint As POINTAPI) _
                            As Long
Private Declare Function GetSystemMetrics _
                                    Lib "user32" ( _
                                ByVal nIndex As Long) _
                            As Long
Private Declare Function GetWindowPlacement _
                                    Lib "user32" ( _
                                ByVal hWnd As Long, _
                                lpwndpl As WINDOWPLACEMENT) _
                            As Long
Private Declare Function GetKeyState _
                                    Lib "user32" ( _
                                ByVal nVirtKey As Long) _
                            As Integer
Private Declare Function GetWindowRect _
                                    Lib "user32" ( _
                                ByVal hWnd As Long, _
                                lpRect As Rect) _
                            As Long


'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»
'»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»»



'############################################
'######            FUNCTIONS           ######
'############################################
Private Const FunctionsStartHere = True


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' shellGetPathFor
'    This function will return the paths of the various different Windows
'    namespace folders (i.e. Desktop, My Documents, etc.)
#If bSHFolderDLL Then
Public Function shellGetPathFor( _
                        ByVal eFolderId As enumShellFolders, _
                        Optional ByVal hWnd As Long = -99) _
                As String
    
    If hWnd = -99 Then
        hWnd = GetDesktopWindow()
    End If

    Dim sReturn As String
    sReturn = Space(MAX_PATH)
    
    If bOK(SHGetFolderPath(hWnd, eFolderId, 0&, SHGFP_TYPE_DEFAULT, sReturn)) Then
        sReturn = sNT(sReturn)
        If Right(sReturn, 1) <> "\" Then
            sReturn = sReturn & "\"
        End If
        shellGetPathFor = sReturn
    End If
    
End Function
#End If ' bSHFolderDLL

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sNT
'    Standing for NullTrim, this function will take in a null terminated string
'    and clip of the extra junk.  Useful for DLL calls that return results in
'    a string buffer.
Public Function sNT( _
                        ByVal sString As String) _
                As String
                
    Dim iNullLoc As Integer
    iNullLoc = InStr(sString, Chr(0))
    If iNullLoc > 0 Then
        sNT = Left(sString, iNullLoc - 1)
    Else
        sNT = sString
    End If
End Function


#If bShellLnkTLB Then
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ShellCreateShorcut
'   This function will create a new .LNK (windows shortcut) file.
'   recreating the Object reference based on the pointer.
' Requires:  ShellLnk.tlb
Public Function shellCreateShorcut( _
                        ByVal sNewShortcutFile As String, _
                        ByVal sExeFile As String, _
                        Optional sWorkDir As String = "", _
                        Optional sExeArgs As String = "", _
                        Optional sIconFile As String = "", _
                        Optional lIconIdx As Long = 0, _
                        Optional sDescription As String = "", _
                        Optional eShowCmd As enumShowWindow = esw_SHOWDEFAULT) _
                As Boolean
    
    Dim rc As Long
    Dim Pidl As Long                                    ' Item id list
    Dim dwReserved As Long                              ' Reserved flag
    
    Dim cShellLink As ShellLinkA                        ' An explorer IShellLinkA(Win 95/Win NT) instance
    Dim cPersistFile As IPersistFile                    ' An explorer IPersistFile instance
    
    If ((sNewShortcutFile = "") Or (sExeFile = "")) Then
        Exit Function    ' Validate min. input requirements.
    End If
    
    On Error GoTo ErrHandler
    Set cShellLink = New ShellLinkA                     ' Create new IShellLink interface
    Set cPersistFile = cShellLink                       ' Implement cShellLink's IPersistFile interface
    
    If UCase(Right(Trim(sNewShortcutFile), 4)) <> ".LNK" Then
        sNewShortcutFile = sNewShortcutFile & ".LNK"
    End If
    
    With cShellLink
        .SetPath sExeFile                                ' set command line exe name & path to new ShortCut.
        
        If (sWorkDir <> "") Then .SetWorkingDirectory sWorkDir ' Set working directory in shortcut
        
        If (sExeArgs <> "") Then .SetArguments sExeArgs   ' Add arguments to command line
        
        If (sIconFile <> "") Then .SetIconLocation sIconFile, lIconIdx ' Set shortcut icon location & index
        
        If (sDescription <> "") Then .SetDescription sDescription & vbNullChar
        
        .SetShowCmd eShowCmd                             ' Set shortcut's startup mode (min,max,normal)
    End With
    
    cShellLink.Resolve 0, SLR_UPDATE
    cPersistFile.Save StrConv(sNewShortcutFile, vbUnicode), 0    ' Unicode conversion hack... This must be done!
    shellCreateShorcut = True                             ' Return Success
    
'---------------------------------------------------------------
ErrHandler:
    Set cPersistFile = Nothing                          ' Destroy Object
    Set cShellLink = Nothing                            ' Destroy Object

End Function
#End If

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' objFromPtr
'    This function provides a return trip back from the VB ObjPtr function,
'    recreating the Object reference based on the pointer.
Public Function objFromPtr( _
                        ByVal pObj As Long) _
                As Object
                
    Dim obj As Object
    ' force the value of the pointer into the temporary object variable
    CopyMemory obj, pObj, 4
    ' assign to the result (this increments the ref counter)
    Set objFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
    
End Function

' when finished, objCopy will allowing transfer binary data into a type variable
' for things like parsing file headers etc.etc.
''''Function objCopy(ByRef oObjectFrom As Test, ByRef oObjectTo As Variant)
''''    Dim pObjFrom As Long, pObjTo As Long
''''    Dim s As String
''''    s = "Jeff"
''''    'pObjFrom = ObjPtr(oObjectFrom)
'''''    pObjTo = ObjPtr(oObjectTo)
''''    CopyMemory pObjTo, pObjFrom, Len(oObjectFrom)
''''
''''End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ShellRecentDocumentsClear
'   This function will add a new file to the "Recent Documents" list.
Public Function shellRecentDocumentsAdd( _
                        ByVal sFileName As String) _
                As Boolean
    
    SHAddToRecentDocs SHARD_PATH, sFileName
        
End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ShellRecentDocumentsClear
'   This function will clear out the "Recent Documents" list.
Public Function shellRecentDocumentsClear()
    
    SHAddToRecentDocs SHARD_PATH, vbNullString
    
End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ShellGetPathTemp
'   This function returns the Windows temporary files directory.
Public Function shellGetPathTemp() As String
    
    Dim sDir As String
    Dim iReturnLen As Integer
    
    sDir = Space(MAX_PATH)
    iReturnLen = GetTempPath(Len(sDir), sDir)
    sDir = Left(sDir, iReturnLen)
    If Right(sDir, 1) <> "\" Then
        sDir = sDir & "\"
    End If
    
    shellGetPathTemp = sDir
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sLeftIs
'    String function to quickly determine if the leftmost characters in the
'    first string, equal the characters passed in the second string.  Compare
'    style can optionally be over-riden.
Public Function sLeftIs( _
                        ByVal sString2Check As String, _
                        ByVal sString2Check4 As String, _
                        Optional ByVal eCompareStyle As VbCompareMethod = vbBinaryCompare) _
                As Boolean
                
    sLeftIs = (StrComp(sPadR(sString2Check, Len(sString2Check4)), sString2Check4, eCompareStyle) = 0)
    
End Function



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sRightIs
'   String function to quickly determine if the rightmost characters in the
'   first string, equal the characters passed in the second string.  Compare
'   style can optionally be over-riden.
Public Function sRightIs( _
                        ByVal sString2Check As String, _
                        ByVal sString2Check4 As String, _
                        Optional ByVal eCompareStyle As VbCompareMethod = vbBinaryCompare) _
                As Boolean
                
    sRightIs = (StrComp(sPadL(sString2Check, Len(sString2Check4)), sString2Check4, eCompareStyle) = 0)
        
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' colHasItem
'   This function will spin through a collection and determine if
'   a passed value exists in the collection.
'   clear out all tab characters as well.
Public Function colHasItem( _
                        ByRef cCol2Check As Variant, _
                        ByRef vItem As Variant, _
                        Optional ByVal bIgnoreCase = True, _
                        Optional ByVal sCharacters2Ignore As String = "", _
                        Optional ByRef lIndex As Long) _
                As Boolean
    
    If cCol2Check.Count > 0 Then
        Dim i As Integer
        Dim iType As Integer
        iType = VarType(vItem)
        For i = 1 To cCol2Check.Count
            Select Case iType
                Case 8
                    If sLeftIs(sStripChars(IIf(bIgnoreCase, UCase(cCol2Check(i)), cCol2Check(i)), sCharacters2Ignore), sStripChars(IIf(bIgnoreCase, UCase(vItem), vItem), sCharacters2Ignore)) Then
                        lIndex = i
                        colHasItem = True
                        Exit For
                    End If
                Case Else
                    If cCol2Check.Item(i) = vItem Then
                        lIndex = i
                        colHasItem = True
                        Exit For
                    End If
            End Select
        Next i
    End If
    
End Function
''    If col2Check.Count > 0 Then
''
''        Dim vTest As Variant
''        On Error Resume Next
'''        vTest = col2Check(vItem)
''        ItemIsInCollection = (Not col2Check(vItem) Is Nothing)
''        ItemIsInCollection = (Err.Number = 0)
''        On Error GoTo 0
''    Else
''        ItemIsInCollection = False
''    End If
''




'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sStripChars
'   Quick string function to remove all specified characters
'   from a string.
Public Function sStripChars( _
                        ByVal sString As String, _
                        Optional ByVal sChars2Remove As String = "") _
                As String
    sStripChars = sConvertChars(sString, sChars2Remove, "", vbTextCompare)
'    Dim i As Integer
'    For i = 1 To Len(sChars2Remove)
'        sString = Replace(sString, Mid(sChars2Remove, i, 1), "")
'    Next i
'    sStripChars = sString
End Function

Public Function sKeepChars( _
                        ByVal sString As String, _
                        ByVal sCharsToKeep As String, _
                        Optional ByVal eCompareMethod As VbCompareMethod = vbBinaryCompare) _
                As String
                
    Dim i As Integer
    Dim sReturn As String
    Dim sCurrChar As String
    For i = 1 To Len(sString)
        sCurrChar = Mid(sString, i, 1)
        If InStr(1, sCharsToKeep, sCurrChar, eCompareMethod) > 0 Then
            sReturn = sReturn & sCurrChar
        End If
    Next
    sKeepChars = sReturn
    
End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sStripTabs
'    Quick function to strip out tabs from a string.
Public Function sStripTabs( _
                        ByVal sString As String) _
                As String
    
    sStripTabs = Trim(Replace(sString, Chr(9), ""))
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sTrimChars
'    Quick string function to trim specified characters from the beginning
'    and ending of strings.  Characters specified compose a LIST of
'    characters to be removed.  To trim strings, use the sTrimString
'    function instead.
Public Function sTrimChars( _
                        ByVal sString As String, _
                        ByVal sCharsToTrim As String, _
                        Optional ByVal eCompareStyle As VbCompareMethod = vbBinaryCompare) _
                As String
    
    Do While InStr(1, sCharsToTrim, Left(sString, 1), eCompareStyle) > 0 And Len(sString) > 0
        sString = Mid(sString, 2)
    Loop
    Do While InStr(1, sCharsToTrim, Right(sString, 1), eCompareStyle) > 0 And Len(sString) > 0
        sString = Left(sString, Len(sString) - 1)
    Loop
    sTrimChars = sString
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sTrimString
'    Quick string function to trim a specified string from
'    the beginning and ending of a string.
Public Function sTrimString( _
                        ByVal sString As String, _
                        ByVal sStringToTrim As String, _
                        Optional ByVal eCompareStyle As VbCompareMethod = vbBinaryCompare) _
                As String
    
    Do While Left(sString, Len(sStringToTrim)) = sStringToTrim
        sString = Mid(sString, Len(sStringToTrim) + 1)
    Loop
    Do While Right(sString, Len(sStringToTrim)) = sStringToTrim
        sString = Left(sString, Len(sString) - Len(sStringToTrim))
    Loop
    sTrimString = sString
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' idebInDevMode
'    This function will tell you whether or not the code is running
'    in the Visual Basic development environment (true), or running
'    as a compiled program (false).  Useful for things like Error
'    Handlers, where you want the behavior to be different when
'    deployed than it is when you are developing.
Public Function idebInDevMode() _
                As Boolean

    On Error GoTo ErrHandler
    'because debug statements are ignored when
    'the app is compiled, the next statment will
    'never be executed in the EXE.
    Debug.Print 1 / 0
    idebInDevMode = False

    Exit Function
    
ErrHandler:
    'If we get an error then we are
    'running in IDE / Debug mode
    idebInDevMode = True
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ideMsgBox
'    Use this function when you want to display a message box only while
'    developing your code, but not when the program is deployed.
Public Function ideMsgBox( _
                        ByVal sMessage As String, _
                        Optional ByVal eMsgBoxStyle As VbMsgBoxStyle = vbOKOnly, _
                        Optional ByVal sTitle As String = "", _
                        Optional ByVal sHelpFile As String = "", _
                        Optional ByVal lHelpContext As Long = 0) _
                As VbMsgBoxResult
                
    If ts.idebInDevMode Then
        ideMsgBox = MsgBox(sMessage, eMsgBoxStyle, sTitle, sHelpFile, lHelpContext)
    End If
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' colInMsgBox
'    Useful when debugging or developing, this function will display
'    a collection in a MsgBox window(s).  Line #s can be optionally
'    added, and the routine has been enhanced to take in the
'    .Fields object from a recordset, but you must specify this is
'    the case in the parameters (bFieldsCol = true, for the sake of
'    speed and simplicity).
Public Function colInMsgBox( _
                        ByRef cCollection As Variant, _
                        Optional ByVal bLineNumbers As Boolean = False, _
                        Optional ByVal bFieldsCol As Boolean = False)
    
    Dim i As Integer
    Dim sDisplay As String
    
    On Error GoTo ErrorHandler
    
    For i = IIf(bFieldsCol, 0, 1) To IIf(bFieldsCol, cCollection.Count - 1, cCollection.Count)
        
        If bFieldsCol Then
            sDisplay = sDisplay & IIf(bLineNumbers, Format(i, "000") & ":  ", "") & sPadR(cCollection(i).Name, 20) & ":" & cCollection(i) & vbCrLf
        Else
            sDisplay = sDisplay & IIf(bLineNumbers, Format(i, "000") & ":  ", "") & cCollection(i) & vbCrLf
        End If
        
        If (i + 1) / 24 = Int((i + 1) / 24) Then
            sDisplay = sDisplay & "<<MORE>>..."
            MsgBox sDisplay, vbOKOnly + vbInformation, "Collection Display"
            sDisplay = "<<CONTINUED>>..." & vbCrLf
        End If
    Next i
    
    MsgBox sDisplay, vbOKOnly + vbInformation, "Collection Display"
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error# " & Err.Number & vbCrLf & Err.Description
    
    
End Function

#If bWindowCls Then
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mdiTopForm
'   When coding an app in an MDI environment, this function
'   (which makes use of the clsWindow class file) will return
'   a reference to the child window that is currently on top and
'   having focus.
Public Function mdiTopForm( _
                        ByRef mdifrmParent As MDIForm, _
                        Optional ByVal bQuiet As Boolean = True) _
                As Form
    
    Dim frmReturn As Form
    Dim hWndMDIclient As Long
    Dim cCurrWin As New clsWindow
    cCurrWin.hWnd = mdifrmParent.hWnd
    cCurrWin.RefreshChildren
    Dim i As Integer
    For i = 1 To cCurrWin.Children.Count
        If UCase(cCurrWin.Children(i).sClassName) = "MDICLIENT" Then
            hWndMDIclient = cCurrWin.Children(i).hWnd
        End If
    Next i
    If hWndMDIclient = 0 Then
        If bQuiet Then
            
        Else
            Err.Raise -12300, "ts.mdiTopForm", LoadResString(12300)
        End If
        
    Else
        cCurrWin.hWnd = hWndMDIclient
        cCurrWin.RefreshChildren
        
        If cCurrWin.Children.Count = 0 Then
            If bQuiet Then
                
            Else
                Err.Raise -12301, "ts.mdiTopForm", LoadResString(12301)
            End If
        Else
            For i = 0 To VB.Forms.Count - 1
                If VB.Forms(i).hWnd = cCurrWin.Children(1).hWnd Then
                    Set frmReturn = VB.Forms(i)
                    Exit For
                End If
            Next i
        End If
    End If
    
    Set mdiTopForm = frmReturn
    Set frmReturn = Nothing
    Set cCurrWin = Nothing
            
End Function
#End If

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' errGetDescriptionBySYS
'    This function will take in an error # and return the text
'    description of that error based on the system wide error
'    string lookup table. Appropriate error #s are acheived via the
'    GetLastError() function or the Err.LastDllError property.
Public Function errGetDescriptionBySYS( _
                        Optional ByVal lErrorNum As Long = 0) _
                As String
    
    If lErrorNum = 0 Then
        lErrorNum = GetLastError()
    End If
        
    Dim sReturn As String * 4096
    Dim lStringLen As Long
    
    lStringLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM + FORMAT_MESSAGE_IGNORE_INSERTS, 0, lErrorNum, LANG__DEFAULT, sReturn, Len(sReturn), 0)
    errGetDescriptionBySYS = Left(sReturn, lStringLen)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' errGetDescriptionByDLL
'    This function will take in an error # and return the text
'    description of that error based on the error string lookup
'    in a DLL file.  Appropriate error #s are acheived via the
'    GetLastError() function or the Err.LastDllError property.
Public Function errGetDescriptionByDLL( _
                        ByVal sDLLName As String, _
                        ByVal lErrorNum As Long) _
                As String
    'A work in progress =o)
    
    Dim sReturn As String * 4096
    Dim sSearchReplace() As Variant
    Dim lStringLen As Long
    Dim lhModule As Long
    lhModule = GetModuleHandle(sDLLName)
    
    lStringLen = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE + FORMAT_MESSAGE_IGNORE_INSERTS, lhModule, lErrorNum, 0, sReturn, Len(sReturn), 0)
    
    Dim lTest As Long
    lTest = FormatMessage(123345, 1234, 1234, 0, sReturn, Len(sReturn), 0)
       
    
    If FormatMessage(FORMAT_MESSAGE_IGNORE_INSERTS + FORMAT_MESSAGE_MAX_WIDTH_MASK + FORMAT_MESSAGE_FROM_SYSTEM, 0, Err.LastDllError, LANG_ENGLISH, sReturn, Len(sReturn), ByVal 0&) Then
        MsgBox "Check it out."
        errGetDescriptionByDLL = sNT(sReturn)
    Else
        MsgBox "Hmmm?"
    End If
    Debug.Print FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, GetModuleHandle("C:\WINDOWS\SYSTEM\KERNEL32.DLL"), 123, 0, sReturn, Len(sReturn), 0)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' debugCharChart
'    This function just creates a very simple ASCII character
'    chart in your immediate window.  Useful when making up
'    cut and "colorful" names for MSN Messenger.  =)
Public Function debugCharChart()
    Dim sChart As String
    Dim i As Integer
    For i = 0 To 255
        sChart = sChart & sPadR(Chr(i), 1) & " "
        If (i + 1) / 16 = Int((i + 1) / 16) Then
            sChart = sChart & vbCrLf
        End If
    Next i
    Debug.Print sChart
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sPadL
'    This function will take in a string and pad it on the
'    left until the pad length is reached, optionally using
'    a character you specify.
Public Function sPadL( _
                        ByVal sInstring As String, _
                        ByVal lPadLength As Long, _
                        Optional sPadChar As String = " ") _
                As String
    sInstring = LTrim(sInstring)
    sPadChar = Left(sPadChar, 1)
    If Len(sInstring) > lPadLength Then
        sPadL = Mid(sInstring, Len(sInstring) - lPadLength + 1)
    Else
        sPadL = String(lPadLength - Len(sInstring), sPadChar) & sInstring
    End If
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sPadR
'    This function will take in a string and pad it on the
'    right until the pad length is reached, optionally using
'    a character you specify.
Public Function sPadR( _
                        ByVal sInstring As String, _
                        ByVal lPadLength As Long, _
                        Optional sPadChar As String = " ") _
                As String
    sInstring = RTrim(sInstring)
    sPadChar = Left(sPadChar, 1)
    If Len(sInstring) > lPadLength Then
        sPadR = Left(sInstring, lPadLength)
    Else
        sPadR = sInstring & String(lPadLength - Len(sInstring), sPadChar)
    End If
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' txtAddO
' txtAddV
'    This simple little function is useful when you have
'    several lines of code that are appending text to a
'    text box (or other object that has a text property).
'    Save yourself some keystrokes and use this.
Public Function txtAddO( _
                        ByRef oObjectWithTextProp As Object, _
                        ByVal sText2Add As String)
                        
    oObjectWithTextProp = oObjectWithTextProp & sText2Add
    
End Function
Public Function txtAddV( _
                        ByRef oObjectWithTextProp As Variant, _
                        ByVal sText2Add As String)
                        
    oObjectWithTextProp = oObjectWithTextProp & sText2Add
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' twipsX
' twipsY
'    These functions returns the proper # of twips for
'    # of pixels passed in (x = horizontal, y = vertical).
'    This calculation covers the minute graphical variations
'    that occur from computer to computer and monitor to monitor.
Public Function twipsX( _
                        ByVal PixelsIn As Variant) _
                As Long
    twipsX = PixelsIn * Screen.TwipsPerPixelX
End Function
Public Function twipsY( _
                        ByVal PixelsIn As Variant) _
                As Long
    twipsY = PixelsIn * Screen.TwipsPerPixelY
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sConvertChars
'    This function will convert all occurences of particular
'    characters in a string to another characters.
Public Function sConvertChars(ByVal sString2Convert As String, ByVal sCharactersToConvert As String, ByVal sChar2Convert2 As String, Optional ByVal eCompareStyle As VbCompareMethod = vbBinaryCompare) As String
    Dim i As Integer
    For i = 1 To Len(sCharactersToConvert)
        sString2Convert = Replace(sString2Convert, Mid(sCharactersToConvert, i, 1), Left(sChar2Convert2, 1), , , eCompareStyle)
    Next i
    sConvertChars = sString2Convert
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sysmenuHandle
'    Returns the system menu handle for the specified
'    window (via hWnd).
Public Function sysmenuHandle(ByVal hWnd As Long) As Long
    sysmenuHandle = GetSystemMenu(hWnd, False)
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sysmenuRemoveCommand
'    This function will remove a command from a specified system menu
'    handle.  Use the sysmenuHandle() function to retrieve the system
'    menu handle for a particular form or window.
Public Function sysmenuRemoveCommand(ByVal smHandle As Long, ByVal eCommandToRemove As enumSysMenuCommands) As Boolean
    RemoveMenu smHandle, eCommandToRemove, MF_BYCOMMAND
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' windowUpdate
'    This function allows locking or unlocking the repaint event
'    of a specified window (via hWnd).
Public Function windowUpdate(ByRef hWnd As Long, ByVal eLockOrUnlock As enumLockWindowUpdates) As Boolean
    
    If eLockOrUnlock = elwLOCK Then
        windowUpdate = LockWindowUpdate(hWnd)
    Else
        windowUpdate = LockWindowUpdate(0)
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlSetFocus
'    This function will set the current focus to the specified
'    control or screen object without throwing an error if the
'    object cannot receive focus.
Public Function ctlSetFocus(ByRef ObjToSetFocusTo As Object) As Boolean
    On Error Resume Next
    ObjToSetFocusTo.SetFocus
    ctlSetFocus = Err.Number = 0
    On Error GoTo 0
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlKeyPress
'    This function is handy for wrapping input to textboxes
'    or other controls that have the KeyPress event to implement
'    standard types of input masks.
'       Example:
'            Private Sub txtPlaceOfEmployment_KeyPress(KeyAscii As Integer)
'                KeyAscii = ts.wrapKeyPress(KeyAscii, Uppercase + NoDoubleQuotes)
'            End Sub
Public Function ctlKeyPress(ByVal KeyAscii As KeyCodeConstants, ByVal TypeToAllow As enumKeyPressAllowTypes) As Integer
    
    Dim ltrKeyAscii As Integer
    ltrKeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    ' By default pass the keystroke through and then optionally kill it
    ctlKeyPress = KeyAscii
    
    ' Default Keystrokes to allow (enter, backspace, delete, escape)
    If _
        KeyAscii = vbKeyReturn Or _
        KeyAscii = vbKeyEscape Or _
        KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Then
        
        Exit Function
    End If
    
    ' NumbersOnly
    If (TypeToAllow And OnlyNumbers) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case (KeyAscii = vbKeySubtract Or KeyAscii = Asc("-")) And (TypeToAllow And AllowNegative)
            Case KeyAscii = Asc("#") And (TypeToAllow And AllowPounds)
            Case KeyAscii = Asc("*") And (TypeToAllow And AllowStars)
            Case KeyAscii = vbKeyDecimal And (TypeToAllow And AllowDecimal)
            Case KeyAscii = vbKeySpace And (TypeToAllow And AllowSpaces)
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' DatesOnly
    If (TypeToAllow And OnlyDates) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = vbKeyDivide Or KeyAscii = Asc("/")
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' TimesOnly
    If (TypeToAllow And OnlyTimes) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = Asc(":") Or KeyAscii = Asc(";")
                ctlKeyPress = Asc(":")
            Case ltrKeyAscii = vbKeyA Or ltrKeyAscii = vbKeyP Or ltrKeyAscii = vbKeyM
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' LettersOnly
    If (TypeToAllow And OnlyLetters) Then
        Select Case True
            Case ltrKeyAscii >= vbKeyA And ltrKeyAscii <= vbKeyZ
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' UpperCase
    If (TypeToAllow And Uppercase) Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    ' No Spaces
    If (TypeToAllow And NoSpaces) And KeyAscii = vbKeySpace Then
        KeyAscii = 0
    End If
    
    ' No Double Quotes
    If (TypeToAllow And NoDoubleQuotes) And KeyAscii = Asc("""") Then
        KeyAscii = Asc("'")
    End If
    
    ' No Single Quotes
    If (TypeToAllow And NoSingleQuotes) And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    
    ctlKeyPress = KeyAscii
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlLocked
'    This function is an enhanced method by which to set the
'    "locked" property of a control.  Some controls have a
'    .Locked property, others a .ReadOnly, others you can
'    only set the .Enabled property.
Public Function ctlLocked( _
                        ByRef ctlObject As Control, _
                        ByVal bLocked As Boolean, _
                        Optional ByVal bPreserveColor As Boolean = False, _
                        Optional ByVal lAlternateLockedColor As SystemColorConstants = vbButtonFace)
    
    On Error Resume Next
    ctlObject.Locked = bLocked
    If Err.Number <> 0 Then
        Err.Clear
        ctlObject.ReadOnly = bLocked
    End If
    If Err.Number <> 0 Then
        Err.Clear
        ctlObject.Enabled = Not bLocked
    End If
    If Not bPreserveColor Then
        If bLocked Then
            ctlObject.BackColor = lAlternateLockedColor
        Else
            ctlObject.BackColor = vbWindowBackground
        End If
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlEnabled
'    This function provides a slightly enhanced method for setting the
'    .Enabled property of a control.  Certain controls (like the CheckBox)
'    have a bug that changes the .BackColor property depending on
'    the .Enabled property.  This routine corrects this MS over-sight.
Public Function ctlEnabled( _
                            ByRef ctlObject, _
                            ByVal bEnabled As Boolean, _
                            Optional ByVal bPreserveBackColor As Boolean = False, _
                            Optional ByVal bAffectTabStop As Boolean = False, _
                            Optional ByVal lAlternateLockedColor As SystemColorConstants = vbButtonFace)
    
    On Error Resume Next
    ctlObject.Enabled = bEnabled
    If bAffectTabStop Then
        ctlObject.TabStop = bEnabled
    End If
    If Not bPreserveBackColor Then
        If ctlObject.Enabled Then
            ctlObject.BackColor = vbWindowBackground
        Else
            ctlObject.BackColor = lAlternateLockedColor
        End If
    End If
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagValue
'    Simple function for attempting to coerce any value into a control's
'    tag.  Useful only when it doesn't really matter whether the value
'    is stored or not (like trying to save NULLs to the tag).  This function
'    has been superceded by the other ctlTagData... functions.
Public Function ctlTagValue(ByRef ctlObject As Control, ByVal vValue As Variant)
    
    On Error Resume Next
    ctlObject.Tag = vValue
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlVisible
'    This function just provides and error proof way to set a controls
'    visibility.  Useful for when the state of the entire interface is unknown.
'    (i.e. an error will occur if you try to set the .Visible property of a control
'    who's form is currently invisible).
Public Function ctlVisible(ByRef ctlObject As Control, ByVal bVisible As Boolean)
    
    On Error Resume Next
    ctlObject.Visible = bVisible
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlProperty
'    This function is an adjunct to the ctlHasProperty function. Both functions
'    support the same enum.  This function will return the value of the property.
Public Function ctlProperty(ByRef ctlObject As Control, ByVal eCtrlProp As enumControlProperties) As Variant
    
    On Error Resume Next
    Select Case True
        Case eCtrlProp = cpWordWrap
            ctlProperty = forceBool(ctlObject.WordWrap)
        Case eCtrlProp = cpMultiLine
            ctlProperty = forceBool(ctlObject.MultiLine)
            
    End Select
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' osVersionInfo
'    This function return a type containing information about the operating
'    system, and thus can be referenced like this:
'        EXAMPLE:
'            If osVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
'                TurnNoseUpInDisgust()
'                End
'            End If
Public Function osVersionInfo() As typeOSVERSIONINFO
    Dim oInfo As typeOSVERSIONINFO
    oInfo.dwOSVersionInfoSize = Len(oInfo)
    GetVersionEx oInfo
    osVersionInfo = oInfo
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' osbWindowsNT
'    This function provides a quick boolean for detecting the Windows NT
'    operating system.  This is need in some functions because they reference
'    dll calls that are not valid on the NT platform.
Public Function osbWindowsNT() As Boolean
    
    osbWindowsNT = (osVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sQ
'    Quick string function to "quote" a string, preparing it for use possibly
'    in a SQL statement or .Find clause.
Public Function sQ( _
                    ByVal sString2Quote As String, _
                    Optional ByVal eQuoteType As enumQuoteTypes = qtSingleTick) _
                As String
    Dim sChar As String
    If eQuoteType = qtSingleTick Then
        sChar = "'"
    Else
        sChar = """"
    End If
    sString2Quote = sString2Quote & ""
    If Left(sString2Quote, 1) <> sChar Then
        sString2Quote = sChar & sString2Quote
    End If
    If Right(sString2Quote, 1) <> sChar Then
        sString2Quote = sString2Quote & sChar
    End If
    sQ = sString2Quote
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sEncode
'    This string function will encrypt a string using a key you pass it.
Public Function sEncode( _
                    ByVal S As String, _
                    key As Long, _
                    salt As Boolean) _
                As String
    
    Dim n As Long, i As Long, ss As String
    Dim k1 As Long, k2 As Long, k3 As Long, _
         k4 As Long, t As Long
    Static saltvalue As String * 4
    
    If salt Then
         For i = 1 To 4
              t = 100 * (1 + Asc(Mid(saltvalue, i, 1))) * Rnd() * (Timer + 1)
              Mid(saltvalue, i, 1) = Chr(t Mod 256)
         Next
         S = Mid(saltvalue, 1, 2) & S & Mid(saltvalue, 3, 2)
    End If
    
    n = Len(S)
    ss = Space(n)
    ReDim sn(n) As Long
    
    k1 = 11 + (key Mod 233): k2 = 7 + (key Mod 239)
    k3 = 5 + (key Mod 241): k4 = 3 + (key Mod 251)
    
    For i = 1 To n: sn(i) = Asc(Mid(S, i, 1)): Next
    
    For i = 2 To n: sn(i) = sn(i) Xor sn(i - 1) Xor _
         ((k1 * sn(i - 1)) Mod 256): Next
    For i = n - 1 To 1 Step -1: sn(i) = sn(i) Xor sn(i + 1) Xor _
         (k2 * sn(i + 1)) Mod 256: Next
    For i = 3 To n: sn(i) = sn(i) Xor sn(i - 2) Xor _
         (k3 * sn(i - 1)) Mod 256: Next
    For i = n - 2 To 1 Step -1: sn(i) = sn(i) Xor sn(i + 2) Xor _
         (k4 * sn(i + 1)) Mod 256: Next
    
    For i = 1 To n: Mid(ss, i, 1) = Chr(sn(i)): Next
    
    sEncode = ss
    saltvalue = Mid(ss, Len(ss) / 2, 4)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sDecode
'    Valid only with data encrypted with the sEncode function in this library, this function
'    will decrypt a string using the key # you pass it.
Public Function sDecode( _
                        ByVal S As String, _
                        key As Long, _
                        salt As Boolean) _
                As String
    
    Dim n As Long, i As Long, ss As String
    Dim k1 As Long, k2 As Long, k3 As Long, k4 As Long
    
    n = Len(S)
    ss = Space(n)
    ReDim sn(n) As Long
    
    k1 = 11 + (key Mod 233): k2 = 7 + (key Mod 239)
    k3 = 5 + (key Mod 241): k4 = 3 + (key Mod 251)
    
    For i = 1 To n: sn(i) = Asc(Mid(S, i, 1)): Next
    
    For i = 1 To n - 2: sn(i) = sn(i) Xor sn(i + 2) Xor _
         (k4 * sn(i + 1)) Mod 256: Next
    For i = n To 3 Step -1: sn(i) = sn(i) Xor sn(i - 2) Xor _
         (k3 * sn(i - 1)) Mod 256: Next
    For i = 1 To n - 1: sn(i) = sn(i) Xor sn(i + 1) Xor _
         (k2 * sn(i + 1)) Mod 256: Next
    For i = n To 2 Step -1: sn(i) = sn(i) Xor sn(i - 1) Xor _
         (k1 * sn(i - 1)) Mod 256: Next
    
    For i = 1 To n: Mid(ss, i, 1) = Chr(sn(i)): Next i
    
    If salt Then sDecode = Mid(ss, 3, Len(ss) - 4) Else sDecode = ss
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sAppend
'    This function will append a string to another string when it is not already the
'    last character or characters in the string (useful for ensuring a string is ended
'    with a vbCrLf or when building paths, a backslash \).
Public Function sAppend(ByVal s2AppendTo As String, ByVal sChars2Append As String) As String
    
    If Right(s2AppendTo, Len(sChars2Append)) <> sChars2Append Then
        sAppend = s2AppendTo & sChars2Append
    Else
        sAppend = s2AppendTo
    End If
    
End Function



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' fileExists
'    This functions returns whether a file exists or not.  Uses the special error class to
'    return error information if the file does not exist.
Public Function fileExists(ByVal File2Look4 As String, Optional ByRef oError) As Boolean
    
    Dim FileHandle As Integer
    
    FileHandle = FreeFile
    Err.Clear
    On Error Resume Next
    Open File2Look4 For Input As #FileHandle
    fileExists = (Err.Number = 0)
    oError.CopyFrom Err
    Close #FileHandle
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' dirExists
'    This function returns whether a directory exists or not.  Uses the special error class
'    to return error information if the directory does not exist.
Public Function dirExists(ByVal sPathName As String, Optional ByRef oError) As Boolean
    
    If GetFileAttributes(sPathName) < 0 Then
        fileExists sPathName, oError  ' We'll use the fileexists function to
                                      ' set the error information on FileNotFound
                                      ' directly from VB
        dirExists = False
    Else
        Dim bExists As Boolean
        bExists = (GetFileAttributes(sPathName) And efaDIRECTORY) > 0
        If Not bExists Then
            ts.errSet oError, -30000, "This was a valid path to a file system object, but the object was not a directory."
        End If
        dirExists = bExists
    End If
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' errSet
'    Used with the error class, this function sets an error
'    object when you don't know whether or not it was
'    passed as a parameter to a function or whether it
'    exists within context yet.
Public Function errSet( _
                        ByRef oError, _
                        ByVal lNumber As Long, _
                        Optional ByVal sDescription As String, _
                        Optional ByVal sSource As String, _
                        Optional ByVal sHelpFile As String, _
                        Optional ByVal lHelpContext As Long, _
                        Optional ByVal lLastDLLError As Long)
    
    On Error Resume Next
    oError.SetErr lNumber, sDescription, sSource, sHelpFile, lHelpContext, lLastDLLError
    On Error GoTo 0
    
End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nLoWord
'    Used for binary manipulation (such as parsing info from a file header), this
'    function return the Low Order Word of a DWord.
Public Function nLoWord(inval As Long) As Integer
    
    nLoWord = nDWord(inval).LoWord
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nHiWord
'    Used for binary manipulation (such as parsing info from a file header), this
'    function return the High Order Word of a DWord.
Public Function nHiWord(inval As Long) As Integer
    
    nHiWord = nDWord(inval).HiWord
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nMakeDWord
'    Used for binary manipulation (such as parsing info from a file header), this
'    function will create a DWord from a Low and High order word that you pass.
Public Function nMakeDWord(wHi As Integer, wLo As Integer) As Long
    
    If wHi And &H8000& Then
        nMakeDWord = (((wHi And &H7FFF&) * 65536) Or (wLo And &HFFFF&)) Or &H80000000
    Else
        nMakeDWord = (wHi * 65536) + wLo
    End If
    
End Function '(Public) Function MakeDWord()

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nMakeWord
'    Used for binary manipulation (such as parsing info from a file header), this
'    function will create a Word from a Low and High order byte that you pass.
Public Function nMakeWord(ByVal bHi As Byte, ByVal bLo As Byte) As Integer
    
    If bHi And &H80 Then
        nMakeWord = (((bHi And &H7F) * 256) + bLo) Or &H8000
    Else
        nMakeWord = (bHi * 256) + bLo
    End If
    
End Function '(Public) Function MakeWord()

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlIsLocked
'    This function provides an enhanced method of determining whether a control is in a
'    "locked" state or not.  For those controls not having a true .Locked property, it will
'    even respond to a LOCKED value when it is stored in the tag with the ctlTagDataSet
'    function in this library.
Public Function ctlIsLocked(ByRef ctl2Check) As Boolean
    
    Dim bUseTag As Boolean
    If ts.ctlHasProperty(ctl2Check, ppTag) Then
        bUseTag = ts.ctlTagDataHas(ctl2Check, "LOCKED")
    End If
    Select Case True
        Case bUseTag
            ctlIsLocked = CBool(ts.ctlTagDataGet(ctl2Check, "LOCKED"))
        Case ts.ctlHasProperty(ctl2Check, ppLocked)
            ctlIsLocked = bAbs(ctl2Check.Locked)
        Case ts.ctlHasProperty(ctl2Check, ppReadOnly) And Not TypeOf ctl2Check Is FileListBox
            ctlIsLocked = bAbs(ctl2Check.ReadOnly)
        Case ts.ctlHasProperty(ctl2Check, ppEnabled)
            ctlIsLocked = Not bAbs(ctl2Check.Enabled)
        Case Else
            ctlIsLocked = False
    End Select
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlHasProperty
'    This function is used when the type of a control is not known.  With this function
'    you can test for the existence of properties before you attempt to actually use
'    them in code.
Public Function ctlHasProperty(ByRef ctl2Check, eProp As enumPossibleProperties) As Boolean
    Dim vTest As Variant
    On Error Resume Next
    Select Case True
        Case eProp = ppLocked
            vTest = ctl2Check.Locked
        Case eProp = ppReadOnly
            vTest = ctl2Check.ReadOnly
        Case eProp = ppEnabled
            vTest = ctl2Check.Enabled
        Case eProp = ppTag
            vTest = ctl2Check.Tag
    End Select
    ctlHasProperty = (Err.Number = 0)
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataSet
'    Part of the ctlTagData set of functions, this function will allow you set a value
'    in a controls .Tag property.
Public Function ctlTagDataSet( _
                            ByRef ctl2Set, _
                            ByVal sDataKey As String, _
                            ByVal sDataValue As String)
    Dim sTag As String
    sTag = ctl2Set.Tag
    sDataSet sTag, sDataKey, sDataValue
    ctl2Set.Tag = sTag
''    Dim sOutTag As String
''    Dim vTag As Variant
''    Dim vLine As Variant
''    Dim bMatched As Boolean
''    vTag = Split(ctl2Set.Tag, "~")
''    Dim i As Integer
''    For i = 0 To UBound(vTag)
''        If vTag(i) <> "" Then
''            vLine = Split(vTag(i), "`")
''            ' Check For Replace
''            If UCase(Trim(vLine(0))) = UCase(Trim(sDataKey)) Then
''                sOutTag = sOutTag & "~" & sDataKey & "`" & sDataValue
''                bMatched = True
''            Else
''                sOutTag = sOutTag & "~" & vTag(i)
''            End If
''        End If
''    Next
''    If Not bMatched Then
''        sOutTag = sOutTag & "~" & sDataKey & "`" & sDataValue
''    End If
''    sOutTag = ts.sAppend(sOutTag, "~")
''    ctl2Set.Tag = sOutTag
End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataSet
'    Part of the ctlTagData set of functions, this function will allow you set a value
'    in a controls .Tag property.
Public Function sDataSet( _
                            ByRef ctl2Set, _
                            ByVal sDataKey As String, _
                            ByVal sDataValue As String)
                            
    Dim sOutTag As String
    Dim vTag As Variant
    Dim vLine As Variant
    Dim bMatched As Boolean
    vTag = Split(ctl2Set, "~")
    Dim i As Integer
    For i = 0 To UBound(vTag)
        If vTag(i) <> "" Then
            vLine = Split(vTag(i), "`")
            ' Check For Replace
            If UCase(Trim(vLine(0))) = UCase(Trim(sDataKey)) Then
                sOutTag = sOutTag & "~" & sDataKey & "`" & sDataValue
                bMatched = True
            Else
                sOutTag = sOutTag & "~" & vTag(i)
            End If
        End If
    Next
    If Not bMatched Then
        sOutTag = "~" & sDataKey & "`" & sDataValue & sOutTag
    End If
    sOutTag = ts.sAppend(sOutTag, "~")
    ctl2Set = sOutTag
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataGet
'    Part of the ctlTagData set of functions, this function will allow you to read
'    a value out of a control's tag property.
Public Function ctlTagDataGet( _
                                ByRef ctl2GetFrom, _
                                ByVal sDataKey As String) _
                As String
    ctlTagDataGet = sDataGet(ctl2GetFrom.Tag, sDataKey)
'    Dim sReturn As String
'    Dim vTag As Variant
'    Dim vLine As Variant
'    vTag = Split(ctl2GetFrom.Tag, "~")
'    Dim i As Integer
'    For i = 0 To UBound(vTag)
'        If vTag(i) <> "" Then
'            vLine = Split(vTag(i), "`")
'            If UCase(Trim(vLine(0))) = UCase(Trim(sDataKey)) Then
'                sReturn = vLine(1)
'                Exit For
'            End If
'        End If
'    Next
'    ctlTagDataGet = sReturn
End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataGet
'    Part of the ctlTagData set of functions, this function will allow you to read
'    a value out of a control's tag property.
Public Function sDataGet( _
                                ByRef ctl2GetFrom, _
                                ByVal sDataKey As String) _
                As String
                
    Dim sReturn As String
    Dim vTag As Variant
    Dim vLine As Variant
    vTag = Split(ctl2GetFrom, "~")
    Dim i As Integer
    For i = 0 To UBound(vTag)
        If vTag(i) <> "" Then
            vLine = Split(vTag(i), "`")
            If UCase(Trim(vLine(0))) = UCase(Trim(sDataKey)) Then
                sReturn = vLine(1)
                Exit For
            End If
        End If
    Next
    sDataGet = sReturn
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataHas
'    Part of the ctlTagData set of functions, this function will allow you to test to
'    see if a value has been stored into the control's tag property.
Public Function ctlTagDataHas( _
                            ByRef ctl2GetFrom, _
                            ByVal sDataKey As String) _
                As Boolean
    On Error Resume Next
    ctlTagDataHas = sDataHas(ctl2GetFrom.Tag, sDataKey)
''    Dim bHas As Boolean
''    Dim vTag As Variant
''    Dim vLine As Variant
''    vTag = Split(ctl2GetFrom.Tag, "~")
''    Dim i As Integer
''    For i = 0 To UBound(vTag)
''        If vTag(i) <> "" Then
''            vLine = Split(vTag(i), "`")
''            If UCase(Trim(vLine(0))) = UCase(Trim(sDataKey)) Then
''                bHas = True
''                Exit For
''            End If
''        End If
''    Next
''    ctlTagDataHas = bHas
End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataHas
'    Part of the ctlTagData set of functions, this function will allow you to test to
'    see if a value has been stored into the control's tag property.
Public Function sDataHas( _
                            ByRef ctl2GetFrom, _
                            ByVal sDataKey As String) _
                As Boolean
                
    Dim bHas As Boolean
    Dim vTag As Variant
    Dim vLine As Variant
    vTag = Split(ctl2GetFrom, "~")
    Dim i As Integer
    For i = 0 To UBound(vTag)
        If vTag(i) <> "" Then
            vLine = Split(vTag(i), "`")
            If UCase(Trim(vLine(0))) = UCase(Trim(sDataKey)) Then
                bHas = True
                Exit For
            End If
        End If
    Next
    sDataHas = bHas
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataDel
'    Part of the ctlTagData set of functions, this function will allow you to remove a value
'    that has been set into a control's .Tag property.
Public Function ctlTagDataDel( _
                            ByRef ctl2Set, _
                            ByVal sDataKey As String)
    Dim sTag As String
    sTag = ctl2Set.Tag
    sDataDel sTag, sDataKey
    ctl2Set.Tag = sTag
'''    Dim sOutTag As String
'''    Dim vTag As Variant
'''    Dim vLine As Variant
'''    Dim bMatched As Boolean
'''    vTag = Split(ctl2Set.Tag, "~")
'''    Dim i As Integer
'''    For i = 0 To UBound(vTag)
'''        If vTag(i) <> "" Then
'''            vLine = Split(vTag(i), "`")
'''            ' Check For Replace
'''            If UCase(Trim(vLine(0))) <> UCase(Trim(sDataKey)) Then
'''                sOutTag = sOutTag & "~" & vTag(i)
'''            End If
'''        End If
'''    Next
'''    sOutTag = ts.sAppend(sOutTag, "~")
'''    ctl2Set.Tag = sOutTag
'''
End Function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTagDataDel
'    Part of the ctlTagData set of functions, this function will allow you to remove a value
'    that has been set into a control's .Tag property.
Public Function sDataDel( _
                            ByRef sData, _
                            ByVal sDataKey As String)
    
    Dim sOutTag As String
    Dim vTag As Variant
    Dim vLine As Variant
    Dim bMatched As Boolean
    vTag = Split(sData, "~")
    Dim i As Integer
    For i = 0 To UBound(vTag)
        If vTag(i) <> "" Then
            vLine = Split(vTag(i), "`")
            ' Check For Replace
            If UCase(Trim(vLine(0))) <> UCase(Trim(sDataKey)) Then
                sOutTag = sOutTag & "~" & vTag(i)
            End If
        End If
    Next
    sOutTag = ts.sAppend(sOutTag, "~")
    sData = sOutTag
    
End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' bNot
'    This simple little function performs a boolean not
'    so any none zero value passed will come back as zero and
'    any zero value passed will come back as -1.
Public Function bNot( _
                        ByVal vValue As Variant) _
                As Boolean
                
    bNot = Not (-1 * vValue)
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' bAbs
'    This simple little function with return the ABSOLUTE
'    boolean value for a value passed.
'    and non zero value passed will come back as -1
Public Function bAbs( _
                        ByVal vValue As Variant) _
                As Boolean
    
    On Error Resume Next
    bAbs = (-1 * vValue)
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' bOK
'    This function allows you to wrap functions that return a zero when
'    they are successful so you can use them in IF or types of logical
'    statements.
Public Function bOK( _
                    ByVal vValue As Variant) _
                As Boolean
    
    On Error Resume Next
    bOK = (vValue = 0)
        
End Function



'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' formNamed
'    This function will return a reference to a form name you pass.
'    Nothing is returned if the form is not loaded.
Public Function formNamed(ByVal sFormName As String) As Form
    
    Dim i As Integer
    For i = 0 To VB.Forms.Count - 1
        If UCase(VB.Forms(i).Name) = UCase(sFormName) Then
            Set formNamed = VB.Forms(i)
            Exit Function
        End If
    Next i
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' formIsLoaded
'    This function allows you to test to see if a form is loaded (by checking
'    for the form name you pass).
Public Function formIsLoaded(ByVal sFormName As String) As Boolean
    
    Dim i As Integer
    For i = 0 To Forms.Count - 1
        If UCase$(Forms(i).Name) = UCase$(sFormName) Then
            formIsLoaded = True
            Exit Function
        End If
    Next i
    formIsLoaded = False
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlTextSelAll
'    This function is used usually in the GotFocus event of text and combobox controls
'    to automatically select all of the text contained in it.
Public Function ctlTextSelAll(ByRef oControl As Object)
    On Error Resume Next
    oControl.SelStart = 0
    oControl.SelLength = Len(oControl.Text)
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' aGetSection
'    This function will return a portion of a variant single column array, beginning with the item # you pass
'    and continuing for an optional amount of items.  If you don't specify how many items to return, the
'    function will return the rest of the items in the array.
Public Function aGetSection(ByVal vArray As Variant, ByVal lStart As Long, Optional ByVal lCount As Long = -1) As Variant
    
    Dim vOutput As Variant
    If lCount = -1 Then
        lCount = UBound(vArray) - lStart
    End If
    ReDim vOutput(lCount)
    Dim l As Long
    For l = 0 To lCount
        vOutput(l) = vArray(l + lStart)
    Next
    aGetSection = vOutput
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' aAppend
'    This function will append 1 single column variant array onto another single column variant
'    array and return the resulting array.
Public Function aAppend(ByVal vArray1 As Variant, ByVal vArray2 As Variant) As Variant
    
    Dim vOutput As Variant
    ReDim vOutput(UBound(vArray1) + UBound(vArray2) + 1)
    Dim l As Long
    Dim lCnt As Long
    lCnt = UBound(vArray1)
    For l = 0 To lCnt
        vOutput(l) = vArray1(l)
    Next l
    For l = 0 To UBound(vArray2)
        vOutput(lCnt + 1 + l) = vArray2(l)
    Next l
    aAppend = vOutput
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' aPrePend
'    This function will pre-pend a specified string to the beginning of every item in a single column
'    variant array.  The default value of the string to be pre-pended is a vb comment.
Public Function aPrePend(ByVal vArray As Variant, Optional ByVal sText2PrePend As String = "' ") As Variant
    Dim l As Long
    For l = 0 To UBound(vArray)
        vArray(l) = sText2PrePend & vArray(l)
    Next
    aPrePend = vArray
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' shellLaunchDoc
'    This function will launch a specified document in an optionally specified
'    directory path and return a boolean on whether it was successful or not.
Public Function shellLaunchDoc(ByVal DocFileName As String, Optional ByVal PathName As String) As Boolean
    
    Dim lReturn As Long
    Dim lWinHandle As Long
    lWinHandle = GetDesktopWindow()
    If Trim(PathName) = "" Then
        PathName = CurDir
    End If
    
'    lReturn = ShellExecute(Me.hwnd, vbNullString, "assocapps.vbp", vbNullString, CurDir, 1)
    lReturn = ShellExecute(lWinHandle, "Open", DocFileName, "", PathName, SW_SHOWNORMAL)
    
    Dim sError As String
    If lReturn <= 32 Then
        'There was an error
        Select Case lReturn
            Case SE_ERR_FNF
                sError = "File not found"
            Case SE_ERR_PNF
                sError = "Path not found"
            Case SE_ERR_ACCESSDENIED
                sError = "Access denied"
            Case SE_ERR_OOM
                sError = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                sError = "DLL not found"
            Case SE_ERR_SHARE
                sError = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                sError = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                sError = "DDE Time out"
            Case SE_ERR_DDEFAIL
                sError = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                sError = "DDE busy"
            Case SE_ERR_NOASSOC
                sError = "No association for file extension"
            Case ERROR_BAD_FORMAT
                sError = "Invalid EXE file or error in EXE image"
            Case Else
                sError = "Unknown error"
        End Select
        MsgBox sError, vbCritical + vbOKOnly, "File Execute Error"
        shellLaunchDoc = False
    Else
        shellLaunchDoc = True
    End If
        
End Function


Public Function mousePosPixels() As POINTAPI
    Dim tPoint As POINTAPI
    GetCursorPos tPoint
    mousePosPixels = tPoint
End Function

Public Function mousePosTwips() As POINTAPI
    Dim tPoint As POINTAPI
    GetCursorPos tPoint
    tPoint.X = twipsX(tPoint.X)
    tPoint.Y = twipsY(tPoint.Y)
    mousePosTwips = tPoint
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' formMoveAbs
'    This function implements a form's .Move method, but ensures that
'    no part of the form is ever carried off the screen.
Public Function formMoveAbs( _
                                            ByRef frm2Move As Form, _
                                            ByVal X As Long, _
                                            ByVal Y As Long, _
                                            Optional ByVal lWidth, _
                                            Optional ByVal lHeight)
    
    If IsMissing(lWidth) Then
        lWidth = frm2Move.Width
    End If
    If IsMissing(lHeight) Then
        lHeight = frm2Move.Height
    End If
    If X < 0 Then
        X = 0
    End If
    If Y < 0 Then
        Y = 0
    End If
    If X + lWidth > Screen.Width Then
        X = Screen.Width - lWidth
    End If
    If Y + lHeight > Screen.Height Then
        Y = Screen.Height - lHeight
    End If
    frm2Move.Move X, Y, lWidth, lHeight
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sInTrim
'    Returns a string with all redundant spaces reduced to a single space.
'    Alternate character(s) may be specified for the space.
Public Function sInTrim(ByVal InputString As String, Optional ByVal AlternateSpaceCharacter As String = " ") As String
    Dim SpaceIsOn As Boolean
    Dim StringLength As Integer
    Dim i As Integer
    Dim outString As String
    
    outString = ""
    
    InputString = AlternateSpaceCharacter & InputString & AlternateSpaceCharacter
    SpaceIsOn = True
    StringLength = Len(InputString)
    
    For i = 1 To StringLength
        If InStr(AlternateSpaceCharacter, Mid(InputString, i, 1)) > 0 Then
            If Not SpaceIsOn Then
                outString = outString & Mid(InputString, i, 1)
                SpaceIsOn = True
            End If
        Else
            SpaceIsOn = False
            outString = outString & Mid(InputString, i, 1)
        End If
    Next
    
    If InStr(AlternateSpaceCharacter, Right(outString, 1)) > 0 And Not outString = "" Then
        outString = Left(outString, Len(outString) - 1)
    End If
    
    sInTrim = outString

End Function

Public Function fontCopyProps(ByRef oFrom, ByRef oTo)
    
    On Error Resume Next
    With oTo
        .FontBold = oFrom.FontBold
        .FontColor = oFrom.FontColor
        .FontItalic = oFrom.FontItalic
        .FontName = oFrom.FontName
        .FontSize = oFrom.FontSize
        .FontStrikethru = oFrom.FontStrikethru
        .FontTransparent = oFrom.FontTransparent
        .FontUnderline = oFrom.FontUnderline
    End With
    
End Function


Public Function sOccurs( _
                        StringToSearch As String, _
                        StringToFind As String) _
                As Long

    If Len(StringToFind) Then
        sOccurs = UBound(Split(StringToSearch, StringToFind))
    End If
    
End Function

Public Function sSplit( _
                        Expression As String, _
                        Optional Delimiter As String = " ", _
                        Optional Limit As Long = -1, _
                        Optional Compare As VbCompareMethod = vbBinaryCompare) _
                As Variant
    
    Dim vArr As Variant
    Dim lCount As Long
    Dim lSize As Long
    Dim lMatch As Long
    Dim lLastMatch As Long
    lMatch = 1
    lLastMatch = 1
    lCount = -1
    Const lDimIncrement = 500
    lSize = lDimIncrement
    ReDim vArr(lSize)
    
    Do While (Limit = -1 Or lCount < Limit) And lMatch > 0
        lMatch = InStr(lLastMatch, Expression, Delimiter, Compare)
        If lMatch > 0 Then
            lCount = lCount + 1
            If lCount > lSize Then
                lSize = lSize + lDimIncrement
                ReDim Preserve vArr(lSize)
            End If
            vArr(lCount) = Mid(Expression, lLastMatch, lMatch - lLastMatch)
            lLastMatch = lMatch + Len(Delimiter)
        Else
            lCount = lCount + 1
            If lCount > lSize Then
                ReDim Preserve vArr(lCount)
            End If
            vArr(lCount) = Mid(Expression, lLastMatch)
        End If
    Loop
    
    ReDim Preserve vArr(lCount)
    
    sSplit = vArr
    
End Function


Public Function timeGetStart() As Date
    Dim dStart As Date
    dStart = Now
    Do While dStart = Now
    Loop
    timeGetStart = Now
End Function



''Function Splits(sExpression As String, _
''                Optional sDelimiters As String = sWhiteSpace, _
''                Optional cMax As Long = -1) As Variant
''    Dim sToken As String, asRet() As String, c As Long
''    ' Error trap to resize on overflow
''    On Error GoTo SplitResize
''    ' Break into tokens and put in an array
''    sToken = GetToken(sExpression, sDelimiters)
''    Do While sToken <> sEmpty
''        If cMax <> -1 Then If c >= cMax Then Exit Do
''        asRet(c) = sToken
''        c = c + 1
''        sToken = GetToken(sEmpty, sDelimiters)
''    Loop
''    ' Size is an estimate, so resize to counted number of tokens
''    ReDim Preserve asRet(0 To c - 1)
''    Splits = asRet
''    Exit Function
''
''SplitResize:
''    ' Resize on overflow
''    Const cChunk As Long = 20
''    If Err.Number = eeOutOfBounds Then
''        ReDim Preserve asRet(0 To c + cChunk) As String
''        Resume              ' Try again
''    End If
''    ErrRaise Err.Number     ' Other VB error for client
''End Function

Public Function aPackEmpty( _
                            ByRef aVar As Variant) _
                As Variant
    
    Dim lCount As Long
    Dim aPack As Variant
    Dim bEmpty As Boolean
    ReDim aPack(UBound(aVar))
    lCount = -1
    
    Dim l As Long
    For l = 0 To UBound(aVar)
        If aVar(l) <> "" Then
            lCount = lCount + 1
            aPack(lCount) = aVar(l)
        End If
    Next
    
    ReDim Preserve aPack(lCount)
    aVar = aPack
    aPackEmpty = aPack
    
End Function

Public Function sWordCnt( _
                        ByVal sString As String, _
                        Optional ByVal Delimiter As String = " ", _
                        Optional Compare As VbCompareMethod = vbBinaryCompare) _
                As Long
    
    Dim aVar As Variant, lCount As Long, l As Long
    aVar = Split(sString, Delimiter, , Compare)
    For l = 0 To UBound(aVar)
        If aVar(l) <> "" Then
            lCount = lCount + 1
        End If
    Next
    sWordCnt = lCount
    
End Function
                        
Public Function sWord( _
                        ByVal sString As String, _
                        ByVal lWordNo As Long, _
                        Optional ByVal Delimiter As String = " ", _
                        Optional Compare As VbCompareMethod = vbBinaryCompare) _
                As String
    
    Dim aVar As Variant, lCount As Long, l As Long
    aVar = Split(sString, Delimiter, , Compare)
    For l = 0 To UBound(aVar)
        If aVar(l) <> "" Then
            lCount = lCount + 1
        End If
        If lCount = lWordNo Then
            sWord = aVar(l)
            Exit Function
        End If
    Next l
    
End Function



'''Public Function prnGetLine(ByVal MemoText As String, ByVal Lineno As Integer, ByRef PRN As Printer, Optional ByVal HorizontalOffset As Long)
'''
'''    Dim LineCnt As Integer
'''    Dim i As Integer
'''    Dim iWordCnt As Integer
'''    Dim sLine As String
'''    Dim sPrevLine As String
'''    Dim lAvailWidth As Long
'''
'''    If IsMissing(HorizontalOffset) Then
'''        lAvailWidth = PRN.ScaleWidth
'''    Else
'''        lAvailWidth = PRN.ScaleWidth - HorizontalOffset
'''    End If
'''
'''    LineCnt = 0
'''    sLine = ""
'''    iWordCnt = WordCnt(MemoText)
'''
'''    PrinterGetLine = ""
'''
'''    For i = 1 To iWordCnt
'''        If PRN.TextWidth(sLine & Word(MemoText, i)) < lAvailWidth Then
'''            If sLine = "" Then
'''                LineCnt = 1
'''                sLine = Word(MemoText, 1)
'''            Else
'''                sLine = sLine & " " & Word(MemoText, i)
'''            End If
'''        Else
'''            If Lineno = LineCnt Then
'''                Exit For
'''            Else
'''                LineCnt = LineCnt + 1
'''                sLine = Word(MemoText, i)
'''            End If
'''        End If
'''    Next i
'''    If Lineno = LineCnt Then
'''        PrinterGetLine = sLine
'''    End If
'''
'''End Function
'''
'''
'''
'''Public Function prnLineCnt(ByVal MemoText As String, ByRef PRN As Printer, Optional ByVal HorizontalOffset As Long) As Integer
'''
'''    Dim LineCnt As Integer
'''    Dim i As Integer
'''    Dim iWordCnt As Integer
'''    Dim sLine As String
'''    Dim sPrevLine As String
'''    Dim lAvailWidth As Long
'''
'''    If IsMissing(HorizontalOffset) Then
'''        lAvailWidth = PRN.ScaleWidth
'''    Else
'''        lAvailWidth = PRN.ScaleWidth - HorizontalOffset
'''    End If
'''
'''    LineCnt = 0
'''    sLine = ""
'''    iWordCnt = WordCnt(MemoText)
'''
'''    For i = 1 To iWordCnt
'''        If PRN.TextWidth(sLine & Word(MemoText, i)) < lAvailWidth Then
'''            If sLine = "" Then
'''                LineCnt = 1
'''                sLine = Word(MemoText, 1)
'''            Else
'''                sLine = sLine & " " & Word(MemoText, i)
'''            End If
'''        Else
'''            LineCnt = LineCnt + 1
'''            sLine = Word(MemoText, i)
'''        End If
'''    Next i
'''
'''    PrinterLineCnt = LineCnt
'''
'''End Function

Public Function prnRows(Optional ByRef PRN As Printer) As Long
    prnDefault PRN
    prnRows = Int((PRN.ScaleHeight) / (PRN.TextHeight(sCompleteKeyboard)))
    
End Function

Public Function prnCurrentRow( _
                                Optional ByRef PRN As Printer, _
                                Optional bZeroBased As Boolean = False) _
                As Long
    
    prnDefault PRN
    prnCurrentRow = Int((PRN.CurrentY + (PRN.TextHeight(sCompleteKeyboard) / 2)) / (PRN.TextHeight(sCompleteKeyboard))) + IIf(bZeroBased, 0, 1)
    
End Function
Public Function prnDefault( _
                            Optional ByRef PRN As Printer)
    
    If PRN Is Nothing Then
        Set PRN = VB.Printer
    End If
    
End Function

Public Function prnPixelsX( _
                            ByVal lWidth As Long, _
                            Optional ByVal eFrom As ScaleModeConstants = vbTwips, _
                            Optional ByRef PRN As Printer) _
                As Single
    
    prnDefault PRN
    prnPixelsX = ScaleX(PRN, lWidth, eFrom, vbPixels)
    
End Function

Public Function prnPixelsY( _
                            ByVal lHeight As Long, _
                            Optional ByVal eFrom As ScaleModeConstants = vbTwips, _
                            Optional ByRef PRN As Printer) _
                As Single
    
    prnDefault PRN
    prnPixelsY = ScaleY(PRN, lHeight, eFrom, vbPixels)
    
End Function
                
Public Function prnTwipsX( _
                            ByVal lWidth As Long, _
                            Optional ByVal eFrom As ScaleModeConstants = vbPixels, _
                            Optional ByRef PRN As Printer) _
                As Single
    
    prnDefault PRN
    prnTwipsX = ScaleX(PRN, lWidth, eFrom, vbTwips)
    
End Function

Public Function prnTwipsY( _
                            ByVal lHeight As Long, _
                            Optional ByVal eFrom As ScaleModeConstants = vbPixels, _
                            Optional ByRef PRN As Printer) _
                As Single
    
    prnDefault PRN
    prnTwipsY = ScaleY(PRN, lHeight, eFrom, vbTwips)
    
End Function

Public Function ScaleX( _
                        ByRef baseObject, _
                        ByVal sglWidth As Single, _
                        ByVal eFrom As ScaleModeConstants, _
                        ByVal eTo As ScaleModeConstants) _
                As Single
    
    On Error Resume Next
    ScaleX = baseObject.ScaleX(sglWidth, eFrom, eTo)
    
End Function

Public Function ScaleY( _
                        ByRef baseObject, _
                        ByVal sglHeight As Single, _
                        ByVal eFrom As ScaleModeConstants, _
                        ByVal eTo As ScaleModeConstants) _
                As Single
    
    On Error Resume Next
    ScaleY = baseObject.ScaleY(sglHeight, eFrom, eTo)
    
End Function

Public Function prnPrint( _
                        ByVal sText As String, _
                        Optional ByRef PRN As Printer, _
                        Optional ByVal bNoWordWrap As Boolean = False) _
                As String
    
    prnDefault PRN
    
    Dim lLineCnt As Long
    Dim aVar As Variant
    Dim bPageComplete As Boolean
    Dim bOneBigLine As Boolean
    
    aVar = Split(sText, vbCrLf)
    Dim l As Long
    Dim sCurrLine As String
    For l = 0 To UBound(aVar)
        sCurrLine = aVar(l)
        bOneBigLine = False
        Do While PRN.TextWidth(sCurrLine) > PRN.ScaleWidth And Not bPageComplete And Not bNoWordWrap
            Dim aWords As Variant
            aWords = Split(sCurrLine, " ")
            Dim lWordsFit As Long
            Do While PRN.TextWidth(Join(ts.aGetSection(aWords, 0, lWordsFit), " ")) < PRN.ScaleWidth And lWordsFit <= UBound(aWords)
                lWordsFit = lWordsFit + 1
            Loop
            If lWordsFit <> UBound(aWords) Then
                lWordsFit = lWordsFit - 1
            End If
            PrintLine PRN, Join(ts.aGetSection(aWords, 0, lWordsFit), " "), lLineCnt, bPageComplete
            If UBound(aWords) > lWordsFit Then
                sCurrLine = Join(ts.aGetSection(aWords, lWordsFit + 1), " ")
            Else
                sCurrLine = ""
                bOneBigLine = True
            End If
        Loop
        If Not bPageComplete And prnCurrentRow <= prnRows Then
            If Not bOneBigLine Then
                PrintLine PRN, sCurrLine, lLineCnt, bPageComplete
            End If
        Else
            
            ' Return Remainder
            aVar(l) = sCurrLine
            prnPrint = Join(ts.aGetSection(aVar, l), vbCrLf)
            Exit For
        End If
    Next
    
End Function
Private Function PrintLine(PRN As Printer, sString As String, lLineCntr As Long, bPageComplete As Boolean)
    lLineCntr = lLineCntr + 1
    If prnCurrentRow = prnRows Then
        PRN.Print sString
        bPageComplete = True
    Else
        PRN.Print sString
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sFileName
'    This function is used to parse keys peices of info from a
'    filename that is passed into it.
Public Function sFileName( _
                            ByVal sFile As String, _
                            ByVal ePortions As enumFileNameParts) _
                As String
    
    Dim lFirstPeriod As Long, lFirstBackSlash As Long
    lFirstPeriod = InStrRev(sFile, ".")
    lFirstBackSlash = InStrRev(sFile, "\")
    Dim sPath As String, sName As String, sExt As String
    If lFirstBackSlash > 0 Then
        sPath = Left(sFile, lFirstBackSlash)
    End If
    If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
        sExt = Mid(sFile, lFirstPeriod + 1)
        sName = Mid(sFile, lFirstBackSlash + 1, lFirstPeriod - lFirstBackSlash - 1)
    Else
        sName = Mid(sFile, lFirstBackSlash + 1)
    End If
    Dim sRet As String
    If ePortions And efpFilePath Then
        sRet = sRet & sPath
    End If
    If ePortions And efpFileName Then
        sRet = sRet & sName
    End If
    If ePortions And efpFileExt Then
        If sRet <> "" Then
            sRet = sRet & "." & sExt
        Else
            sRet = sRet & sExt
        End If
    End If
    sFileName = sRet
    
End Function



Public Function prnPrintWithHeader( _
                                    ByVal sHeader As String, _
                                    ByVal sText As String, _
                                    Optional ByVal bPrintPageNos As Boolean = True, _
                                    Optional ByVal PRN As Printer)
    
    prnDefault PRN
    Do While sText <> ""
        Dim sLine As String
        If bPrintPageNos Then
            sLine = "Pg: " & PRN.Page & "   "
        End If
        sLine = sLine & sHeader
        prnPrint sLine, PRN, True
        prnPrint String(255, "="), PRN, True
        sText = prnPrint(sText, PRN)
        If sText <> "" Then
            PRN.NewPage
        End If
        DoEvents
    Loop
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' lstSumSelected
'    This function is useful for using a standard
'    listbox control as an enumerated list where
'    the list is setup as:
'    itemdata1 = 2 ^ 0
'    itemdata2 = 2 ^ 1
'    itemdata3 = 2 ^ 2
'    etc.
Public Function listSumSelected( _
                                ByRef lstControl As ListBox) _
                As Double
    
    Dim l As Long
    Dim lSum As Long
    For l = 0 To lstControl.ListCount - 1
        If lstControl.Selected(l) Then
            lSum = lSum + forceDouble(lstControl.ItemData(l))
        End If
    Next l
    listSumSelected = lSum
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' lstApplySelect
'    This function is useful for using a standard
'    listbox control as an enumerated list where
'    the list is setup as:
'    itemdata1 = 2 ^ 0
'    itemdata2 = 2 ^ 1
'    itemdata3 = 2 ^ 2
'    etc.
Public Function listApplySelected( _
                                ByRef lstControl As ListBox, _
                                ByVal dAmount As Double)
    
    Dim l As Long
    For l = 0 To lstControl.ListCount - 1
        If (forceDouble(lstControl.ItemData(l)) And dAmount) <> 0 Then
            lstControl.Selected(l) = True
        Else
            lstControl.Selected(l) = False
        End If
    Next l
    
    On Error Resume Next
    lstControl.ListIndex = 0
    lstControl.ListIndex = -1
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' shell16bitFileName
'    This function takes in a 32 bit "long" filename and returns
'    a compatible 16 bit version of it.
Public Function shell16bitFileName(ByVal LongFileName As String) As String
    
    Dim sReturn  As String
    sReturn = Space(Len(LongFileName))
    
    Dim lReturnLen As Long
    lReturnLen = GetShortPathName(LongFileName, sReturn, Len(sReturn))
    sReturn = Left(sReturn, lReturnLen)
    If Len(Trim(sReturn)) = 0 Then
        sReturn = LongFileName
    End If
    shell16bitFileName = sReturn
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlFlush
'    This function will force the changes in a control
'    to be written into the underlying recordset.
Public Function ctlFlush( _
                        ByRef oControl As Control)
                        
    On Error Resume Next
    Dim bEnabled As Boolean
    Dim lSelStart As Long, lSelLength As Long
    bEnabled = oControl.Enabled
    lSelStart = oControl.SelStart
    lSelLength = oControl.SelLength
    oControl.Enabled = False
    oControl.Enabled = bEnabled
    oControl.SetFocus
    DoEvents
    oControl.SelStart = lSelStart
    oControl.SelLength = lSelLength
    oControl.SetSelection 0, 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' menuPopup
'    This function enhances the .PopupMenu method of forms by
'    also ensuring that the popup menu in question is enabled.
Public Function menuPopup( _
                        ByRef SourceObject, _
                        ByRef MenuObject)
    
    On Error Resume Next
    If MenuObject.Enabled Then
        SourceObject.PopupMenu MenuObject
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' formHeaderHeight
'    This function will return the "header height" of any VB form.
'    This differs from the simple metrics of a border and a
'    caption because this takes into account the size of the menu
'    bars on the form too.
Public Function formHeaderHeight( _
                                                    FormToCheck As Form) _
                        As Long
    
    Dim iSaveScaleMode As Integer
    iSaveScaleMode = FormToCheck.ScaleMode
    If FormToCheck.ScaleMode <> vbTwips Then
        FormToCheck.ScaleMode = vbTwips
    End If
    formHeaderHeight = FormToCheck.Height - FormToCheck.ScaleHeight
    If FormToCheck.ScaleMode <> iSaveScaleMode Then
        FormToCheck.ScaleMode = iSaveScaleMode
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sysMetrics
'    This function provides access to the metrics of the Windows
'    system.  See the enum for a possible list of return values.
Public Function sysMetrics( _
                                        ByVal lMetric As enumSystemMetrics) _
                        As Long
    
'    On Error Resume Next
    sysMetrics = GetSystemMetrics(lMetric)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' windowPlacement
'    This function will return the "placement" of any window.
'    Placement includes additional info besides just coords.
Public Function WINDOWPLACEMENT( _
                                                ByVal hWnd As Long) _
                        As mAPIconstants.WINDOWPLACEMENT
    
    Dim Ret As mAPIconstants.WINDOWPLACEMENT
'    On Error Resume Next
    GetWindowPlacement hWnd, Ret
    WINDOWPLACEMENT = Ret
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' keyState
'    This function will determine the state of any virtual key
'    on the keyboard.  Possible states are:
'       .Pressed
'       .Toggled
'    Toggled applies to keys suchs as the CapsLock or NumLock
'    keys which hold a toggled setting on the keyboard.
Public Function keyState( _
                                    ByVal key As enumVIrtualKeys) _
                        As typeKeyState
    
    Dim iRet As Integer
    iRet = GetKeyState(key)
    Dim Ret As typeKeyState
    Ret.Pressed = (iRet And 128) <> 0
    Ret.Toggled = (iRet And 1) <> 0
    keyState = Ret
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' keyShiftState
'    This function will calculate the "shift" state typically returned through the
'    KeyPress event from anywhere and any point in code.
Public Function keyShiftState() _
                        As ShiftConstants
    
    Dim Ret As ShiftConstants
    If keyState(VK_CONTROL).Pressed Then
        Ret = Ret + vbCtrlMask
    End If
    If keyState(VK_SHIFT).Pressed Then
        Ret = Ret + vbShiftMask
    End If
    If keyState(VK_MENU).Pressed Then
        Ret = Ret + vbAltMask
    End If
    keyShiftState = Ret
    
End Function



Public Sub formGradient( _
                                    ByRef frmWork As Form, _
                                    ByVal FromColor As enumColors, _
                                    ByVal ToColor As enumColors)
    
    Static bDoingGradientBackGroundFill As Boolean
    
    If Not bDoingGradientBackGroundFill And frmWork.ScaleHeight > 0 And frmWork.WindowState <> vbMinimized Then
        
        bDoingGradientBackGroundFill = True
        
        Dim intLoop As Integer
        Dim bSaveAutoRedraw As Boolean
        Dim iDrawMode As Integer
        Dim iDrawStyle As Integer
        Dim lDrawWidth As Long
        Dim lScaleHeight As Long
        Dim iScaleMode As Integer
        
        With frmWork
            bSaveAutoRedraw = .AutoRedraw
            iDrawStyle = .DrawStyle
            iDrawMode = .DrawMode
            iScaleMode = .ScaleMode
            lDrawWidth = .DrawWidth
            lScaleHeight = .ScaleHeight
            
            .AutoRedraw = True
            .DrawStyle = vbInsideSolid
            .DrawMode = vbCopyPen
            .ScaleMode = vbPixels
            .DrawWidth = 2
            .ScaleHeight = 256
        End With
        
        Dim iFromRed, iFromBlue, iFromGreen As Integer
        Dim iToRed, iToBlue, iToGreen As Integer
        Dim iCurrRed, iCurrBlue, iCurrGreen As Integer
        iFromRed = colorRGBValue(FromColor, rgbRed)
        iFromBlue = colorRGBValue(FromColor, rgbBlue)
        iFromGreen = colorRGBValue(FromColor, rgbGreen)
        iToRed = colorRGBValue(ToColor, rgbRed)
        iToBlue = colorRGBValue(ToColor, rgbBlue)
        iToGreen = colorRGBValue(ToColor, rgbGreen)
        
        For intLoop = 0 To 255
            iCurrRed = iFromRed - ((iFromRed - iToRed) * (intLoop / 255))
            iCurrGreen = iFromGreen - ((iFromGreen - iToGreen) * (intLoop / 255))
            iCurrBlue = iFromBlue - ((iFromBlue - iToBlue) * (intLoop / 255))
            frmWork.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(iCurrRed, iCurrGreen, iCurrBlue), BF
        Next
        
        frmWork.Refresh
        
        With frmWork
            .AutoRedraw = bSaveAutoRedraw
            .DrawStyle = iDrawStyle
            .DrawMode = iDrawMode
            .DrawWidth = lDrawWidth
            .ScaleMode = iScaleMode
        End With
        
        bDoingGradientBackGroundFill = False
    End If
    
End Sub

Public Function colorRGBValue( _
                                        ByVal lColorToGetFrom As enumColors, _
                                        ByVal iValueToGet As enumRGBbases) _
                        As Byte
    
    Dim Ret As Byte
    Select Case True
        Case iValueToGet = rgbBlue
            Ret = nLoByte(nHiWord(lColorToGetFrom))
        Case iValueToGet = rgbGreen
            Ret = nHiByte(nLoWord(lColorToGetFrom))
        Case iValueToGet = rgbRed
            Ret = nLoByte(nLoWord(lColorToGetFrom))
    End Select
    colorRGBValue = Ret
    
End Function


Public Function nLoByte( _
                                    inval As Integer) _
                        As Byte
    
    nLoByte = nWord(inval).LoByte
    
End Function

Public Function nHiByte( _
                                    inval As Integer) _
                        As Byte
    
    nHiByte = nWord(inval).HiByte
    
End Function


Public Function nWord( _
                                inval As Integer) _
                        As TwoBytes
    
    Dim oWord As OneWord
    oWord.Word = inval
    Dim oBytes As TwoBytes
    LSet oBytes = oWord
    nWord = oBytes
    
End Function

Public Function nDWordBytes( _
                            inval As Long) _
                As FourBytes
    
    Dim oDWord As OneDWord
    Dim oWords As FourBytes
    
    oDWord.dWord = inval
    LSet oWords = oDWord
    nDWordBytes = oWords
                    
End Function

Public Function nDWord( _
                                inval As Long) _
                        As TwoWords
    
    Dim oDWord As OneDWord
    Dim oWords As TwoWords
    
    oDWord.dWord = inval
    LSet oWords = oDWord
    nDWord = oWords
    
End Function

Public Function nGetBit( _
                                inval, _
                                ByVal ByteNumber As Integer) _
                        As Byte
    
    ByteNumber = ByteNumber - 1
    Dim lCheck As Long
    lCheck = 2 ^ ByteNumber
    nGetBit = Abs((inval And lCheck) <> 0)
    
End Function


Public Function formCenterInMDI( _
                                                ByRef ChildFormToCenter As Form, _
                                                ByRef ParentFormToCenterIn As MDIForm)
    On Error Resume Next
    ChildFormToCenter.Move (ParentFormToCenterIn.ScaleWidth / 2) - (ChildFormToCenter.Width / 2), (ParentFormToCenterIn.ScaleHeight / 2) - (ChildFormToCenter.Height / 2)
    On Error GoTo 0
    
End Function



Public Function pixelsX(ByVal TwipsIn As Long) As Integer
    pixelsX = TwipsIn / Screen.TwipsPerPixelX
End Function

Public Function pixelsY(ByVal TwipsIn As Long) As Integer
    pixelsY = TwipsIn / Screen.TwipsPerPixelY
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sInLine
'    This function will return the line # that contains the specified
'    string.  Word wrapping is not in effect, this just parses on
'    the carriage returns.
Public Function sInLine( _
                                ByVal TextToSearch As String, _
                                ByVal StringToSearchFor As String, _
                                Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) _
            As Long
    
    Dim aLines As Variant
    aLines = Split(TextToSearch, vbCrLf)
    Dim l As Long
    For l = 0 To UBound(aLines)
        If InStr(1, aLines(l), StringToSearchFor, Compare) > 0 Then
            sInLine = l + 1
            Exit Function
        End If
    Next
    
End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sWordBF
'    Returns a specific word from a string.  Default word identifier is a space.
'    Optional character or characters can be used in place of space.
'    This function is unique from the Split() command in that words are
'    identified regardless of how many spaces are in between them.
'    The BF stands for Brute Force.  This function was developed in VB5
'    and in certain cases such as LARGE text files, actually runs faster
'    than the equivalent sWord which relies on the Split() function.
Public Function sWordBF( _
                        ByVal InputString As String, _
                        ByVal WordNum As Integer, _
                        Optional ByVal AlternateSpaceCharacter As String = " ", _
                        Optional ByVal bMultipleKindsOfSpaceChars As Boolean = False, _
                        Optional ByVal bIgnoreCase As Boolean = True) _
                As String
    
    Dim SpaceIsOn As Boolean
    Dim StringLength As Long
    Dim i As Long
    Dim outString As String
    Dim CurrWord As Integer
    Dim iMatchLen As Integer
    If bMultipleKindsOfSpaceChars Then
        iMatchLen = 1
    Else
        iMatchLen = Len(AlternateSpaceCharacter)
    End If
    
    outString = ""
    CurrWord = 0
    
    Dim sWorkString As String
    InputString = AlternateSpaceCharacter & InputString & AlternateSpaceCharacter
    sWorkString = InputString
    'SpaceIsOn = True
    StringLength = Len(sWorkString)
   
    If bIgnoreCase Then
        AlternateSpaceCharacter = UCase(AlternateSpaceCharacter)
        sWorkString = UCase(sWorkString)
    End If
    
    For i = 1 To StringLength
        If InStr(AlternateSpaceCharacter, Mid(sWorkString, i, iMatchLen)) > 0 Then
            SpaceIsOn = True
            i = i + iMatchLen - 1
        Else
            If SpaceIsOn Then
                SpaceIsOn = False
                CurrWord = CurrWord + 1
                If CurrWord > WordNum Then
                    Exit For
                End If
            End If
            If CurrWord = WordNum Then
                outString = outString & Mid(InputString, i, 1)
            End If
        End If
    Next i
    
    If Trim(outString) <> "" And InStr(AlternateSpaceCharacter, Right(outString, iMatchLen)) > 0 Then
        outString = Left(outString, Len(outString) - iMatchLen)
    End If
    
    sWordBF = outString
    
End Function


Public Function sCleanSplit( _
                            ByVal Expression As String, _
                            Optional ByVal Delimiter As String = " ", _
                            Optional ByVal Limit As Long = -1, _
                            Optional Compare As VbCompareMethod = vbBinaryCompare) _
                As Variant
    
    Dim varItems As Variant, i As Long
    
    varItems = Split(Expression, Delimiter, Limit, Compare)
     
    For i = 0 To UBound(varItems)
        If Len(varItems(i)) = 0 Then varItems(i) = Delimiter
    Next
    
    sCleanSplit = Filter(varItems, Delimiter, False)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sLineCnt
'     ==================================================
'     LineCnt()
'    Returns the total number of lines in a multi-line string.
'     ==================================================
Public Function sLineCnt( _
                        ByVal StringToCountFrom As String) _
                As Integer
    
    sLineCnt = sWordCnt(StringToCountFrom, vbCrLf)
    
End Function


Public Function windowCoords( _
                            ByVal hWnd As Long) _
                As Rect
                   
    Dim Ret As Rect
    GetWindowRect hWnd, Ret
    windowCoords = Ret
    
End Function



Public Function sProper(ByVal txt As String) As String
    
    Dim need_cap As Boolean
    Dim i As Integer
    Dim ch As String
    
    txt = LCase(txt)
    need_cap = True
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ'", UCase(ch)) > 0 Then
            If need_cap Then
                Mid$(txt, i, 1) = UCase(ch)
                need_cap = False
            End If
        Else
            need_cap = True
        End If
    Next i
    sProper = txt
    
End Function


Public Function sIs( _
                    ByVal sStringToCheck As String, _
                    ByVal WhatToCheck As enumStringIsTypes) _
                As Boolean
    
    Dim sCheck As String
    Dim i As Integer
    Select Case True
        Case WhatToCheck = esiOnlyNumbers
            sCheck = " 0123456789"
        Case WhatToCheck = esiMathematicalCharacters
            sCheck = " 0123456789./-+=*"
        Case WhatToCheck = esiNumbersAndMoneyPunctuation
            sCheck = " 0123456789.$,-"
        Case WhatToCheck = esiNumbersAndNumericPunctuation
            sCheck = " 0123456789.-"
    End Select
    
    sIs = True
    For i = 1 To Len(sStringToCheck)
        If Not InStr(sCheck, Mid(sStringToCheck, i, 1)) > 0 Then
            sIs = False
            Exit For
        End If
    Next i
End Function


Public Function sEquals( _
                        String1 As Variant, _
                        String2 As Variant) _
                As Boolean
    sEquals = Trim(UCase(CStr(String1))) = Trim(UCase(CStr(String2)))
End Function


Public Function fileUniqueName( _
                                ByVal sCurrFile As String) _
                As String

    Dim i As Integer
    Dim iExt As Integer
    Dim sExt As String
    Dim sBase As String
    i = 1
    iExt = InStrRev(sCurrFile, ".")
    If iExt = 0 Then
        sExt = ""
        sBase = sCurrFile
    Else
        sExt = Mid(sCurrFile, iExt)
        sBase = Left(sCurrFile, iExt - 1)
    End If
    Do While fileExists(sCurrFile)
        i = i + 1
        sCurrFile = sBase & "(" & Trim(i) & ")" & sExt
        
    Loop
    fileUniqueName = sCurrFile
    
End Function


Public Function ctlApplyMinMax( _
                                ByRef oProgressBar, _
                                ByVal lMin As Long, _
                                ByVal lMax As Long)
    
    On Error Resume Next
    Dim lCurMin As Long
    Dim lCurMax As Long
    lCurMin = oProgressBar.Min
    lCurMax = oProgressBar.max
    
    If lMin > lCurMax Then
        oProgressBar.max = lMax
        oProgressBar.Min = lMin
    Else
        oProgressBar.Min = lMin
        oProgressBar.max = lMax

    End If
        
End Function


Public Function fileWaitUntilExists( _
                                    ByVal sFileToLookFor As String, _
                                    Optional ByVal iTimeout As Integer = 15) _
                As Boolean
    
    Dim iCounter As Integer
    Do While Not fileExists(sFileToLookFor) And iCounter < iTimeout
        DoEvents
        iCounter = iCounter + 1
        
    Loop
    fileWaitUntilExists = fileExists(sFileToLookFor)
    
End Function



Public Function sNTrim( _
                        ByVal vValue As Variant) _
                As String
    sNTrim = Trim(vValue & "")
End Function

Public Function sNRTrim( _
                        ByVal vValue As Variant) _
                As String
    sNRTrim = RTrim(vValue & "")
End Function

Public Function sRemoveQ( _
                        ByVal sStringToUnquote As String) _
                As String
    ' This function will remove quotes that are wrapped around a text string you want.
    
    Dim sQuoteType As String
    Select Case Left(Trim(sStringToUnquote), 1)
        Case """"
            sQuoteType = """"
        Case "'"
            sQuoteType = "'"
    End Select
    If sQuoteType <> "" Then
        sStringToUnquote = sTrimChars(sStringToUnquote, sQuoteType)
    End If
    
    sRemoveQ = sStringToUnquote
    
End Function

Public Function sWord2Asc( _
                        ByVal sWord As String) _
                As Variant
    ' Converts a byte WORD into it's corresping value
    '  i.e.  chr(0) & chr(67) = 67
    '        chr(3) & chr(31) = 799
    Dim vReturn As Variant
    Dim dBase As Double
    Dim iCharVal As Integer
    dBase = 1
    Dim i As Integer
    For i = Len(sWord) To 1 Step -1
        iCharVal = Asc(Mid(sWord, i, 1))
        vReturn = vReturn + (iCharVal * dBase)
        dBase = dBase * 256
    Next i
    sWord2Asc = vReturn
    
End Function

Public Function sAsc2Word( _
                        ByVal vAsc As Variant) _
                As String
    
    Dim dBase As Double
    Dim vMask As Variant
    Dim dChar As Variant
    Dim sRet As String
    dBase = 1
    Do While vAsc <> 0
        sRet = Chr(vAsc And 255) & sRet
        vAsc = Int(vAsc / 256)
    Loop
    sAsc2Word = sRet
End Function

Public Function nHexVal( _
                        ByVal HexString As String) _
                As Variant
    
    Dim l As Long
    Dim sChar As String
    Dim sReturn As String
    sReturn = "0"
    For l = 1 To Len(HexString)
        sChar = Mid(HexString, Len(HexString) - l + 1, 1)
'        sChar = sPadR(sChar, l, "0")
        sReturn = sReturn + CLng("&H" & sChar) * ((16 ^ (l - 1)))
    Next l
    nHexVal = sReturn
    
End Function

Public Function sRTrim( _
                        ByVal sTrimMe As String, _
                        Optional ByVal sChars2Trim As String = " ") _
                As String
    Do While Right(sTrimMe, Len(sChars2Trim)) = sChars2Trim
        sTrimMe = Left(sTrimMe, Len(sTrimMe) - Len(sChars2Trim))
    Loop
    sRTrim = sTrimMe
End Function


Public Function optSelected( _
                        ByRef ctlOptions, _
                        Optional ByVal iDefault As Integer = -1) _
                As Integer
    
    optSelected = iDefault
    On Error Resume Next
    Dim i As Integer
    For i = 0 To ctlOptions.Count - 1
        If ctlOptions(i).Value <> 0 Then
            optSelected = i
            Exit Function
        End If
    Next i
    
End Function

Public Function bGetSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, Optional ByVal DefaultValue As Boolean) As Boolean
    bGetSetting = forceBool(GetSetting(appname, Section, key, CStr(DefaultValue)))
End Function
Public Function iGetSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, Optional ByVal DefaultValue As Integer) As Integer
    iGetSetting = forceInt(GetSetting(appname, Section, key, CStr(DefaultValue)))
End Function
Public Function lGetSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, Optional ByVal DefaultValue As Long) As Long
    lGetSetting = forceLong(GetSetting(appname, Section, key, CStr(DefaultValue)))
End Function
Public Function dGetSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, Optional ByVal DefaultValue As Date) As Date
    dGetSetting = forceDate(GetSetting(appname, Section, key, DefaultValue))
End Function

Public Function bSaveSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, ByVal Value As Boolean) As Boolean
    SaveSetting appname, Section, key, CStr(Value)
End Function
Public Function iSaveSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, ByVal Value As Integer) As Integer
    SaveSetting appname, Section, key, CStr(Value)
End Function
Public Function lSaveSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, ByVal Value As Long) As Long
    SaveSetting appname, Section, key, CStr(Value)
End Function
Public Function dSaveSetting(ByVal appname As String, ByVal Section As String, ByVal key As String, ByVal Value As Date) As Long
    SaveSetting appname, Section, key, Value
End Function


Public Function forceLong(ByVal InputThing As Variant) As Long
    
    On Error GoTo BadInput
    forceLong = CLng(InputThing)
    Exit Function
    
BadInput:
    forceLong = 0

End Function

Public Function forceSingle(ByVal InputThing As Variant) As Single
    
    On Error GoTo BadInput
    forceSingle = CSng(InputThing)
    Exit Function
    
BadInput:
    forceSingle = 0

End Function

Public Function forceDouble(ByVal InputThing As Variant) As Double
    
    On Error GoTo BadInput
    forceDouble = CDbl(InputThing)
    Exit Function
    
BadInput:
    forceDouble = 0
    
End Function

Public Function forceCurrency(ByVal InputThing As Variant) As Currency
    
    On Error GoTo BadInput
    forceCurrency = CCur(InputThing)
    Exit Function
    
BadInput:
    forceCurrency = 0
    
End Function



Public Function forceInt(ByVal InputThing As Variant) As Integer

    On Error GoTo BadInput
    forceInt = CInt(InputThing)
    Exit Function
    
BadInput:
    forceInt = 0

End Function


Public Function forceStr(ByVal InputThing As Variant) As String

    On Error GoTo BadInput
    forceStr = CStr(InputThing)
    Exit Function
    
BadInput:
    forceStr = ""

End Function

Public Function forceBool(ByVal InputThing As Variant) As Boolean
    
    Select Case VarType(InputThing)
        Case vbString
            forceBool = InStr("~Y~YES~TRUE~+~POSITIVE~T~-1~1~", "~" & UCase(Trim(InputThing)) & "~") > 0
        Case Else
            On Error GoTo BadInput
            forceBool = CBool(InputThing)
    End Select
    Exit Function
    
BadInput:
    forceBool = False
    
End Function

Public Function forceDate(ByVal InputThing As Variant, Optional ByVal bEmptyOnError As Boolean = True) As Variant
    
    On Error GoTo BadInput
    forceDate = CDate(InputThing)
    Exit Function

BadInput:
    If bEmptyOnError Then
        forceDate = Empty
    Else
        forceDate = #1/1/100#
    End If
    
End Function

Public Function forceVariant(ByVal InputThing As Variant) As Variant
    
    On Error GoTo BadInput
    forceVariant = CVar(InputThing)
    Exit Function

BadInput:
    
    
End Function


Public Sub timePause( _
                    ByVal Seconds As Double)
    
    Dim iHoldTime As Single
    iHoldTime = Timer
    While Timer < iHoldTime + Seconds
        DoEvents
    Wend
    
End Sub


' ==============================================
' KillFile()
'     returns a logical on whether it was able to successfully
'     delete the file you passed.
' ==================================================
Public Function fileKill(ByVal File2Delete As String) As Boolean

    On Error GoTo DidNotKill
    Kill File2Delete
    fileKill = True
    Exit Function
    
DidNotKill:
    fileKill = False
    
End Function


' ==============================================
' ObjectIsSet()
'     Returns whether the object you have passed has been set yet or not.
' ==================================================
Public Function objIsSet( _
                        ByRef ObjToCheck As Object) _
                As Boolean
    
    Dim iTester As Integer
    On Error GoTo IsNotSet
    objIsSet = Not (ObjToCheck Is Nothing)
    On Error GoTo 0
    'ObjectIsSet = True
    Exit Function
    
IsNotSet:
    objIsSet = False

End Function


' ==================================================
' Replicate()
'    Returns a string composed of the string passed repeated X number of times.
' ==================================================
Function sReplicate( _
                    ByVal CharacterOrString As String, _
                    ByVal Times2Repeat As Integer) _
        As String

    Dim ReturnString As String
    Dim i As Integer
    ReturnString = ""
    
    For i = 1 To Times2Repeat
        ReturnString = ReturnString & CharacterOrString
        
    Next
    sReplicate = ReturnString
    
End Function


Public Function fileRead( _
                        ByVal FileNameToRead As String) _
                As String
    
    Dim fhand As Integer
    Dim ReturnStr As String
    Dim CurrChar As String
    CurrChar = Space(1024 * 8)
    ReturnStr = ""
    fhand = FreeFile()
    On Error Resume Next
    Open FileNameToRead For Binary Access Read As #fhand
    If Err.Number = 0 Then
        On Error GoTo 0
        
        Do While Not EOF(fhand)
            DoEvents
            Get fhand, , CurrChar
            ReturnStr = ReturnStr & CurrChar
        Loop
        
    End If
    Close #fhand
    On Error GoTo 0
    fileRead = ReturnStr
    
End Function


Public Function dirSet( _
                        ByVal sPathName As String) _
                As Boolean
                
    On Error Resume Next
    ChDrive sPathName
    ChDir sPathName
    dirSet = Err.Number = 0
    On Error GoTo 0
    

End Function


Public Function dirKill( _
                        ByVal sDirectory As String, _
                        Optional ByRef StatusObject As Object) _
                As Boolean
    
    Dim sFileName As String
    Dim Files As Collection
    Dim i As Integer
    
    On Error GoTo DeleteError
    
    ' Get a list of files it contains.
    Set Files = New Collection
    sFileName = Dir$(sDirectory & "\*.*", vbReadOnly + vbHidden + vbSystem + vbDirectory)
    Do While Len(sFileName) > 0
        If (sFileName <> "..") And (sFileName <> ".") Then
            Files.Add sDirectory & "\" & sFileName
        End If
        sFileName = Dir$()
    Loop
    
    ' Delete the files.
    For i = 1 To Files.Count
        
        DoEvents
        
        sFileName = Files(i)
        ' See if it is a directory.
        If GetAttr(sFileName) And vbDirectory Then
            ' It is a directory. Delete it.
            dirKill sFileName
        Else
            ' It's a file. Delete it.
            On Error Resume Next
            StatusObject.Caption = sDirectory
            StatusObject.Refresh
            On Error GoTo DeleteError
            SetAttr sFileName, vbNormal
            Kill sFileName
        End If
    Next i
    
    ' The directory is now empty. Delete it.
    On Error Resume Next
    StatusObject.Caption = sDirectory
    StatusObject.Refresh
    On Error GoTo DeleteError
    RmDir sDirectory
    
    dirKill = True
    Set Files = Nothing
    Exit Function
    
DeleteError:
    
    dirKill = False
    On Error Resume Next
    Set Files = Nothing
    On Error GoTo 0
    
End Function

Public Function fileRename( _
                                    ByVal sFrom As String, _
                                    ByVal sTO As String) _
                        As Boolean
    
    On Error Resume Next
    Name sFrom As sTO
    fileRename = (Err.Number = 0)
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nPercent
'    This simple function looks almost useless, but it incorporates an error trap for those
'    situations where you are trying to calculate a percentage with a base value of 0.
'    (i.e.  divide by zero error)
Public Function nPercent( _
                                    ByVal Amount, _
                                    ByVal max) _
                        As Integer
    
    On Error Resume Next
    nPercent = CInt((Amount / max) * 100)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' formByHWnd
'    Return a VB form object based open a window handle
'    you pass.
Public Function formByHWnd( _
                                        ByVal hWnd As Long) _
                        As Form
                        
    Dim fCurr As Form
    For Each fCurr In Forms
        If fCurr.hWnd = hWnd Then
            Set formByHWnd = fCurr
            Exit Function
        End If
    Next fCurr
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nLowest
'    Take in a list of up to 9 numbers and return the
'    lowest number in the list.
Public Function nLowest( _
                                    Optional ByVal lValue1, _
                                    Optional ByVal lValue2, _
                                    Optional ByVal lValue3, _
                                    Optional ByVal lValue4, _
                                    Optional ByVal lValue5, _
                                    Optional ByVal lValue6, _
                                    Optional ByVal lValue7, _
                                    Optional ByVal lValue8, _
                                    Optional ByVal lValue9) _
                        As Double
    
    Dim dLowest As Double
    Const cMax = 1.84467440737096E+19
    dLowest = cMax
    
    If Not IsMissing(lValue1) Then
        If lValue1 < dLowest Then
            dLowest = lValue1
        End If
    End If
    
    If Not IsMissing(lValue2) Then
        If lValue2 < dLowest Then
            dLowest = lValue2
        End If
    End If
    
    If Not IsMissing(lValue3) Then
        If lValue3 < dLowest Then
            dLowest = lValue3
        End If
    End If
    
    If Not IsMissing(lValue4) Then
        If lValue4 < dLowest Then
            dLowest = lValue4
        End If
    End If
    
    If Not IsMissing(lValue5) Then
        If lValue5 < dLowest Then
            dLowest = lValue5
        End If
    End If
    
    If Not IsMissing(lValue6) Then
        If lValue6 < dLowest Then
            dLowest = lValue6
        End If
    End If
    
    If Not IsMissing(lValue7) Then
        If lValue7 < dLowest Then
            dLowest = lValue7
        End If
    End If
    
    If Not IsMissing(lValue8) Then
        If lValue8 < dLowest Then
            dLowest = lValue8
        End If
    End If
    
    If Not IsMissing(lValue9) Then
        If lValue9 < dLowest Then
            dLowest = lValue9
        End If
    End If
    
    If dLowest = cMax Then
        dLowest = 0
    End If
    
    nLowest = dLowest
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nHighest
'    Take in a list of up to 9 numbers and return the highest
'    one.
Public Function nHighest( _
                                    Optional ByVal lValue1, _
                                    Optional ByVal lValue2, _
                                    Optional ByVal lValue3, _
                                    Optional ByVal lValue4, _
                                    Optional ByVal lValue5, _
                                    Optional ByVal lValue6, _
                                    Optional ByVal lValue7, _
                                    Optional ByVal lValue8, _
                                    Optional ByVal lValue9) _
                        As Double

    Dim dHighest As Double
    Const cMax = -1.84467440737096E+19
    dHighest = cMax
     
    If Not IsMissing(lValue1) Then
        If lValue1 > dHighest Then
            dHighest = lValue1
        End If
    End If

    If Not IsMissing(lValue2) Then
        If lValue2 > dHighest Then
            dHighest = lValue2
        End If
    End If

    If Not IsMissing(lValue3) Then
        If lValue3 > dHighest Then
            dHighest = lValue3
        End If
    End If

    If Not IsMissing(lValue4) Then
        If lValue4 > dHighest Then
            dHighest = lValue4
        End If
    End If

    If Not IsMissing(lValue5) Then
        If lValue5 > dHighest Then
            dHighest = lValue5
        End If
    End If

    If Not IsMissing(lValue6) Then
        If lValue6 > dHighest Then
            dHighest = lValue6
        End If
    End If

    If Not IsMissing(lValue7) Then
        If lValue7 > dHighest Then
            dHighest = lValue7
        End If
    End If

    If Not IsMissing(lValue8) Then
        If lValue8 > dHighest Then
            dHighest = lValue8
        End If
    End If

    If Not IsMissing(lValue9) Then
        If lValue9 > dHighest Then
            dHighest = lValue9
        End If
    End If

    If dHighest = cMax Then
        dHighest = 0
    End If
    
    nHighest = dHighest

End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' aQuickSort
'    '== Sort a 2 dimensional array on SortField
'    '==
'    '== Quicksort is the fastest array sorting routine for
'    '== unordered arrays.  Its big O is  n log n
'    '==
'    '== Parameters:
'    '== vec       - array to be sorted
'    '== SortField - The field to sort on (2nd dimension value)
'    '== loBound and hiBound are simply the upper and lower
'    '== bounds of the array's 1st dimension.  It's probably
'    '== easiest to use the LBound and UBound functions to
'    '== set these.
Public Sub aQuickSort( _
                                vec, _
                                loBound, _
                                hiBound, _
                                SortField)
 

  Dim pivot(), loSwap, hiSwap, temp, counter
  ReDim pivot(UBound(vec, 2))

  '== Two items to sort
  If hiBound - loBound = 1 Then
    If vec(loBound, SortField) > vec(hiBound, SortField) Then
        Call aSwapRows(vec, hiBound, loBound)
    End If
  End If

  '== Three or more items to sort
  
  For counter = 0 To UBound(vec, 2)
    pivot(counter) = vec(Int((loBound + hiBound) / 2), counter)
    vec(Int((loBound + hiBound) / 2), counter) = vec(loBound, counter)
    vec(loBound, counter) = pivot(counter)
  Next

  loSwap = loBound + 1
  hiSwap = hiBound
  
  Do
    '== Find the right loSwap
    While loSwap < hiSwap And vec(loSwap, SortField) <= pivot(SortField)
      loSwap = loSwap + 1
    Wend
    '== Find the right hiSwap
    While vec(hiSwap, SortField) > pivot(SortField)
      hiSwap = hiSwap - 1
    Wend
    '== Swap values if loSwap is less then hiSwap
    If loSwap < hiSwap Then Call aSwapRows(vec, loSwap, hiSwap)


  Loop While loSwap < hiSwap
  
  For counter = 0 To UBound(vec, 2)
    vec(loBound, counter) = vec(hiSwap, counter)
    vec(hiSwap, counter) = pivot(counter)
  Next
    
  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    If loBound < (hiSwap - 1) Then Call aQuickSort(vec, loBound, hiSwap - 1, SortField)
    '== 2 or more items in second section
    If hiSwap + 1 < hiBound Then Call aQuickSort(vec, hiSwap + 1, hiBound, SortField)
    
End Sub  'QuickSort
Public Sub aSwapRows(ary, row1, row2)
  '== This proc swaps two rows of an array
  Dim X, tempvar
  For X = 0 To UBound(ary, 2)
    tempvar = ary(row1, X)
    ary(row1, X) = ary(row2, X)
    ary(row2, X) = tempvar
  Next
End Sub  'SwapRows


Public Function aRippleSort( _
                                        ByRef ArrayToSort As Variant)

  Dim NumOfEntries As Integer, NumOfTimes As Integer
  Dim i As Integer, J As Integer
  Dim temp As Variant
    
    If Not IsEmpty(ArrayToSort) Then
  NumOfEntries = UBound(ArrayToSort)
  NumOfTimes = NumOfEntries - 1

  For i = LBound(ArrayToSort) To NumOfTimes
    For J = i + 1 To NumOfEntries
      If ArrayToSort(J) < ArrayToSort(i) Then     ' swap array items
        temp = ArrayToSort(J)
        ArrayToSort(J) = ArrayToSort(i)
        ArrayToSort(i) = temp
      End If
    Next J
  Next i
  End If

End Function

Public Function aShellSort( _
                                    a() As Variant)
                                    
    Dim i As Integer, J As Integer
    Dim Low As Integer, hi As Integer
    Dim PushPop As Variant
   Low = LBound(a)
   hi = UBound(a)
   J = (hi - Low + 1) \ 2
   Do While J > 0
     For i = Low To hi - J
      If a(i) > a(i + J) Then
         PushPop = a(i)
         a(i) = a(i + J)
         a(i + J) = PushPop
      End If
     Next i
     For i = hi - J To Low Step -1
      If a(i) > a(i + J) Then
         PushPop = a(i)
         a(i) = a(i + J)
         a(i + J) = PushPop
      End If
     Next i
     J = J \ 2
   Loop

End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' aSortNPack
'    This function sorts an array and removes any repeating
'    values in the list.
Public Function aSortNPack( _
                                        ByRef WorkArray As Variant)
                                        
    aRippleSort WorkArray
    Dim iArrTop As Integer
    iArrTop = UBound(WorkArray)
    Dim i As Integer
    Dim CurrArrCell As Integer
    CurrArrCell = LBound(WorkArray)
    For i = LBound(WorkArray) + 1 To iArrTop
        If WorkArray(i) = WorkArray(CurrArrCell) Then
        Else
            CurrArrCell = CurrArrCell + 1
            WorkArray(CurrArrCell) = WorkArray(i)
        End If
    Next i
    ReDim Preserve WorkArray(CurrArrCell)
    
End Function


Public Function filePrintTo( _
                                    ByVal FileNameToWrite2 As String, _
                                    ByVal String2WriteOut As String, _
                                    Optional ByVal bOverWrite As Boolean) _
                        As Boolean
    
    If bOverWrite Then
        On Error Resume Next
        Kill FileNameToWrite2
        On Error GoTo 0
    End If
    If Right(String2WriteOut, 2) = vbCrLf Then
        String2WriteOut = Left(String2WriteOut, Len(String2WriteOut) - 2)
    End If
    On Error GoTo didntwork
    Dim FileHandle As Integer
    FileHandle = FreeFile
    Open FileNameToWrite2 For Append Access Write Lock Read As #FileHandle
    Print #FileHandle, String2WriteOut
    Close #FileHandle
    filePrintTo = True
    Exit Function
    
didntwork:
    filePrintTo = False
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' fileAttr
'    Used to supercede the varios FileIs functions, this function
'    is useful for quickly testing the attributes of a file (i.e.
'    hidden, system, directory, etc).
Public Function fileAttr( _
                                ByRef FileNameToCheck As String, _
                                ByVal eWhatAttributes As VbFileAttribute) _
                        As Boolean
    
    On Error Resume Next
    fileAttr = CBool(GetAttr(FileNameToCheck) And eWhatAttributes)
    
End Function



Public Function colCopy(ByRef CollectionToCopy As Variant, ByRef CollectionToCopyTo As Collection) As Long
    Dim lCount As Long
    For lCount = 0 To CollectionToCopy.Count - 1
        Dim vNewObject As Variant
        If IsObject(CollectionToCopy(lCount)) Then
            Set vNewObject = CollectionToCopy(lCount)
        Else
            vNewObject = CollectionToCopy(lCount)
        End If
        CollectionToCopyTo.Add vNewObject, vNewObject
    Next lCount
    colCopy = lCount
End Function

Public Function colRemoveAll(ByRef CollectionToClear As Collection)
    Dim iCnt As Integer
    iCnt = CollectionToClear.Count
    Do While iCnt > 0
        CollectionToClear.Remove iCnt
        iCnt = iCnt - 1
    Loop
End Function


Public Function coordsCopy( _
                                        ByRef oFrom, _
                                        ByRef oTo)
    
    On Error Resume Next
    oTo.Left = oFrom.Left
    oTo.Top = oFrom.Top
    oTo.Height = oFrom.Height
    oTo.Width = oFrom.Width
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' wsStateText
'    This function returns the plain text of the .State of
'    a winsock control.
Public Function wsStateText(ByRef WinSockCtl As Object) As String
    Select Case WinSockCtl.State
        Case 0
            wsStateText = "Closed"
        Case 1
            wsStateText = "Open"
        Case 2
            wsStateText = "Listening"
        Case 3
            wsStateText = "Connection Pending"
        Case 4
            wsStateText = "Resolving Host"
        Case 5
            wsStateText = "Host Resolved"
        Case 6
            wsStateText = "Connecting"
        Case 7
            wsStateText = "Connected"
        Case 8
            wsStateText = "Peer Is Closing The Connection"
        Case 9
            wsStateText = "Error"
        Case Else
            wsStateText = "UNKNOWN"
    End Select
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' tbTrim
'    This function will take in a toolbar object and trim it to
'    account for changing size of buttons.
Public Function tbTrim( _
                                ByRef ToolBarToTrim As Object)
    
    ToolBarToTrim.ButtonWidth = 20
    ToolBarToTrim.ButtonHeight = 20
    Dim iMaxButton As Integer
    iMaxButton = ToolBarToTrim.Buttons.Count
    On Error Resume Next
        ToolBarToTrim.Width = ToolBarToTrim.Buttons(iMaxButton).Left + ToolBarToTrim.Buttons(iMaxButton).Width
    On Error GoTo 0
    
    ToolBarToTrim.Height = ToolBarToTrim.Buttons(iMaxButton).Top + ToolBarToTrim.Buttons(iMaxButton).Height - ts.twipsY(1)
        
End Function


Public Function timeHasDate( _
                                            ByVal vTime) _
                        As Boolean
    
    vTime = CDate(vTime)
    Dim sDate As String
    sDate = DatePart("m", vTime) & "/" & DatePart("d", vTime) & "/" & DatePart("yyyy", vTime)
    If CDate(sDate) = 0 Then
        timeHasDate = False
    Else
        timeHasDate = True
    End If
    
End Function


Public Function timeToSeconds( _
                                                ByVal vTime, _
                                                Optional ByVal eReference As enumSecondsTypes = estSinceMidnight) _
                        As Double
    
    vTime = CDate(vTime)
    If Not timeHasDate(vTime) Then
        vTime = vTime + Date
    End If
    
    Dim dBase As Date
    Select Case True
        Case eReference = estSinceBeginningOfWeek
            dBase = Date - Weekday(Now) + 1
        Case eReference = estSinceFirstOfTheMonth
            dBase = Date - Day(Date) + 1
        Case eReference = estSinceFirstOfTheYear
            dBase = Date - DatePart("y", Date) + 1
        Case eReference = estSinceMidnight
            dBase = Date
        Case eReference = estSinceZero
            dBase = 0
    End Select
    timeToSeconds = DateDiff("s", dBase, vTime)
    Exit Function
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' timeFromSeconds
'    This function will generate a time value based on a number of
'    seconds that you pass.
Public Function timeFromSeconds( _
                                                ByVal lSeconds) _
                        As Date
    
    Dim sHour As String
    Dim sMinute As String
    Dim sSeconds As String
    
    sHour = Format(Int(lSeconds / 60 / 60), "00")
    lSeconds = lSeconds - (Val(sHour) * 60 * 60)
    sMinute = Format(Int(lSeconds / 60), "00")
    lSeconds = lSeconds - (Val(sMinute) * 60)
    sSeconds = Format(lSeconds, "00")
    
    timeFromSeconds = CDate(sHour & ":" & sMinute & ":" & sSeconds)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' timeAsElapsed
'    This function will take in a time and return a string expressing it
'    as "## hours, ##minutes, ## seconds"
Public Function timeAsElapsed( _
                                ByVal vTime As Date, _
                                Optional ByVal eFormat As enumTimeElapsedFormats = etfHoursMinutesSeconds) _
                        As String
    
    Dim lHours As Long
    Dim lMinutes As Long
    Dim lSeconds As Long
    lHours = DatePart("h", vTime)
    lMinutes = DatePart("n", vTime)
    lSeconds = DatePart("s", vTime)
    
    Select Case True
        Case eFormat = etfHoursMinutesSeconds
            Dim sRet As String
            If lHours > 0 Then
                If lHours > 1 Then
                    sRet = lHours & " hours "
                Else
                    sRet = lHours & " hours "
                End If
            End If
            If lMinutes > 0 Then
                If lMinutes > 1 Then
                    sRet = sRet & lMinutes & " minutes "
                Else
                    sRet = sRet & lMinutes & " minute "
                End If
            End If
            If lSeconds > 1 Then
                sRet = sRet & lSeconds & " seconds "
            Else
                sRet = sRet & lSeconds & " second "
            End If
        Case eFormat = etfLCD
            If lHours > 0 Then
                sRet = sRet & Format(lHours, "######00") & ":"
            End If
            sRet = sRet & Format(lMinutes, "00") & ":" & Format(lSeconds, "00")
    End Select
    
    timeAsElapsed = sRet
    
End Function


Public Function shellFolderIDFromText( _
                                    ByVal sName As String) _
                As enumShellFolders
    
    Dim Ret As enumShellFolders
    Select Case UCase(sName)
        Case "MY COMPUTER"
            Ret = CSIDL_PERSONAL
        Case "DESKTOP"
            Ret = CSIDL_DESKTOPDIRECTORY
        Case "MY DOCUMENTS"
            Ret = CSIDL_COMMON_DOCUMENTS
        Case "NETWORK NEIGHBORHOOD"
            Ret = CSIDL_NETHOOD
        Case "ENTIRE NETWORK"
            Ret = CSIDL_NETWORK
    End Select
    shellFolderIDFromText = Ret
    
End Function


Public Function cboSyncListIndex( _
                                ByRef cbo As ComboBox)
    
    'ts.ctlTagDataSet cbo, "Synchronizing", "True"
    Dim l As Long
    For l = 0 To cbo.ListCount - 1
        If cbo.List(l) = cbo.Text Then
            If cbo.ListIndex <> l Then
                cbo.ListIndex = l
            End If
            Exit Function
        End If
    Next l
    'ts.ctlTagDataDel cbo, "Synchronizing"
    
End Function


Public Function sHasExtendedChars( _
                            ByVal sToCheck) _
                As Boolean
    
    Dim l As Long
    For l = 1 To Len(sToCheck)
        Dim iAsc As Integer
        iAsc = Asc(Mid(sToCheck, l, 1))
        If iAsc < 32 Or iAsc > 127 Then
            sHasExtendedChars = True
            Exit Function
        End If
    Next l
    
End Function

Public Function cboItemData( _
                            ByRef cbo As ComboBox, _
                            Optional ByVal lDefault As Long = -1) _
                As Long
    
    cboItemData = lDefault
    Dim l As Long
    For l = 0 To cbo.ListCount - 1
        If cbo.Text = cbo.List(l) Then
            cboItemData = cbo.ItemData(l)
            Exit Function
        End If
    Next l
                    
End Function


Public Function dirCreate( _
                        ByVal sPath As String) _
                As Boolean
                
    Dim tSec As SECURITY_ATTRIBUTES
    tSec.lpSecurityDescriptor = 0
    tSec.bInheritHandle = True
    tSec.nLength = Len(tSec)
    
    Dim Ret As Long
                    
    If ts.dirExists(sPath) Then
        dirCreate = True
        Exit Function
    End If
    
    Dim sCurrPath As String
    sPath = ts.sAppend(sPath, "\")
    sPath = Replace(sPath, "\\", "??")
    Dim i As Integer
    Dim iCnt As Integer
    iCnt = sWordCnt(sPath, "\")
    For i = 1 To iCnt
        sCurrPath = sCurrPath & Replace(sWord(sPath, i, "\"), "??", "\\") & "\"
        Ret = CreateDirectory(sCurrPath, tSec)
        If i = iCnt Then
            dirCreate = (Ret <> 0)
        End If
    Next i
    
End Function



Public Function windowDesktopHWnd() _
                As Long
    windowDesktopHWnd = GetDesktopWindow()
    
End Function


Public Function nDecimalSet( _
                            dblValue As Double, _
                            intDecimals As Integer) _
                As Double
                
    Dim sWrk As String
    sWrk = Format(dblValue, "0." & String$(intDecimals + 1, "0"))
    If intDecimals = 0 Then
        nDecimalSet = CDbl(Left$(sWrk, Len(sWrk) - 2))
    Else
        nDecimalSet = CDbl(Left$(sWrk, Len(sWrk) - 1))
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sMoney
'    Simple little function to take a value and format it as money.
Public Function sMoney( _
                                    ByVal vInput) _
                        As String
    
    If Not IsNull(vInput) Then
        sMoney = Format(vInput, "$#,0.00")
    Else
        sMoney = "$0.00"
    End If
        
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' tabMouseUp
'    This little function is primarily to be used to cause the right click to act like a left click when the tabs
'    are being clicked.  Assumptions made are:
'       There are not multiple rows of tabs and
'       The TabsPerRow property is the same as the Tabs property.
Public Function tabMouseUp( _
                                        ByRef tabCtl, _
                                        ByVal Button As MouseButtonConstants, _
                                        ByVal Shift As ShiftConstants, _
                                        ByVal X As Long, _
                                        ByVal Y As Long) _
                        As Boolean
    
    If Button = vbRightButton Then
        ' Calculate if being a tab clicked
        If Y < tabCtl.TabHeight Then
            
            ' Calculate the tab being clicked
            Dim TabClicked As Long
            TabClicked = Int(X / (tabCtl.Width / tabCtl.Tabs))
            If TabClicked <> tabCtl.Tab Then
                tabCtl.Tab = TabClicked
                DoEvents
            End If
            
        End If
        
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' tvRemoveChildren
'    This function is used to remove all of the children from a particular
'    node in a treeview.
Public Function tvRemoveChildren( _
                                ByRef tv, _
                                ByRef Node) _
                As Long
                    
    Do While Node.Children > 0
        tv.Nodes.Remove Node.Child.Index
    Loop
                
End Function


Public Function nodeExists( _
                            ByVal sKey As String, _
                            ByRef tv) _
                As Boolean
    
    Dim sTest As String
    On Error Resume Next
    sTest = tv.Nodes(sKey).key
    nodeExists = (Err.Number = 0)
    On Error GoTo 0
    
End Function


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' NoEmpty
'    Use this function where you wish to assign a default for empty values (typically
'    coming from the database).  Detects nulls also.
Public Function NoEmpty( _
                                    ByRef vVariable, _
                                    ByVal vDefault As Variant) _
                        As Variant
    
    If IsEmpty(NoNull(vVariable)) Then
        NoEmpty = vDefault
    Else
        NoEmpty = vVariable
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' NoNull
'    NoNull is a wrapper to put around field values coming
'    out of a database when you cannot have them be null.
Public Function NoNull(ByRef vValue As Variant) As Variant
    If Not IsNull(vValue) Then
        NoNull = vValue
        Exit Function
    End If
    NoNull = Empty
End Function

Public Function timeFileToDate( _
                                ft As FILETIME) _
                As Date
                    
    Dim tSysTime As SYSTEMTIME
    FileTimeToSystemTime ft, tSysTime
    timeFileToDate = ts.timeSysToDate(tSysTime)
                    
End Function
                            
Public Function timeDateToFile( _
                                ByVal dDate As Date) _
                As FILETIME
                
    Dim tRet As FILETIME
    Dim tSys As SYSTEMTIME
    tSys = timeDateToSys(dDate)
    SystemTimeToFileTime tSys, tRet
    timeDateToFile = tRet
    
End Function

Public Function timeSysToDate( _
                            st As SYSTEMTIME) _
                As Date
    
    timeSysToDate = CDate(Format(st.wMonth, "00") & "/" & Format(st.wDay, "00") & "/" & Format(st.wYear, "0000") & " " & Format(st.wHour, "00") & ":" & Format(st.wMinute, "00") & ":" & Format(st.wSecond, "00"))
    
End Function

Public Function timeDateToSys( _
                            ByVal dDateTime As Date) _
                As SYSTEMTIME
                    
    Dim tRet As SYSTEMTIME
    tRet.wDay = Day(dDateTime)
    tRet.wMonth = Month(dDateTime)
    tRet.wYear = Year(dDateTime)
    tRet.wHour = Hour(dDateTime)
    tRet.wMinute = Minute(dDateTime)
    tRet.wSecond = Second(dDateTime)
    timeDateToSys = tRet
    
End Function



Public Function fileCount( _
                            ByVal sSpec As String) _
                As Long
                    
    Dim tInfo As WIN32_FIND_DATA
    Dim lCnt As Long
    Dim lFind As Long, lMatch As Long
    lFind = FindFirstFile(sSpec, tInfo)
    lMatch = 99
    Do While lFind > 0 And lMatch > 0
        lCnt = lCnt + 1
        lMatch = FindNextFile(lFind, tInfo)
    Loop
    fileCount = lCnt
    
End Function


Public Function fileAttributes( _
                            ByVal sFileName As String) _
                As enumFileAttributes
    
    fileAttributes = GetFileAttributes(sFileName)
    
End Function

Public Function fileOpenStructure( _
                                ByVal sFileName As String) _
                As OFSTRUCT
    
    Dim tOF As OFSTRUCT
    Dim lHandle As Long
    lHandle = OpenFile(sFileName, tOF, 0)
    CloseHandle lHandle
    fileOpenStructure = tOF
    
End Function


Public Function fileInformation( _
                            ByVal sFileName As String) _
                As BY_HANDLE_FILE_INFORMATION
                    
    Dim tInfo As BY_HANDLE_FILE_INFORMATION
    Dim tOF As OFSTRUCT
    Dim lHandle As Long
    lHandle = OpenFile(sFileName, tOF, 0)
    If lHandle > 0 Then
        GetFileInformationByHandle lHandle, tInfo
    End If
    fileInformation = tInfo
    CloseHandle lHandle
    
End Function


Public Function fileExpandedName( _
                                ByVal sFileName As String) _
                As String
                    
    Dim sBuffer As String
    sBuffer = Space(1024)
    GetExpandedName sFileName, sBuffer
    fileExpandedName = ts.sNT(sBuffer)
    
End Function

Public Function fileShortName( _
                                ByVal sFileName As String) _
                As String
                    
    Dim sBuffer As String
    sBuffer = Space(1024)
    GetShortPathName sFileName, sBuffer, Len(sBuffer)
    fileShortName = ts.sNT(sBuffer)
    
End Function

Public Function formCenterInForm( _
                                                ByRef FormToCenter As Form, _
                                                ByRef FormToCenterIn As Form)
        
    Dim lLeft As Long, lTop As Long
    lTop = (FormToCenterIn.Top + (FormToCenterIn.Height / 2)) - (FormToCenter.Height / 2)
    lLeft = (FormToCenterIn.Left + (FormToCenterIn.Width / 2)) - (FormToCenter.Width / 2)
    
    ts.formMoveAbs FormToCenter, lLeft, lTop
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' dateBOM
'    This function will return a date that represents the absolute
'    beginning of the month.
Public Function dateBOM( _
                                    ByVal dIn As Date) _
                        As Date
    
    Dim lYear As Long
    Dim lMonth As Long
    lYear = Year(dIn)
    lMonth = Month(dIn)
    dateBOM = CDate(Format(lMonth, "00") & "/01/" & Format(lYear, "0000"))
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' dateEOM
'    This function will return the date that represents the absolute
'    end of the month.
Public Function dateEOM( _
                                    ByVal dIn As Date) _
                        As Date
    
    Dim lYear As Long
    Dim lMonth As Long
    Dim lDay As Long
    
    lYear = Year(dIn)
    lMonth = Month(dIn)
    lDay = Day(dIn)
    
    Do While Month(dIn) = lMonth
        lDay = Day(dIn)
        dIn = dIn + 1
    Loop
    
    dateEOM = CDate(Format(lMonth, "00") & "/" & Format(lDay, "00") & "/" & Format(lYear, "0000"))
    
End Function


Public Function cboShowDropDown( _
                                                ByRef oCombo As ComboBox, _
                                                Optional ByVal bShow As Boolean = True)
    
    SendMessage oCombo.hWnd, CB_SHOWDROPDOWN, ByVal bShow, 0
    
End Function

Public Function cboDroppedDown( _
                                                ByRef oCombo As ComboBox) _
                        As Boolean
    
    cboDroppedDown = SendMessage(oCombo.hWnd, CB_GETDROPPEDSTATE, 0, 0)
    
End Function


Public Function cboMaxWidth( _
                                            ByRef oCombo As ComboBox, _
                                            ByVal lWidth As Long)
    
    SendMessage oCombo.hWnd, CB_LIMITTEXT, lWidth, 0
    
End Function

Public Function cboSelectString( _
                                            ByRef oCombo As ComboBox, _
                                            Optional ByVal sSearch As String = "~~~NONE~~~") _
                        As Long
    
    If sSearch = "~~~NONE~~~" Then
        sSearch = oCombo.Text
    End If
    sSearch = sSearch & Chr(0)
    cboSelectString = SendMessage(oCombo.hWnd, CB_SELECTSTRING, ByVal -1, ByVal sSearch)
    
End Function

Public Function cboFindString( _
                                        ByRef oCombo As ComboBox, _
                                        Optional ByVal sSearch As String = "~~~NONE~~~") _
                        As Long
    If sSearch = "~~~NONE~~~" Then
        sSearch = oCombo.Text
    End If
    sSearch = sSearch & Chr(0)
    cboFindString = SendMessage(oCombo.hWnd, CB_FINDSTRING, ByVal -1, ByVal sSearch)
    
End Function

Public Function cboFindStringExact( _
                                        ByRef oCombo, _
                                        Optional ByVal sSearch As String = "~~~NONE~~~") _
                        As Long
    
    If sSearch = "~~~NONE~~~" Then
        sSearch = oCombo.Text
    End If
    sSearch = sSearch & Chr(0)
    cboFindStringExact = SendMessage(oCombo.hWnd, CB_FINDSTRINGEXACT, ByVal -1, ByVal sSearch)
    
End Function

Public Function cboChange( _
                                    ByRef oCombo As ComboBox) _
                        As Boolean
    
    Dim lCursorPos As Long
    lCursorPos = oCombo.SelStart
    If ts.cboSelectString(oCombo) > -1 Then
        oCombo.SelStart = lCursorPos
        oCombo.SelLength = 24000
        cboChange = True
    End If
    
End Function


Public Function cboKeyDown( _
                                    ByRef oCombo As ComboBox, _
                                    ByVal KeyCode As Integer, _
                                    ByVal Shift As Integer)
    
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) And Shift = 0 Then
        ts.cboShowDropDown oCombo
    End If
    
End Function


Public Function cboKeyPress( _
                                    ByRef oCombo As ComboBox, _
                                    ByRef KeyAscii As Integer)
    
    If KeyAscii <> 0 And KeyAscii <> 13 And KeyAscii <> 10 And KeyAscii <> 8 Then
        ts.cboShowDropDown oCombo
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' nPower
'    This function will return the power of a number based upon a root number.
Public Function nPower( _
                        ByVal BigNumber, _
                        ByVal RootNumber) _
                As Long
    
    Dim lCnt As Long
    Do While BigNumber / RootNumber >= 1
        BigNumber = BigNumber / RootNumber
        lCnt = lCnt + 1
    Loop
    nPower = lCnt
    
End Function



Public Function aAddTo( _
                                    ByRef vArr As Variant, _
                                    ByVal vValue As Variant) _
                        As Variant
    
    Dim lSize As Long
    If IsEmpty(vArr) Then
        ReDim vArr(0)
        lSize = 0
    Else
        lSize = UBound(vArr)
        lSize = lSize + 1
    ReDim Preserve vArr(lSize)
    End If
    
    vArr(lSize) = vValue
    aAddTo = vArr
    
End Function


Public Function cmdSetStyle( _
                                ByRef cmd, _
                                ByVal eStyle As enumButtonStyles)
    
    SendMessage cmd.hWnd, BM_SETSTYLE, ByVal eStyle, ByVal -1
    
End Function

Public Function cmdSetState( _
                                ByRef cmd As CommandButton, _
                                ByVal bSelected As Boolean)
    
    SendMessage cmd.hWnd, BM_SETSTATE, ByVal bSelected, ByVal 0
    
End Function


Public Function sysColor( _
                                    ByVal eSysColor As enumSysColors) _
                        As Long
    
    sysColor = GetSysColor(eSysColor)
    
End Function


Public Function txtInsert( _
                                    ByRef txtCtl As TextBox, _
                                    ByVal sText As String)
    
    Dim lStart As Long
    Dim lLen As Long
    lStart = txtCtl.SelStart
    lLen = txtCtl.SelLength
'    If lStart = 0 Then
'        lStart = 1
'    End If
    
    Dim sLeft As String
    Dim sRight As String
    sLeft = Left(txtCtl.Text, lStart)
    sRight = Right(txtCtl.Text, Len(txtCtl.Text) - (lStart + lLen))
    
    txtCtl.Text = sLeft & sText & sRight
    txtCtl.SelStart = lStart + Len(sText)
'    txtCtl.Refresh
'    txtCtl.SelStart = lStart
'    txtCtl.SelLength = lLen
    
End Function



Public Function abQueryPos( _
                                        ByVal eEdge As enumAppBarEdges) _
                        As APPBARDATA
    
    Dim tBarData As APPBARDATA
    tBarData.uEdge = eEdge
    tBarData.cbSize = Len(tBarData)
    SHAppBarMessage ABM_GETAUTOHIDEBAR, tBarData
    abQueryPos = tBarData
    
End Function

Public Function abState(ByVal hWnd As Long) As enumAppBarStates
    Dim tBarData As APPBARDATA
    tBarData.hWnd = hWnd
    tBarData.cbSize = Len(tBarData)
    abState = SHAppBarMessage(ABM_GETSTATE, tBarData)
    
End Function

Public Function abData(ByVal hWnd As Long) As APPBARDATA
    
    Dim tData As APPBARDATA
    tData.hWnd = hWnd
    tData.cbSize = Len(tData)
    SHAppBarMessage ABM_GETTASKBARPOS, tData
    abData = tData
    
End Function

Public Function abPosition(ByRef abData As APPBARDATA)
    
    SHAppBarMessage ABM_GETTASKBARPOS, abData
    
End Function

#If bWindowCls Then

Public Function abEnumWindows() As Collection
    
    Dim cRet As New Collection
    
    Dim cWindow As New clsWindow
    cWindow.hWnd = GetDesktopWindow
    cWindow.RefreshChildren False
    Dim tBarData As APPBARDATA
    tBarData.cbSize = Len(tBarData)
    
    Dim l As Long
    For l = 1 To cWindow.Children.Count
        If (UCase(cWindow.Children(l).sClassName) = "BASEBAR" Or UCase(cWindow.Children(l).sClassName) = "SHELL_TRAYWND") And (cWindow.Children(l).eWindowStyle And WS_VISIBLE) Then
            cRet.Add CLng(cWindow.Children(l).hWnd)
        End If
    Next l
        
    Set abEnumWindows = cRet
    Set cWindow = Nothing
            
End Function



Public Function debugListAppBars()
    Dim c As Collection
    Set c = ts.abEnumWindows
    Dim cWindow As New clsWindow
    Dim l As Long
    For l = 1 To c.Count
        Dim sPrint As String
        sPrint = sPadR(l, 5) & sPadR(Hex(c(l)), 10)
        Dim tData As APPBARDATA
        tData = ts.abData(c(l))
        cWindow.hWnd = c(l)
        sPrint = sPrint & "L:" & cWindow.Left & "  R: " & cWindow.Right & " T:" & cWindow.Top & " B:" & cWindow.Bottom & "  EDGE:" & rectGetEdge(cWindow)
        Debug.Print sPrint
        
    Next l
    Set cWindow = Nothing
    
End Function

#End If

Public Function rectGetEdge(ByRef rc) As enumAppBarEdges
    rectGetEdge = ABE_FLOATING
    Select Case True
        Case rc.Top = rc.Left And rc.Bottom > rc.Right
            rectGetEdge = ABE_LEFT
        Case rc.Top = rc.Left And rc.Bottom < rc.Right
            rectGetEdge = ABE_TOP
        Case rc.Top > rc.Left
            rectGetEdge = ABE_BOTTOM
        Case Else
            rectGetEdge = ABE_RIGHT
    End Select
End Function
'UINT CMainFrame::GetEdge(CRect rc)
'      {
'      UINT uEdge = -1;
'
'      if (rc.top == rc.left && rc.bottom > rc.right)
'      {
'          uEdge = ABE_LEFT;
'      }
'      else if (rc.top == rc.left && rc.bottom < rc.right)
'      {
'          uEdge = ABE_TOP;
'      }
'      else if (rc.top > rc.left )
'      {
'          uEdge = ABE_BOTTOM;
'      }
'      Else
'      {
'          uEdge = ABE_RIGHT;
'      }
'
'         return uEdge;
'      }
'


Public Function fileLongPath(ByVal sFile As String) As String
    Dim sRet As String
    sRet = Space(2048)
    Dim lLen As Long
    lLen = GetLongPathName(sFile, sRet, Len(sRet))
    fileLongPath = Left(sRet, lLen)
End Function

#If bWindowCls Then
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' shellGetDesktopRect
'    This function will return the current usable desktop space.
'    This function takes into account dockbars on all edges including
'    their states (auto-hide, always-on-top, etc.)
Public Function shellGetDesktopRectPixels() As Rect
    
    Dim rRect As Rect
    Dim cWindow As New clsWindow
    cWindow.hWnd = GetDesktopWindow
    cWindow.RefreshChildren False
    cWindow.Children(cWindow.Children.Count).RefreshChildren True
    rRect.Left = cWindow.Children(cWindow.Children.Count).Children(1).Left
    rRect.Top = cWindow.Children(cWindow.Children.Count).Children(1).Top
    rRect.Right = cWindow.Children(cWindow.Children.Count).Children(1).Right
    rRect.Bottom = cWindow.Children(cWindow.Children.Count).Children(1).Bottom
    shellGetDesktopRectPixels = rRect
    Set cWindow = Nothing
    
End Function

Public Function shellGetDesktopRectTwips() As Rect
    Dim rRect As Rect
    rRect = shellGetDesktopRectPixels
    With rRect
        .Top = twipsY(.Top)
        .Bottom = twipsY(.Bottom)
        .Left = twipsX(.Left)
        .Right = twipsX(.Right)
    End With
    shellGetDesktopRectTwips = rRect
End Function

#End If


Public Function byteArray(ByVal Value) As Byte()
    Dim bArr() As Byte
    Dim lLen As Long
    lLen = LenB(Value)
    ReDim Preserve bArr(lLen)
    CopyMemory VarPtr(bArr(0)), VarPtr(Value), lLen
    byteArray = bArr
    
End Function

Public Function byteArrayToVar(ByRef bArr() As Byte) As Variant
    Dim vRet As Variant
    Dim lLen As Long
    lLen = UBound(bArr)
    CopyMemory vRet, bArr(0), lLen
    byteArrayToVar = vRet
    
End Function

Public Function TestByteArray()
''    Dim aBytes() As Byte
'''    aBytes = byteArray("This is a test")
'''    Debug.Print byteArrayToVar(aBytes)
''
''    Dim Jeff2 As New Collection
''    Jeff2.Add "TEST"
''    Jeff2.Add "John"
''    Jeff2.Add "Clark"
''
''    Dim test2 As String
''    test2 = "Blah blah blah"
''
''    Dim jeff As Variant
''    ReDim jeff(5)
''    jeff(1) = "Jeff"
''    jeff(2) = "Clark"
''    jeff(3) = "Murl"
''    jeff(4) = "Karen"
''    jeff(5) = "John"
''
''    aBytes = baFromVA(jeff)
'''    aBytes = baFromCol(Jeff2)
''
''    Dim test As Variant
''    test = baToVA(aBytes)
''
''    Debug.Print test(4)
''    Debug.Print test(2)
''
''
End Function


Public Function baFromVar(ByVal Value As Variant) As Byte()
    
    Dim bRet() As Byte
    Dim lLen As Long
    lLen = LenB(Value)
    ReDim bRet(lLen)
    bRet = Value
    baFromVar = bRet
    
End Function

Public Function baToVar(ByRef bArr() As Byte) As Variant
    baToVar = bArr
End Function




#If bSendArrayDll Then
Public Function baFromVA(ByRef vArr As Variant) As Byte()
    Dim bRet() As Byte
    Dim lLen As Long
    ReDim bRet(2048)
    GetBufferFromVariantArray vArr, lLen, bRet
    ReDim Preserve bRet(lLen)
    baFromVA = bRet
End Function
Public Function baToVA(ByRef bArr As Variant) As Variant
    Dim vRet As Variant
    GetVariantArrayFromBuffer UBound(bArr), bArr, vRet
    baToVA = vRet
End Function
#End If


Public Function byteFromPtr(ByVal lPtr As Long) As Byte
    Dim bRet As Byte
    On Error Resume Next
    CopyMemory lPtr, ByVal bRet, 1
    On Error GoTo 0
    byteFromPtr = bRet
End Function


Public Function varSize(ByRef var) As Long
    Dim lSize As Long
    Dim lPtr As Long
    lPtr = VarPtr(var)
    Dim lHeap As Long
    lHeap = GetProcessHeap
    varSize = HeapSize(lHeap, 0, lPtr)
End Function


Public Function nRemainder(ByVal dValue As Double) As Double
    nRemainder = dValue - Int(dValue)
End Function


Public Function fileCompactPath( _
                                ByVal sFileName As String, _
                                Optional ByVal iChars As Integer = 30) _
                As String
    
    Dim sBuf As String
    sBuf = Space(Len(sFileName))
    Dim lResult As Long
    lResult = PathCompactPathEx(sBuf, sFileName, iChars + 1, 0)
    If lResult Then
        fileCompactPath = ts.sNT(sBuf)
    Else
        fileCompactPath = sFileName
    End If
        
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' chk3StateMouseUp
'    This function is used on the MouseUp events of CheckBoxes that have had there
'    style set (with the ts.cmdSetStyle function) to a 3 state check box.
Public Function chk3StateMouseUp( _
                                ByRef ctlCheck As CheckBox, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    If Button = vbLeftButton Then
        With ctlCheck
            Select Case True
                Case .Value = 1
                    .Value = 2
                Case .Value = 2
                    .Value = 0
                Case Else
                    .Value = 1
            End Select
        End With
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sCreateGUID
'    This function will create a globally unique identifier.
Public Function sCreateGUID() As String
    
    Dim lResult As Long, uGUID As GUID
    Dim sGUID As String
    
    On Error Resume Next
    
    lResult = CoCreateGuid(uGUID)
    
    If lResult <> GUID_OK Then
        sCreateGUID = Empty
        Exit Function
    End If
    
    sGUID = String$(GUID_LENGTH, 0)
    
    lResult = StringFromGUID2(uGUID, StrPtr(sGUID), 1 + GUID_LENGTH)
    
    If lResult <> 1 + GUID_LENGTH Then
        sCreateGUID = Empty
    Else
        sCreateGUID = sGUID
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sCreateUKey
'    This function will scramble a guid and return it as a 32 character
'    unique key.
Public Function sCreateUKey() As String
    Dim sGUID As String
    sGUID = ts.sCreateGUID
    sGUID = Replace(sGUID, "{", "")
    sGUID = Replace(sGUID, "-", "")
    sGUID = Replace(sGUID, "}", "")
    Dim l As Long
    Dim sRet As String
    For l = Len(sGUID) To 1 Step -2
        sRet = sRet & Mid(sGUID, l, 1)
    Next l
    For l = Len(sGUID) - 1 To 1 Step -2
        sRet = sRet & Mid(sGUID, l, 1)
    Next l
    sCreateUKey = sRet
End Function


Public Function volumeInformation(ByVal sDrive As String) As typeVolumeInformation
    
    Dim Ret As typeVolumeInformation
    Ret.sRootPathName = sDrive
    Ret.sFileSystemName = Space(1024)
    Ret.sVolumeName = Space(1024)
    GetVolumeInformation Ret.sRootPathName, Ret.sVolumeName, Len(Ret.sVolumeName), Ret.lVolumeSerialNo, Ret.lMaximumComponentLength, Ret.lFileSystemFlags, Ret.sFileSystemName, Len(Ret.sFileSystemName)
    Ret.sFileSystemName = ts.sNT(Ret.sFileSystemName)
    Ret.sVolumeName = ts.sNT(Ret.sVolumeName)
    volumeInformation = Ret
    
End Function

Public Function fileRoot(ByVal sFileName As String) As String
    
    Dim lngResult As Long
    lngResult = PathStripToRoot(sFileName)
    If lngResult <> 0 Then
        If InStr(sFileName, vbNullChar) > 0 Then
            fileRoot = Left$(sFileName, InStr(sFileName, vbNullChar) - 1)
        Else
            fileRoot = sFileName
        End If
    End If
    
End Function



#If bRegistryCls Then
Public Function driveMapping(ByVal sDrive As String) As String
    Dim cReg As New cRegistry
    cReg.ClassKey = HKEY_USERS
    cReg.SectionKey = ".DEFAULT\Network\Persistent\" & Left(sDrive, 1)
    cReg.ValueKey = "RemotePath"
    driveMapping = cReg.Value
    Set cReg = Nothing
    
End Function

#End If


Public Function fileSetTime(ByVal sFileName As String, ByVal dLastWrite As Date)
    
    Dim timeCreated As FILETIME
    Dim timeAccess As FILETIME
    Dim timeWrite As FILETIME
    Dim tInfo As BY_HANDLE_FILE_INFORMATION
    Dim tOF As OFSTRUCT
    Dim lHandle As Long
    lHandle = OpenFile(sFileName, tOF, 0)
    
    If lHandle > 0 Then
        GetFileTime lHandle, timeCreated, timeAccess, timeWrite
        timeWrite = ts.timeDateToFile(dLastWrite)
        SetFileTime lHandle, timeCreated, timeAccess, timeWrite
    End If
    CloseHandle lHandle
        
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sqlFix
'    This functions provides a simple "quick fix"
'    for dealing with SQL server and the use of
'    quotes and ticks in a SQL statement.  If the
'    string has any quotes (invalid syntax in SQL
'    Server back end, yet fine in jet), THE ASSUMPTION
'    is made that you marked your strings with quotes
'    and no quotes were allowed in your actual data.
'    To deal with quotes in data, you'll need to use
'    the Replace() function when making your SQL string.
Public Function sqlFix(ByRef sSQL As String)
    If InStr(sSQL, """") > 0 Then
        sSQL = Replace(sSQL, "'", "''")
        sSQL = Replace(sSQL, """", "'")
    End If
    
End Function





Public Function mouseLeftClick(ByVal XPos As Integer, ByVal YPos As Integer)
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN, XPos, YPos, 0, 0
    'DoEvents
End Function

Public Function mouseLeftUp(ByVal XPos As Integer, ByVal YPos As Integer)
    mouse_event MOUSEEVENTF_LEFTUP, XPos, YPos, 0, 0
    'DoEvents
End Function

Public Function mouseMiddleClick(ByVal XPos As Integer, ByVal YPos As Integer)
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MIDDLEDOWN, XPos, YPos, 0, 0
    'DoEvents
End Function

Public Function mouseRightClick(ByVal XPos As Integer, ByVal YPos As Integer)
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_RIGHTDOWN, XPos, YPos, 0, 0
    'DoEvents
End Function

Public Function mouseRightUp(ByVal XPos As Integer, ByVal YPos As Integer)
    mouse_event MOUSEEVENTF_RIGHTUP, XPos, YPos, 0, 0
    'DoEvents
End Function

Public Function ctlSendMessage(ByVal hWnd As Long, ByVal eMessage As enumWindowsMessages, Optional ByVal wParam As Long = 0, Optional ByVal lParam As Long = 0)
    SendMessage hWnd, eMessage, wParam, ByVal lParam
End Function


Public Function nDWordFromWords( _
                                ByVal LoWord As Integer, _
                                ByVal HiWord As Integer) _
                        As Long
    
    Dim oDWord As OneDWord
    Dim oWords As TwoWords
    
    oWords.LoWord = LoWord
    oWords.HiWord = HiWord
    LSet oDWord = oWords
    nDWordFromWords = oDWord.dWord
    
End Function



Public Sub fileSetLength(ByVal FileName As String, ByVal NewLength As Long)
   'Will cut the length of a file to the length specified.
   
   Dim l As Long
   Dim hFile As Long
   Const ZERO = 0
   
   'if file is smaller than or equal to requsted length, exit.
   If FileLen(FileName) <= NewLength Then Exit Sub
   'open the file
   hFile = CreateFile(FileName, GENERIC_WRITE, ZERO, ByVal ZERO, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
   'if file not open exit
   If hFile = -1 Then Exit Sub
   'seek to position
   l = SetFilePointer(hFile, NewLength, ZERO, ZERO)
   'and mark here as end of file
   SetEndOfFile hFile
   'close the file
   l = CloseHandle(hFile)
   
End Sub

Public Function fileLength(ByVal sFileName As String) As Long

    Dim FileHandle As Integer
    
    FileHandle = FreeFile
    On Error Resume Next
    Open sFileName For Input As #FileHandle
    fileLength = LOF(FileHandle)
    Close #FileHandle
    On Error GoTo 0
    
End Function


Public Function rectMake(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long) As Rect
    Dim tRet As Rect
    tRet.Bottom = lBottom
    tRet.Top = lTop
    tRet.Left = lLeft
    tRet.Right = lRight
    rectMake = tRet
End Function


#If bTreeView Then

Public Function tvSetFirstVisibleNode(ByVal tv As TreeView, ByVal Node As Node)
    Dim hItem As Long
    Dim selNode As Node
    
    ts.windowUpdate tv.hWnd, elwLOCK
    
    ' remember the node currently selected
    Set selNode = tv.SelectedItem
    ' make the Node the select Node in the control
    Set tv.SelectedItem = Node
    ' now we can get its handle
    hItem = SendMessage(tv.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    ' restore node that was selected
    Set tv.SelectedItem = selNode
    ' make it the fist visible Node
    SendMessage tv.hWnd, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, ByVal hItem
    
    ts.windowUpdate tv.hWnd, elwUNLOCK
    
End Function


Function tvFirstVisibleNode(ByVal tv As TreeView) As Node
    Dim hItem As Long
    Dim selNode As Node
    
    ' remember the node currently selected
    Set selNode = tv.SelectedItem
    ' get the handle of the first visible Node
    hItem = SendMessage(tv.hWnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal 0&)
    ' make it the selected Node
    SendMessage tv.hWnd, TVM_SELECTITEM, TVGN_CARET, ByVal hItem
    ' return the result as a Node object
    Set tvFirstVisibleNode = tv.SelectedItem
    ' restore node that was selected
    Set tv.SelectedItem = selNode
    
End Function

Function tvNodeLevel(ByVal Node As Node) As Integer
    Do Until (Node.Parent Is Nothing)
        tvNodeLevel = tvNodeLevel + 1
        Set Node = Node.Parent
    Loop
End Function

#End If


Public Sub bmpTile(Target As Object, BITMAP As StdPicture)
    Dim X As Single
    Dim Y As Single
    Dim bmpWidth As Single
    Dim bmpHeight As Single

    ' get Bitmap's size, in the coordinate system
    ' of the target Form or BitmapBox
    bmpWidth = Target.ScaleX(BITMAP.Width, vbHimetric, Target.ScaleMode)
    bmpHeight = Target.ScaleY(BITMAP.Height, vbHimetric, Target.ScaleMode)
    
    ' tile the Bitmap
    For X = 0 To Target.Width Step bmpWidth
        For Y = 0 To Target.Height Step bmpHeight
            Target.PaintPicture BITMAP, X, Y
        Next
    Next
End Sub



' Enable extended matching to any type combobox control
'
' Extended matching means that as soon as you type in the edit area
' of the ComboBox control, the routine searches for a partial match
' in the list area and highlights the characters left to be typed.
'
' To enable this capability you have only to call this routine
' from within the KeyPress routine of the ComboBox, as follows:
'
' Private Sub Combo1_KeyPress(KeyAscii As Integer)
'    ComboBoxExtendedMatching Combo1, KeyAscii
' End Sub

Sub ComboBoxExtendedMatching(cbo As ComboBox, KeyAscii As Integer, _
    Optional CompareMode As VbCompareMethod = vbTextCompare)
    Dim Index As Long
    Dim Text As String
    
    ' if user pressed a control key, do nothing
    If KeyAscii <= 32 Then Exit Sub
    
    ' produce new text, cancel automatic key processing
    Text = Left$(cbo.Text, cbo.SelStart) & Chr$(KeyAscii) & Mid$(cbo.Text, _
        cbo.SelStart + 1 + cbo.SelLength)
    KeyAscii = 0
    
    ' search the current item in the list
    For Index = 0 To cbo.ListCount - 1
        If InStr(1, cbo.List(Index), Text, CompareMode) = 1 Then
            ' we've found a match
            cbo.ListIndex = Index
            Exit For
        End If
    Next
    
    ' if no matching item
    If Index = cbo.ListCount Then
        cbo.Text = Text
    End If
    
    ' highlight trailing chars in the edit area
    cbo.SelStart = Len(Text)
    cbo.SelLength = 9999
    
End Sub


Public Function txtForceNumeric(TextBox As TextBox, Optional Force As Boolean = True)
    Dim Style As Long
    Const GWL_STYLE = (-16)
    Const ES_NUMBER = &H2000
    
    ' get current style
    Style = GetWindowLong(TextBox.hWnd, GWL_STYLE)
    If Force Then
        Style = Style Or ES_NUMBER
    Else
        Style = Style And Not ES_NUMBER
    End If
    ' enforce new style
    SetWindowLong TextBox.hWnd, GWL_STYLE, Style
End Function


Public Sub txtForceCase(TextBox As TextBox, Optional ConvertCase As Integer)
    
    Dim Style As Long
    Const GWL_STYLE = (-16)
    Const ES_UPPERCASE = &H8&
    Const ES_LOWERCASE = &H10&
    
    ' get current style
    Style = GetWindowLong(TextBox.hWnd, GWL_STYLE)
    
    Select Case ConvertCase
        Case 0
            ' restore default style
            Style = Style And Not (ES_UPPERCASE Or ES_LOWERCASE)
        Case 1
            ' convert to uppercase
            Style = Style Or ES_UPPERCASE
        Case 2
            ' convert to lowecase
            Style = Style Or ES_LOWERCASE
    End Select
    ' enforce new style
    SetWindowLong TextBox.hWnd, GWL_STYLE, Style
End Sub


Sub cboSetDropDownHeight(CB As ComboBox, ByVal newHeight As Long)
    
    Dim lpRect As Rect
    Dim wi As Long
    
    ' get combobox rectangle, relative to screen
    GetWindowRect CB.hWnd, lpRect
    wi = lpRect.Right - lpRect.Left
    
    ' convert to form's client coordinates
    ScreenToClientAny CB.Parent.hWnd, lpRect
    
    ' enforce the new height
    MoveWindow CB.hWnd, lpRect.Left, lpRect.Top, wi, newHeight, True

End Sub






Public Function ctlAddOfficeBorder(ByVal hWnd As Long)
    
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong hWnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function



#If bTreeView Then


Public Function tvRemoveNode(ByRef oTV As TreeView, ByRef oNode As Node)
    tvRemoveChildren oTV, oNode
    oTV.Nodes.Remove oNode.Index
    
End Function

#End If


Public Function ARRTEST()
    Dim vArr As Variant
    vArr = aRedim(vArr, 4, 4)
    
    vArr = aRedim(vArr, 2, 2)
    
    vArr = aRedim(vArr, 5, 3)
    
    vArr(2, 2) = "JEFF"
    
    vArr = aRedim(vArr, 4, 8)
    
    Debug.Print vArr(2, 2)
    
End Function



Public Function aRedim(ByRef vArr As Variant, ByVal X As Long, Optional ByVal Y As Long = -1) As Variant
    
    Dim vRet As Variant
    If Y > -1 Then
        ReDim vRet(X, Y)
    Else
        ReDim vRet(X)
    End If
    
    Dim lTest As Long
    On Error Resume Next
    lTest = LBound(vArr)
    If Err.Number = 0 Then
        On Error GoTo 0
        
        Dim l As Long
        For l = LBound(vRet) To UBound(vRet)
            If Y > -1 Then
                Dim m As Long
                For m = LBound(vRet, 2) To UBound(vRet, 2)
                    On Error Resume Next
                    vRet(l, m) = vArr(l, m)
                    On Error GoTo 0
                Next m
            Else
                On Error Resume Next
                vRet(l) = vArr(l)
                On Error GoTo 0
            End If
        Next l
    End If
    On Error GoTo 0
    
    aRedim = vRet
    
End Function

Public Function shellCreateTempFile( _
                                Optional ByVal ThreeCharPrefix As String = "TMP") _
                As String
    
    Dim sFileName As String
    Dim sUniqueNo As Long
    
    sFileName = Space(255)
    If Len(ThreeCharPrefix) < 3 Then
        ThreeCharPrefix = ThreeCharPrefix & String(3 - Len(ThreeCharPrefix), "_")
    End If
    
    sUniqueNo = GetTempFileName(shellGetPathTemp, ThreeCharPrefix, 0, sFileName)
    If sUniqueNo = 0 Then
        shellCreateTempFile = ""
    Else
        shellCreateTempFile = Left(sFileName, InStr(sFileName, Chr(0)) - 1)
    End If
    
End Function


Public Function ctlGetEditStyle( _
                                ByVal hWnd As Long) _
                As enumEditStyles
    
    ctlGetEditStyle = GetWindowLong(hWnd, GWL_STYLE)
    
End Function

Public Function ctlSetEditStyle( _
                                ByVal hWnd As Long, _
                                ByVal eStyle As enumEditStyles)
    
    Dim lCurrStyle As Long
    lCurrStyle = ctlGetEditStyle(hWnd)
    If (lCurrStyle And eStyle) = 0 Then
        SetWindowLong hWnd, GWL_STYLE, lCurrStyle Or eStyle
    End If
    
End Function
                                
Public Function ctlUnSetEditStyle( _
                                ByVal hWnd As Long, _
                                ByVal eStyle As enumEditStyles)
    
    Dim lCurrStyle As Long
    lCurrStyle = ctlGetEditStyle(hWnd)
    If (lCurrStyle And eStyle) <> 0 Then
        SetWindowLong hWnd, GWL_STYLE, lCurrStyle Xor eStyle
    End If
    
End Function
                                



Public Function folderGet( _
                        ByVal hWndOwner As Long, _
                        ByVal sPrompt As String, _
                        Optional ByVal iOptions As enumBrowseForFlags = BIF_RETURNONLYFSDIRS) _
                As String
    '
    ' Opens the system dialog for browsing for a folder.
    '
    Dim iNull    As Integer
    Dim lpIDList As Long
    Dim lResult  As Long
    Dim sPath    As String
    Dim udtBI    As BrowseInfo
    
    If iOptions = BIF_NONESPECIFIED Then
        iOptions = BIF_RETURNONLYFSDIRS
    End If
    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = iOptions
'        .pszDisplayName = VarPtr(sDefaultPath)
    End With
    
    lpIDList = SHBrowseForFolder(udtBI)
    
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
    End If
    
    folderGet = sPath
    
End Function


Public Function shellCopy( _
                        ByVal sDirOrFileNameFrom As String, _
                        ByVal sDirOrFileNameTo As String, _
                        Optional ByVal CopyFlags As enumFileOpFlags = FOF_ALLOWUNDO) _
                As Boolean
    
    Dim lResult  As Long
    Dim lFlags   As Long
    Dim SHFileOp As SHFILEOPSTRUCT
    
    Screen.MousePointer = vbHourglass
    
    sDirOrFileNameFrom = RTrim(LTrim(sDirOrFileNameFrom))
    If Len(sDirOrFileNameFrom) > 3 And Right(sDirOrFileNameFrom, 1) = "\" Then
        sDirOrFileNameFrom = Left(sDirOrFileNameFrom, Len(sDirOrFileNameFrom) - 1)
    End If
    sDirOrFileNameTo = RTrim(LTrim(sDirOrFileNameTo))
    If Len(sDirOrFileNameTo) > 3 And Right(sDirOrFileNameTo, 1) = "\" Then
        sDirOrFileNameTo = Left(sDirOrFileNameTo, Len(sDirOrFileNameTo) - 1)
    End If
    
    With SHFileOp
        .wFunc = FO_COPY
        .pFrom = sDirOrFileNameFrom & vbNullChar & vbNullChar
        .pTo = sDirOrFileNameTo & vbNullChar & vbNullChar
        .fFlags = CopyFlags
    End With
    lResult = SHFileOperation(SHFileOp)
    '
    ' If User hit Cancel button while operation is in progress,
    ' the fAborted parameter will be true
    '
    Screen.MousePointer = vbDefault
    If lResult <> 0 Or SHFileOp.fAborted Then
        shellCopy = False
    Else
        shellCopy = True
    End If
    
End Function

Public Function shellDelete( _
                            ByVal sDirOrFileName As String, _
                            Optional ByVal DeleteFlags As enumFileOpFlags = FOF_ALLOWUNDO) _
                As Boolean
    
    Dim lResult  As Long
    Dim lFlags   As Long
    Dim SHFileOp As SHFILEOPSTRUCT
    
    Screen.MousePointer = vbHourglass
    
    sDirOrFileName = RTrim(LTrim(sDirOrFileName))
    If Len(sDirOrFileName) > 3 And Right(sDirOrFileName, 1) = "\" Then
        sDirOrFileName = Left(sDirOrFileName, Len(sDirOrFileName) - 1)
    End If
    
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sDirOrFileName & vbNullChar & vbNullChar
        .pTo = vbNullChar & vbNullChar
        .fFlags = DeleteFlags
    End With
    On Error Resume Next
    lResult = SHFileOperation(SHFileOp)
    On Error GoTo 0
    '
    ' If User hit Cancel button while operation is in progress,
    ' the fAborted parameter will be true
    '
    Screen.MousePointer = vbDefault
    If lResult <> 0 Or SHFileOp.fAborted Then
        shellDelete = False
    Else
        shellDelete = True
    End If
    
End Function


Public Function shellMove( _
                        ByVal sDirOrFileNameFrom As String, _
                        ByVal sDirOrFileNameTo As String, _
                        Optional ByVal MoveFlags As enumFileOpFlags = FOF_ALLOWUNDO + FOF_RENAMEONCOLLISION, _
                        Optional ByRef sNewFileName As String) _
                As Boolean
    
    Dim lResult  As Long
    Dim lFlags   As Long
    Dim SHFileOp As SHFILEOPSTRUCT
    
    Screen.MousePointer = vbHourglass
    
    sDirOrFileNameFrom = RTrim(LTrim(sDirOrFileNameFrom))
    If Len(sDirOrFileNameFrom) > 3 And Right(sDirOrFileNameFrom, 1) = "\" Then
        sDirOrFileNameFrom = Left(sDirOrFileNameFrom, Len(sDirOrFileNameFrom) - 1)
    End If
    sDirOrFileNameTo = RTrim(LTrim(sDirOrFileNameTo))
    If Len(sDirOrFileNameTo) > 3 And Right(sDirOrFileNameTo, 1) = "\" Then
        sDirOrFileNameTo = Left(sDirOrFileNameTo, Len(sDirOrFileNameTo) - 1)
    End If
    
    With SHFileOp
        .wFunc = FO_MOVE
        .pFrom = sDirOrFileNameFrom & vbNullChar & vbNullChar
        .pTo = sDirOrFileNameTo & vbNullChar & vbNullChar
        .fFlags = MoveFlags
    End With
    lResult = SHFileOperation(SHFileOp)
    '
    ' If User hit Cancel button while operation is in progress,
    ' the fAborted parameter will be true
    '
    Screen.MousePointer = vbDefault
    sNewFileName = SHFileOp.pTo
    If lResult <> 0 Or SHFileOp.fAborted Then
        shellMove = False
    Else
        shellMove = True
    End If
    
End Function

Public Function shellRename( _
                            ByVal sDirOrFileNameFrom As String, _
                            ByVal sDirOrFileNameTo As String, _
                            Optional ByVal RenameFlags As enumFileOpFlags = FOF_ALLOWUNDO + FOF_RENAMEONCOLLISION, _
                            Optional ByRef sNewFileName As String) _
                As Boolean
    
    Dim lResult  As Long
    Dim lFlags   As Long
    Dim SHFileOp As SHFILEOPSTRUCT
    
    Screen.MousePointer = vbHourglass
    
    sDirOrFileNameFrom = RTrim(LTrim(sDirOrFileNameFrom))
    If Len(sDirOrFileNameFrom) > 3 And Right(sDirOrFileNameFrom, 1) = "\" Then
        sDirOrFileNameFrom = Left(sDirOrFileNameFrom, Len(sDirOrFileNameFrom) - 1)
    End If
    sDirOrFileNameTo = RTrim(LTrim(sDirOrFileNameTo))
    If Len(sDirOrFileNameTo) > 3 And Right(sDirOrFileNameTo, 1) = "\" Then
        sDirOrFileNameTo = Left(sDirOrFileNameTo, Len(sDirOrFileNameTo) - 1)
    End If
    
    With SHFileOp
        .wFunc = FO_RENAME
        .pFrom = sDirOrFileNameFrom & vbNullChar & vbNullChar
        .pTo = sDirOrFileNameTo & vbNullChar & vbNullChar
        .fFlags = RenameFlags
    End With
    lResult = SHFileOperation(SHFileOp)
    '
    ' If User hit Cancel button while operation is in progress,
    ' the fAborted parameter will be true
    '
    Screen.MousePointer = vbDefault
    sNewFileName = SHFileOp.pTo
    If lResult <> 0 Or SHFileOp.fAborted Then
        shellRename = False
    Else
        shellRename = True
    End If
    
End Function



Public Function odbcConnectString( _
                                ByVal sDSN As String, _
                                Optional ByVal sUser As String = "", _
                                Optional ByVal sPWD As String = "") _
                As String
    
    odbcConnectString = "ODBC;DSN=" & sDSN & ";UID=" & sUser & ";PWD=" & sPWD & ";"
    
End Function





Public Function windowCopyImage(ByVal hWndFrom As Long, ByVal hDCto As Long, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0)  ' ByRef objTo as Object)
    
    Dim hWnd As Long
    Dim tr As Rect
    Dim hdc As Long
    
    ' Note: objTo must have hDC,Picture,Width and Height
    ' properties and should have AutoRedraw = True
    
    ' Get the size of the desktop window:
    
    'hwnd = GetDesktopWindow()
    GetWindowRect hWndFrom, tr
    
    ' Set the object to the relevant size:
'    objTo.Width = (tR.Right - tR.Left) * Screen.TwipsPerPixelX
'    objTo.Height = (tR.Bottom - tR.Top) * Screen.TwipsPerPixelY
    
    ' Now get the desktop DC:
    hdc = GetDC(hWndFrom)
    'hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    ' Copy the contents of the desktop to the object:
    'BitBlt objTo.hDC, 0, 0, (tR.Right - tR.Left), (tR.Bottom - tR.Top), hDC, 0, 0, SRCCOPY
    Debug.Print "Left = " & Left & "   Top = " & Top
    BitBlt hDCto, 0, 0, (tr.Right - tr.Left), (tr.Bottom - tr.Top), hdc, Left, Top, SRCCOPY
    ' Ensure we clear up DC GDI has given us:
    DeleteDC hdc
    
End Function


Public Function fileAttributesEx(ByVal sFileName As String) As WIN32_FILE_ATTRIBUTE_DATA
    Dim tRet As WIN32_FILE_ATTRIBUTE_DATA
    GetFileAttributesEx sFileName, 1, tRet
    fileAttributesEx = tRet
End Function


Public Function timeETC( _
                        ByVal dStart As Date, _
                        ByVal lMax As Long, _
                        ByVal lCurr As Long, _
                        Optional ByVal eFormat As enumTimeElapsedFormats = etfHoursMinutesSeconds) _
                As String
                
    Dim lElapsed As Long
    lElapsed = DateDiff("s", dStart, Now)
    Dim dPer As Double
    dPer = lElapsed / lCurr
    Dim lRemaining As Long
    lRemaining = dPer * (lMax - lCurr)
    timeETC = ts.timeAsElapsed(ts.timeFromSeconds(lRemaining), eFormat)
    
End Function




Public Function picFromBmp( _
                            ByVal hBmp As Long, _
                            Optional ByVal hPal As Long = 0) _
                As IPicture
    
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
    Set picFromBmp = IPic
    
End Function


Public Function picMappedFromRes( _
                                ByVal lResBmpID As Long, _
                                ByVal lColorFrom1 As Long, _
                                ByVal lColorTo1 As Long, _
                                Optional ByVal lColorFrom2 As Long = 0, _
                                Optional ByVal lColorTo2 As Long = 0) As IPicture
    
    Dim aMap(0 To 1) As COLORMAP
    aMap(0).from = lColorFrom1
    aMap(0).to = lColorTo1
    aMap(1).from = lColorFrom2
    aMap(1).to = lColorTo2
    
    
    
    Dim cMap As twoCOLORMAPs
    cMap.from1 = lColorFrom1
    cMap.to1 = lColorTo1
    cMap.from2 = lColorFrom2
    cMap.to2 = lColorTo1
    
    cMap.from1 = &HFFFFFF
    cMap.to1 = ecOrange
    cMap.from2 = &HFFFF ' vbYellow
    cMap.to2 = vbMagenta
    
    Dim c1Map As COLORMAP
    c1Map.from = vbYellow
    c1Map.to = vbGreen
    
    Set picMappedFromRes = picFromBmp(CreateMappedBitmap(App.hInstance, lResBmpID, 0, c1Map, 1))
    
End Function


Public Function txtUnselected(ByRef txtCtl As Object) As String
    Dim sReturn As String
    sReturn = txtCtl.Text
    If txtCtl.SelLength > 0 Then
        sReturn = Left(sReturn, txtCtl.SelStart) & Mid(sReturn, txtCtl.SelStart + txtCtl.SelLength + 1)
    End If
    txtUnselected = sReturn
End Function


Public Function FormAlwaysOnTop(ByRef FormToWorkWith As Form, ByVal YesorNo As Boolean)
    
    Dim lReturn As Long
    If YesorNo Then
        lReturn = SetWindowPos(FormToWorkWith.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
        If lReturn = 0 Then
            ' We had an error of some kind
        End If
    Else
        lReturn = SetWindowPos(FormToWorkWith.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    End If
        
End Function


Public Function windowSetForeground(ByVal hWnd As Long) As Boolean
    Dim lReturn As Long
    lReturn = SetForegroundWindow(hWnd)
End Function


Public Function pbStart(ByRef pb As Object, ByVal lMax)
    pb.Value = 0
    pb.Min = 0
    pb.max = lMax
    
End Function


Public Function dayOfYear( _
                        Optional ByVal dDate As Date) _
                As Integer
    
    If dDate = 0 Then
        dDate = Date
    End If
    dayOfYear = dDate - CDate("01/01/" & Year(dDate)) + 1
    
End Function


#If bWindowCls Then
Public Function formCenterInScreen(ByRef frm As Form)
    
    Dim rec As Rect
    
    
    
    
End Function

Public Function formFillScreen(ByRef frm)
    With ts.shellGetDesktopRectTwips
        frm.WindowState = vbNormal
        frm.Move .Left, .Top, .Right - .Left, .Bottom - .Top
    End With
    
End Function


#End If


Public Function sLTrim( _
                        ByVal sString As String, _
                        ByVal sStringToTrim As String) _
                As String
    
    Do While Left(sString, Len(sStringToTrim)) = sStringToTrim
        sString = Mid(sString, Len(sStringToTrim) + 1)
    Loop
    sLTrim = sString
    
End Function


Public Function dateSFormat(ByVal sDateString As String, ByVal sFormat As String) As String
    Select Case UCase(sFormat)
        Case "MM/DD/YYYY"
            
            Dim vNumbers As Variant
            vNumbers = Split(sWord(sDateString, 1), "/")
            If UBound(vNumbers) = 2 Then
                Select Case True
                    Case Len(vNumbers(0)) = 4
                        sDateString = sPadL(vNumbers(1), 2, "0") & "/" & sPadL(vNumbers(2), 2, "0") & "/" & Format(dateY2KYear(vNumbers(0)), "0000")
                    Case Else
                        sDateString = sPadL(vNumbers(0), 2, "0") & "/" & sPadL(vNumbers(1), 2, "0") & "/" & Format(dateY2KYear(vNumbers(2)), "0000")
                End Select
            End If
        Case Else
            MsgBox "dateSFormat - Unsupported format string"
    End Select
    dateSFormat = sDateString
    
End Function

Public Function dateY2KYear(ByVal lYear As Variant) As Long
    lYear = CLng(lYear)
    Select Case True
        Case lYear < 20
            lYear = lYear + 2000
        Case lYear < 100
            lYear = lYear + 1900
    End Select
    dateY2KYear = lYear
    
End Function



Public Function dateBOW( _
                        ByVal dDate As Date, _
                        Optional WeekStart As VbDayOfWeek = vbSunday) _
                As Date
    
    dateBOW = dDate - (Weekday(dDate, WeekStart) - 1)
    
End Function

Public Function dateEOW( _
                        ByVal dDate As Date, _
                        Optional WeekStart As VbDayOfWeek = vbSunday) _
                As Date
    
    dateEOW = dDate + (7 - Weekday(dDate, WeekStart))
    
End Function


Public Function sysSetLocaleInfo( _
                                ByVal eInfoType As enumLocaleInfoTypes, _
                                ByVal sValue As String) _
                As Boolean
    
    Dim lID As Long
    lID = GetSystemDefaultLCID()
    
    sysSetLocaleInfo = (SetLocaleInfo(lID, eInfoType, sValue) <> 0)
    
End Function

Public Function sysGetLocaleInfo( _
                                ByVal eInfoType As enumLocaleInfoTypes) _
                As String
    
    Dim lID As Long
    lID = GetSystemDefaultLCID()
    
    Dim sValue As String
    sValue = Space(1024)
    If GetLocaleInfo(lID, eInfoType, sValue, Len(sValue)) Then
        sValue = ts.sNT(sValue)
    End If
    
    sysGetLocaleInfo = Trim(sValue)
        
End Function



Public Function sysForceEnglishDate() As Boolean
    Dim bRet As Boolean
    bRet = True
    If ts.sysGetLocaleInfo(LOCALE_SDATE) <> "/" Then
        bRet = sysSetLocaleInfo(LOCALE_SDATE, "/")
    End If
    If bRet Then
        If UCase(Trim(sysGetLocaleInfo(LOCALE_SSHORTDATE))) <> "MM/DD/YYYY" Then
            bRet = sysSetLocaleInfo(LOCALE_SSHORTDATE, "MM/dd/yyyy")
        End If
    End If
    sysForceEnglishDate = bRet
    
End Function



#If bRegistryCls And bSHFolderDLL Then

Public Function odbcCreateSQLServer2000DSN( _
                                            ByVal DSNName As String, _
                                            ByVal DSNServer As String, _
                                            ByVal DSNDatabase As String, _
                                            ByVal DSNLogin As String, _
                                            Optional ByVal DSNDesc As String, _
                                            Optional ByVal DSNPassword As String) _
                As Boolean
    
    Dim cReg As cRegistry
    Set cReg = New cRegistry
    
    cReg.ClassKey = HKEY_LOCAL_MACHINE
    cReg.SectionKey = "SOFTWARE\ODBC\ODBC.INI\" & DSNName
    If cReg.KeyExists Then
        cReg.DeleteKey
    End If
    cReg.CreateKey
    cReg.ValueType = REG_SZ
    cReg.ValueKey = "Driver"
    cReg.Value = ts.shellGetPathFor(CSIDL_SYSTEM) & "SQLSRV32.DLL"
    If Trim(DSNDesc) <> "" Then
        cReg.ValueKey = "Description"
        cReg.Value = DSNDesc
    End If
    If Trim(DSNServer) <> "" Then
        cReg.ValueKey = "Server"
        cReg.Value = DSNServer
    End If
    If Trim(DSNDatabase) <> "" Then
        cReg.ValueKey = "Database"
        cReg.Value = DSNDatabase
    End If
    If Trim(DSNLogin) <> "" Then
        cReg.ValueKey = "LastUser"
        cReg.Value = DSNLogin
    End If
    If Trim(DSNPassword) <> "" Then
        cReg.ValueKey = "LastPassword"
        cReg.Value = DSNPassword
    End If
    
    Set cReg = Nothing
    odbcCreateSQLServer2000DSN = True
    
End Function



#End If


Public Function dcSetFont( _
                        ByVal hdc As Long, _
                        ByVal oFont As StdFont, _
                        Optional ByVal lFontColor As OLE_COLOR = -99, _
                        Optional ByVal eAlignment As enumTextAlignment = -99)
    
    Dim oLogFont As LOGFONT
    With oLogFont
        .lfCharSet = oFont.Charset
        Dim l As Long
        For l = 1 To Len(oFont.Name)
            .lfFaceName(l) = Asc(Mid(oFont.Name, l, 1))
        Next l
        .lfHeight = -MulDiv(GetDeviceCaps(hdc, LOGPIXELSY), oFont.Size, 72)
        .lfItalic = oFont.Italic
        .lfStrikeOut = oFont.Strikethrough
        .lfUnderline = oFont.Underline
        .lfWeight = oFont.Weight
        .lfQuality = 2
        
    End With
    
    SelectObject hdc, CreateFontIndirect(oLogFont)
    
    If lFontColor <> -99 Then
        dcSetTextColor hdc, lFontColor
    End If
    If eAlignment <> -99 Then
        dcSetTextAlignment hdc, eAlignment
    End If
    
End Function

Public Function dcGetFont(ByVal hdc As Long) As StdFont
    
    Set dcGetFont = New StdFont
    
    Dim sFontName As String
    sFontName = Space(1024)
    GetTextFace hdc, Len(sFontName), sFontName
    
    Dim tMetrics As TEXTMETRIC
    GetTextMetrics hdc, tMetrics
    
    Dim dY As Long
    dY = GetDeviceCaps(hdc, LOGPIXELSY)
    
    With dcGetFont
        .Name = sNT(sFontName)
        .Bold = (tMetrics.tmWeight = 700)
        .Charset = tMetrics.tmCharSet
        .Italic = tMetrics.tmItalic
        .Size = (tMetrics.tmAscent / tMetrics.tmDigitizedAspectY) * 72
'        Select Case True
'            Case tMetrics.tmPitchAndFamily And TMPF_FIXED_PITCH
'                .Size = (tMetrics.tmAscent / tMetrics.tmDigitizedAspectY) * 72
'            Case tMetrics.tmPitchAndFamily And TMPF_TRUETYPE
'                .Size = (tMetrics.tmAscent + tMetrics.tmInternalLeading) / 2
'        End Select
        .Strikethrough = tMetrics.tmStruckOut
        .Underline = tMetrics.tmUnderlined
        .Weight = tMetrics.tmWeight
    End With
    
End Function

Public Function dcGetTextColor(ByVal hdc As Long) As OLE_COLOR
    dcGetTextColor = GetTextColor(hdc)
End Function

Public Function dcSetTextColor(ByVal hdc As Long, ByVal lColor As OLE_COLOR)
    SetTextColor hdc, lColor
End Function

Public Function dcSetTextAlignment(ByVal hdc As Long, ByVal eAlignment As enumTextAlignment)
    SetTextAlign hdc, eAlignment
    
End Function

Public Function dcGetTextAlignment(ByVal hdc As Long) As enumTextAlignment
    dcGetTextAlignment = GetTextAlign(hdc)
End Function

Public Function vbAlignToTextAlign(ByVal eAlignment As AlignmentConstants) As enumTextAlignment
    Select Case True
        Case eAlignment = vbCenter
            vbAlignToTextAlign = TA_CENTER
        Case eAlignment = vbLeftJustify
            vbAlignToTextAlign = TA_LEFT
        Case eAlignment = vbRightJustify
            vbAlignToTextAlign = TA_RIGHT
        Case Else
            vbAlignToTextAlign = TA_LEFT
    End Select
End Function
