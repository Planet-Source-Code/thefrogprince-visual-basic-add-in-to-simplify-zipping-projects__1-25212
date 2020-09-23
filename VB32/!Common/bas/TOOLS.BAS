Attribute VB_Name = "Tools"
Option Explicit
'**************
'* Cool Stuff *
'**************

' First exception
Private Declare Function GetInputState Lib "user32" () As Long


'------------------------------------------------------
' A list of possible key states for SHIFT
'''''''Public Enum enumSHIFTVALUES
'''''''    svShiftKey = 0
'''''''    svControlKey = 1
'''''''    svAltKey = 2
'''''''End Enum




'------------------------------------------------------
' A list of possible types of input (used with the EditMask function)
Public Enum enumType
    etDATE_SLASH = 1
    etPOS_NUM = 2
    etLETTER = 3
    etTIME = 4
End Enum


'''''
''''''------------------------------------------------------
'''''' A list of possible ways to check a string.
'''''Public Enum StringIsTypes
'''''    OnlyNumbers = 1
'''''    NumbersAndMoneyPunctuation = 2
'''''    NumbersAndNumericPunctuation = 3
'''''    MathematicalCharacters = 4
'''''
'''''End Enum



'------------------------------------------------------
' A list of properties that may need to be tested for.
' Don't see one you need?  Just add it to the bottom and use the next number.
' Then... code the test in the function ItemHasProperty()...
' And then return... if you want it... in GetProperty().
' Because everything is variant... we can still get away with
' this in TOOLS.bas.
Public Enum PotentialControlProperties
    Justification = 1
    MaxCharacters = 2
    MaxLength = 3
    Text = 4
    SelBookmarks = 5
    LockEdits = 6
    Connection = 7
    Caption = 8
    Count = 9
    Connect = 10
    Name = 11
    DateCreated = 12
    LastModified = 13
    LastUpdated = 14
    Enabled = 15
    ListCount = 16
    ListIndex = 17
    SimpleText = 18
    Panels = 19
    Style = 20
    List = 21
End Enum

'------------------------------------------------------
' Jeff's very early proto-type for a SQL parser.  =o)
' Goes with the function ParseSQLSelect()
Public Type SQLSelectObject
    Fields As String
    From As String
    Where As String
    OrderBy As String
    Having As String
    GroupBy As String
End Type

''''''------------------------------------------------------
'''''' List of your primary RGB colors.
'''''' Goes with the function GetRGBValue()
'''''Public Enum RGBBaseColors
'''''    rgbRed = 1
'''''    rgbBlue = 2
'''''    rgbGreen = 3
'''''End Enum

' KeyAsciiAllowTypes has been renamed and moved to mAPIconstants.
'''------------------------------------------------------
''' List of different styles of keyboard entry allowed.
''' Goes with the function KeyAsciiAllow()
''Public Enum KeyAsciiAllowTypes
''    NumbersOnly = 1
''    Uppercase = 2
''    NoSpaces = 4
''    NoSingleQuotes = 8
''    NoDoubleQuotes = 16
''    AllowDecimal = 32
''    AllowNegative = 64
''End Enum


'------------------------------------------------------
' List of different kinds of alignment that can take place between objects in a screen.
Public Enum AlignmentConstantsNEW
    vbLeft = 1
    vbTop = 2
    vbright = 4
    vbBottom = 8
    vbMiddleVertical = 16
    vbMiddleHorizontal = 32
    vbRightFlush = 64
    vbLeftFlush = 128
End Enum

'------------------------------------------------------
' List of different kinds of centering that can take place between objects in a screen.
Public Enum AlignmentCenteringConstants
    CenterHorizontally = vbHorizontal
    CenterVertically = vbVertical
    Center = vbHorizontal + vbVertical
End Enum

'

'------------------------------------------------------
' A public property of the TOOLS module that sets what the dividing line is for the never ending Y2K issue.  =o)
Public Y2Kthreshold As Integer     ' Number to divide Y2K issues on (default is 40)


'------------------------------------------------------
' Need a public property indicating whether the GradientColorBackground() function is currently being run.
' This switch prevents it from running multiple instances of itself.
Public bDoingGradientBackGroundFill As Boolean


'------------------------------------------------------
' Used with the SaveObjectCoordinates and RestoreObjectCoordinates functions to
' quickly backup an objects positioning properties, allow modifications, and then
' quick restoration to the original state.
'''Public Type ObjectCoordinates
'''    objLeft As Long
'''    objTop As Long
'''    objWidth As Long
'''    objHeight As Long
'''
'''End Type


Public Type ManualGridBookmark
    FirstRow As Variant
    Row As Long
End Type






' ==================================================
' BitIsOn()
'   Returns whether or not a certain bit in a byte is set on.  Possible bits are:
'   87654321
' ==================================================
'''''Public Function BitIsOn(ByVal ValToCheck As Integer, BitToCheck As Variant) As Boolean
'''''
'''''    Select Case BitToCheck
'''''        Case 1
'''''            BitIsOn = (ValToCheck And 1) <> 0
'''''        Case 2
'''''            BitIsOn = (ValToCheck And 2) <> 0
'''''        Case 3
'''''            BitIsOn = (ValToCheck And 4) <> 0
'''''        Case 4
'''''            BitIsOn = (ValToCheck And 8) <> 0
'''''        Case 5
'''''            BitIsOn = (ValToCheck And 16) <> 0
'''''        Case 6
'''''            BitIsOn = (ValToCheck And 32) <> 0
'''''        Case 7
'''''            BitIsOn = (ValToCheck And 64) <> 0
'''''        Case 8
'''''            BitIsOn = (ValToCheck And 128) <> 0
'''''        Case 9
'''''            BitIsOn = (ValToCheck And 256) <> 0
'''''        Case 10
'''''            BitIsOn = (ValToCheck And 512) <> 0
'''''        Case 11
'''''            BitIsOn = (ValToCheck And 1024) <> 0
'''''        Case 12
'''''            BitIsOn = (ValToCheck And 2048) <> 0
'''''        Case 13
'''''            BitIsOn = (ValToCheck And 4096) <> 0
'''''        Case 14
'''''            BitIsOn = (ValToCheck And 8192) <> 0
'''''        Case 15
'''''            BitIsOn = (ValToCheck And 16384) <> 0
'''''        Case 16
'''''            BitIsOn = (ValToCheck And 32768#) <> 0
'''''        Case Else
'''''            BitIsOn = False
'''''    End Select
'''''
'''''End Function


''''''' ==================================================
''''''' FileExists()
'''''''    Returns whether a file exists or not.
''''''' ==================================================
''''''Public Function fileExists(ByVal File2Look4 As String) As Boolean
''''''
''''''    Dim FileHandle As Integer
''''''
''''''    FileHandle = FreeFile
''''''
''''''    Err.Clear
''''''    On Error Resume Next
''''''    Open File2Look4 For Input As #FileHandle
''''''    fileExists = (Err.Number = 0)
''''''    Close #FileHandle
''''''    On Error GoTo 0
''''''
''''''    ' Old Method Follows
'''''''    If Not IsBlank(File2Look4) Then
'''''''        On Error Resume Next
'''''''        FileExists = (Not Dir(Trim(File2Look4), vbNormal + vbHidden + vbSystem + vbDirectory) = "") And (FileLen(File2Look4) <> 0)
'''''''        If Err.Number <> 0 Then
'''''''            FileExists = False
'''''''        End If
'''''''    Else
'''''''        FileExists = False
'''''''    End If
''''''
''''''End Function


' ==================================================
' FnPadZees()
'   Returns whether or not a certain bit in a byte is set on.  Possible bits are:
'   87654321
' ==================================================
''''''Public Function FnPadZees(ByRef sInstring As String, ByRef iFieldlen As Integer) As String
''''''    ' Author: Mark
''''''    Dim iStringLen As Integer
''''''    iStringLen = Len(Trim$(sInstring))
''''''    If iStringLen < iFieldlen Then
''''''        FnPadZees = sInstring & String(iFieldlen - iStringLen, "Z")
''''''    Else
''''''        FnPadZees = Left$(sInstring, iFieldlen - 1) & "Z"
''''''    End If
''''''End Function


''''Public Function InLine(ByVal String2Search As String, ByVal String2SearchFor As String) As Integer
''''
''''    Dim i As Integer
''''    Dim cntLines As Integer
''''    cntLines = Tools.LineCnt(String2Search)
''''
''''    InLine = 0
''''    For i = 1 To cntLines
''''        If InStr(UCase(Tools.GetLine(String2Search, i)), UCase(String2SearchFor)) > 0 Then
''''            InLine = i
''''            Exit Function
''''        End If
''''    Next i
''''
''''End Function

'''''' ==================================================
'''''' InTrim()
''''''    Returns a string with all redundant spaces reduced to a single space.
''''''    Alternate character(s) may be specified for the space.
'''''' ==================================================
'''''Public Function InTrim(ByVal InputString As String, Optional ByVal AlternateSpaceCharacter As String) As String
'''''    Dim SpaceIsOn As Boolean
'''''    Dim StringLength As Integer
'''''    Dim i As Integer
'''''    Dim outString As String
'''''
'''''    outString = ""
'''''
'''''    If IsBlank(AlternateSpaceCharacter) Then
'''''        AlternateSpaceCharacter = " "
'''''    End If
'''''
'''''    InputString = AlternateSpaceCharacter & InputString & AlternateSpaceCharacter
'''''    SpaceIsOn = True
'''''    StringLength = Len(InputString)
'''''
'''''    For i = 1 To StringLength
'''''        If InStr(AlternateSpaceCharacter, Mid(InputString, i, 1)) > 0 Then
'''''            If Not SpaceIsOn Then
'''''                outString = outString & Mid(InputString, i, 1)
'''''                SpaceIsOn = True
'''''            End If
'''''        Else
'''''            SpaceIsOn = False
'''''            outString = outString & Mid(InputString, i, 1)
'''''        End If
'''''    Next
'''''
'''''    If InStr(AlternateSpaceCharacter, Right(outString, 1)) > 0 And Not outString = "" Then
'''''        outString = Left(outString, Len(outString) - 1)
'''''    End If
'''''
'''''    InTrim = outString
'''''
'''''End Function

''''' ==================================================
''''' IsBlank()
'''''    Returns whether something is blank or not.
'''''    Checks for NULL and zero length strings.
''''' ==================================================
''''Public Function IsBlank(ByVal VarToCheck As Variant) As Boolean
''''
''''    IsBlank = False
''''    If IsMissing(VarToCheck) Then
''''        IsBlank = True
''''        Exit Function
''''    End If
''''    If IsNull(VarToCheck) Then
''''        IsBlank = True
''''        Exit Function
''''    End If
''''    If IsEmpty(VarToCheck) Then
''''        IsBlank = True
''''        Exit Function
''''    End If
''''    If Len(Trim(VarToCheck)) = 0 Then
''''        IsBlank = True
''''        Exit Function
''''    End If
''''
''''End Function


'''''' ==================================================
'''''' Padl()
''''''    Pads a string on the left side to the specified string length using spaces or the optionally supplied alternate space character.
'''''' ==================================================
'''''Function Padl(ByVal InputString As String, LengthToPadTo As Integer, Optional PaddingCharacter As String) As String
'''''    Dim ReturnString As String
'''''    If Len(PaddingCharacter) = 0 Then
'''''        PaddingCharacter = " "
'''''    End If
'''''
'''''    ReturnString = Replicate(PaddingCharacter, LengthToPadTo - Len(InputString)) & InputString
'''''
'''''    Padl = ReturnString
'''''
'''''End Function
    
'''''Function PadLeft(ByVal InputString As String, LengthToPadTo As Integer, Optional PaddingCharacter As String) As String
'''''
'''''    PadLeft = Padl(InputString, LengthToPadTo, PaddingCharacter)
'''''
'''''End Function

''''Function PadRight(ByVal InputString As String, LengthToPadTo As Integer, Optional PaddingCharacter As String) As String
''''    Select Case True
''''        Case Len(RTrim(InputString)) > LengthToPadTo
''''            PadRight = Left(InputString, LengthToPadTo)
''''        Case IsMissing(PaddingCharacter)
''''            PadRight = Format(RTrim(InputString), "!" & String(LengthToPadTo, "@"))
''''        Case Len(PaddingCharacter) = 0
''''            PadRight = Format(RTrim(InputString), "!" & String(LengthToPadTo, "@"))
''''        Case Len(PaddingCharacter) <> 0
''''            ' Could possible speed up using Marks logic in PadZs
''''            Dim RetStr As String
''''            RetStr = Format(RTrim(InputString), "!" & String(LengthToPadTo, "@"))
''''            Dim i As Integer
''''            For i = Len(RetStr) To 1 Step -1
''''                If Mid(RetStr, i, 1) = " " Then
''''                    Mid(RetStr, i, 1) = PaddingCharacter
''''                End If
''''            Next i
''''            PadRight = RetStr
''''        Case Else   ' LOGIC FALLTHRU -- DO NOTHING
''''            PadRight = InputString
''''    End Select
''''End Function

''''Public Function ReadFile(ByVal FileNameToRead As String) As String
''''
''''    Dim fhand As Integer
''''    Dim ReturnStr As String
''''    Dim CurrChar As String
''''    CurrChar = Space(1024 * 8)
''''    ReturnStr = ""
''''    fhand = FreeFile()
''''    On Error Resume Next
''''    Open FileNameToRead For Binary Access Read As #fhand
''''    If Err.Number = 0 Then
''''        On Error GoTo 0
''''
''''        Do While Not EOF(fhand)
''''            DoEvents
''''            Get fhand, , CurrChar
''''            ReturnStr = ReturnStr & CurrChar
''''        Loop
''''
''''    End If
''''    Close #fhand
''''    On Error GoTo 0
''''    ReadFile = ReturnStr
''''
''''End Function


'''''' ==================================================
'''''' Replicate()
''''''    Returns a string composed of the string passed repeated X number of times.
'''''' ==================================================
'''''Function Replicate(ByVal CharacterOrString As String, Times2Repeat As Integer) As String
'''''
'''''    Dim ReturnString As String
'''''    Dim i As Integer
'''''    ReturnString = ""
'''''
'''''    For i = 1 To Times2Repeat
'''''        ReturnString = ReturnString & CharacterOrString
'''''
'''''    Next
'''''    Replicate = ReturnString
'''''
'''''End Function

'''''' ==================================================
'''''' StrTran()
''''''    This function will take in a string and convert all occurrences of a character
''''''    or a string to an alternate character or string.
'''''' ==================================================
'''''Function StrTran(ByVal InputString As String, ByRef String2Look4 As String, ByRef String2Change2 As String) As String
'''''
'''''    Dim ReturnString As String
'''''    ReturnString = Replace(InputString, String2Look4, String2Change2)
''''''    Dim Location As Integer
''''''    ReturnString = ""
''''''    While InStr(InputString, String2Look4)
''''''        Location = InStr(InputString, String2Look4)
''''''        ReturnString = ReturnString & Left(InputString, Location - 1) & String2Change2
''''''        InputString = Mid(InputString, InStr(InputString, String2Look4) + Len(String2Look4))
''''''    Wend
''''''    If Not IsBlank(InputString) Then
''''''        ReturnString = ReturnString & InputString
''''''    End If
'''''    StrTran = ReturnString
'''''
'''''End Function

'''''' ==================================================
'''''' Word()
''''''    Returns a specific word from a string.  Default word identifier is a space.
''''''    Optional character or characters can be used in place of space.
'''''' ==================================================
'''''Public Function Word(ByVal InputString As String, ByVal WordNum As Integer, Optional ByVal AlternateSpaceCharacter As String = " ", Optional ByVal bMultipleKindsOfSpaceChars As Boolean = False, Optional ByVal bIgnoreCase As Boolean = True) As String
'''''
'''''    Dim SpaceIsOn As Boolean
'''''    Dim StringLength As Long
'''''    Dim i As Long
'''''    Dim outString As String
'''''    Dim CurrWord As Integer
'''''    Dim iMatchLen As Integer
'''''    If bMultipleKindsOfSpaceChars Then
'''''        iMatchLen = 1
'''''    Else
'''''        iMatchLen = Len(AlternateSpaceCharacter)
'''''    End If
'''''
'''''    outString = ""
'''''    CurrWord = 0
'''''
'''''    If IsBlank(AlternateSpaceCharacter) Then
'''''        AlternateSpaceCharacter = " "
'''''    End If
'''''
'''''    Dim sWorkString As String
'''''    InputString = AlternateSpaceCharacter & InputString & AlternateSpaceCharacter
'''''    sWorkString = InputString
'''''    'SpaceIsOn = True
'''''    StringLength = Len(sWorkString)
'''''
'''''    If bIgnoreCase Then
'''''        AlternateSpaceCharacter = UCase(AlternateSpaceCharacter)
'''''        sWorkString = UCase(sWorkString)
'''''    End If
'''''
'''''    For i = 1 To StringLength
'''''        If InStr(AlternateSpaceCharacter, Mid(sWorkString, i, iMatchLen)) > 0 Then
'''''            SpaceIsOn = True
'''''            i = i + iMatchLen - 1
'''''        Else
'''''            If SpaceIsOn Then
'''''                SpaceIsOn = False
'''''                CurrWord = CurrWord + 1
'''''                If CurrWord > WordNum Then
'''''                    Exit For
'''''                End If
'''''            End If
'''''            If CurrWord = WordNum Then
'''''                outString = outString & Mid(InputString, i, 1)
'''''            End If
'''''        End If
'''''    Next i
'''''
'''''    If Not IsBlank(outString) And InStr(AlternateSpaceCharacter, Right(outString, iMatchLen)) > 0 Then
'''''        outString = Left(outString, Len(outString) - iMatchLen)
'''''    End If
'''''
'''''    Word = outString
'''''
'''''End Function

'''''' ==================================================
'''''' WordCnt()
''''''    Returns the total number of words in a string.  Default word delimiter is a
''''''    space... but you can specify an optional character to be used as space or '
''''''    even multiple characters (i.e. commas and dashes).
'''''' ==================================================
'''''Public Function WordCnt(ByVal InputString As String, Optional ByVal AlternateSpaceCharacter As String = " ", Optional ByVal bMultipleKindsOfSpaceChars As Boolean = False, Optional ByVal bIgnoreCase As Boolean = True) As Integer
'''''
'''''    Dim SpaceIsOn As Boolean
'''''    Dim StringLength As Long
'''''    Dim i As Long
'''''    Dim TempWordCnt As Integer
'''''    Dim iMatchLen As Integer
'''''    If bMultipleKindsOfSpaceChars Then
'''''        iMatchLen = 1
'''''    Else
'''''        iMatchLen = Len(AlternateSpaceCharacter)
'''''    End If
'''''
'''''    SpaceIsOn = True
'''''    TempWordCnt = 0
'''''
'''''    Dim sWorkString As String
'''''    InputString = AlternateSpaceCharacter & InputString & AlternateSpaceCharacter
'''''    sWorkString = InputString
'''''    StringLength = Len(InputString)
'''''
'''''    If bIgnoreCase Then
'''''        AlternateSpaceCharacter = UCase(AlternateSpaceCharacter)
'''''        sWorkString = UCase(sWorkString)
'''''    End If
'''''
'''''    For i = 1 To StringLength
'''''        If InStr(AlternateSpaceCharacter, Mid(sWorkString, i, iMatchLen)) > 0 Then
'''''            If Not SpaceIsOn Then
'''''                SpaceIsOn = True
'''''                TempWordCnt = TempWordCnt + 1
'''''            End If
'''''            i = i + iMatchLen - 1
'''''        Else
'''''            SpaceIsOn = False
'''''        End If
'''''    Next
'''''
'''''    WordCnt = TempWordCnt
'''''
'''''End Function

' ==================================================
' GetLine()
'    Needs the WORD() function, returns a string containing the text of a
'    particular line number from a multi-line string.
' ==================================================
Public Function GetLine(ByVal StringToGetFrom As String, ByVal LineNumber As Integer) As String

    GetLine = sWord(StringToGetFrom, LineNumber, vbCrLf)

End Function

'''' ==================================================
'''' LineCnt()
''''    Returns the total number of lines in a multi-line string.
'''' ==================================================
'''Public Function LineCnt(ByVal StringToCountFrom As String) As Integer
'''
'''    LineCnt = WordCnt(StringToCountFrom, vbCrLf)
'''
'''End Function

''''''' ==================================================
''''''' ObjectIsSet()
'''''''     Returns whether the object you have passed has been set yet or not.
''''''' ==================================================
''''''Public Function ObjectIsSet(ByRef ObjToCheck As Object) As Boolean
''''''
''''''    Dim iTester As Integer
''''''    On Error GoTo IsNotSet
''''''    ObjectIsSet = Not (ObjToCheck Is Nothing)
''''''    On Error GoTo 0
''''''    'ObjectIsSet = True
''''''    Exit Function
''''''
''''''IsNotSet:
''''''    ObjectIsSet = False
''''''
''''''End Function

''''Public Function IsObjectSet(ByRef ObjToCheck As Object) As Boolean
''''
''''    IsObjectSet = ObjectIsSet(ObjToCheck)
''''
''''End Function


''''' ==================================================
''''' KillFile()
'''''     returns a logical on whether it was able to successfully
'''''     delete the file you passed.
''''' ==================================================
''''Public Function KillFile(ByVal File2Delete As String) As Boolean
''''
''''    On Error GoTo DidNotKill
''''    Kill File2Delete
''''    KillFile = True
''''    Exit Function
''''
''''DidNotKill:
''''    KillFile = False
''''
''''End Function





''''Public Sub Pause(ByVal Seconds As Double)
''''
''''    Dim iHoldTime As Single
''''    iHoldTime = Timer
''''    While Timer < iHoldTime + Seconds
''''        DoEvents
''''    Wend
''''
''''End Sub


'''''Public Function Occurs(ByVal String2Check As String, ByVal Char2Look4 As String) As Long
'''''
'''''    Dim CurrPlace As Long
'''''    Dim CurrCount As Long
'''''
'''''    CurrCount = 0
'''''    For CurrPlace = 1 To Len(String2Check)
'''''        If Mid(String2Check, CurrPlace, 1) = Char2Look4 Then
'''''            CurrCount = CurrCount + 1
'''''        End If
'''''    Next CurrPlace
'''''
'''''    Occurs = CurrCount
'''''
'''''End Function



''''Public Function GetExt(ByVal FullFileName As String) As String
''''
''''    Dim i As Integer
''''    Dim iStrLen As Integer
''''    Dim sOutStr As String
''''
''''    sOutStr = ""
''''    iStrLen = Len(FullFileName)
''''
''''    If InStr(FullFileName, ".") Then
''''        For i = iStrLen To 1 Step -1
''''            If Mid(FullFileName, i, 1) = "." Then
''''                Exit For
''''            End If
''''            sOutStr = Mid(FullFileName, i, 1) & sOutStr
''''        Next i
''''    End If
''''
''''    GetExt = sOutStr
''''
''''End Function

Public Function PrinterLineCnt(ByVal MemoText As String, ByRef PRN As Printer, Optional ByVal HorizontalOffset As Long) As Integer
        
    Dim LineCnt As Integer
    Dim i As Integer
    Dim iWordCnt As Integer
    Dim sLine As String
    Dim sPrevLine As String
    Dim lAvailWidth As Long
    
    If IsMissing(HorizontalOffset) Then
        lAvailWidth = PRN.ScaleWidth
    Else
        lAvailWidth = PRN.ScaleWidth - HorizontalOffset
    End If
    
    LineCnt = 0
    sLine = ""
    iWordCnt = sWordCnt(MemoText)
    
    For i = 1 To iWordCnt
        If PRN.TextWidth(sLine & sWord(MemoText, i)) < lAvailWidth Then
            If sLine = "" Then
                LineCnt = 1
                sLine = sWord(MemoText, 1)
            Else
                sLine = sLine & " " & sWord(MemoText, i)
            End If
        Else
            LineCnt = LineCnt + 1
            sLine = sWord(MemoText, i)
        End If
    Next i
    
    PrinterLineCnt = LineCnt
    
End Function

Public Function PrinterGetLine(ByVal MemoText As String, ByVal Lineno As Integer, ByRef PRN As Printer, Optional ByVal HorizontalOffset As Long)
    
    Dim LineCnt As Integer
    Dim i As Integer
    Dim iWordCnt As Integer
    Dim sLine As String
    Dim sPrevLine As String
    Dim lAvailWidth As Long
    
    If IsMissing(HorizontalOffset) Then
        lAvailWidth = PRN.ScaleWidth
    Else
        lAvailWidth = PRN.ScaleWidth - HorizontalOffset
    End If
    
    LineCnt = 0
    sLine = ""
    iWordCnt = sWordCnt(MemoText)
    
    PrinterGetLine = ""
    
    For i = 1 To iWordCnt
        If PRN.TextWidth(sLine & sWord(MemoText, i)) < lAvailWidth Then
            If sLine = "" Then
                LineCnt = 1
                sLine = sWord(MemoText, 1)
            Else
                sLine = sLine & " " & sWord(MemoText, i)
            End If
        Else
            If Lineno = LineCnt Then
                Exit For
            Else
                LineCnt = LineCnt + 1
                sLine = sWord(MemoText, i)
            End If
        End If
    Next i
    If Lineno = LineCnt Then
        PrinterGetLine = sLine
    End If

End Function

Public Sub AutoCenterToScreen(ByRef InForm As Form)
    If InForm.WindowState = 0 Then
        InForm.Move (Screen.Width / 2) - (InForm.Width / 2), (Screen.Height / 2) - (InForm.Height / 2)
    End If
End Sub

''''Public Function File_Ext(ByVal FileIn As String) As String
''''    File_Ext = GetFileExt(FileIn)
''''End Function
''''Public Function GetFileExt(ByVal FileName As String) As String
''''    Dim outString As String
''''    Dim i As Integer
''''
''''    outString = ""
''''
''''    For i = Len(FileName) To 1 Step -1
''''        If Mid(FileName, i, 1) = "." Then
''''            Exit For
''''        End If
''''        outString = Mid(FileName, i, 1) & outString
''''    Next i
''''
''''    GetFileExt = outString
''''
''''End Function


'''''Public Function wsStatus(ByRef WinSockCtl As Object) As String
'''''    Select Case WinSockCtl.state
'''''        Case 0
'''''            wsStatus = "Closed"
'''''        Case 1
'''''            wsStatus = "Open"
'''''        Case 2
'''''            wsStatus = "Listening"
'''''        Case 3
'''''            wsStatus = "Connection Pending"
'''''        Case 4
'''''            wsStatus = "Resolving Host"
'''''        Case 5
'''''            wsStatus = "Host Resolved"
'''''        Case 6
'''''            wsStatus = "Connecting"
'''''        Case 7
'''''            wsStatus = "Connected"
'''''        Case 8
'''''            wsStatus = "Peer Is Closing The Connection"
'''''        Case 9
'''''            wsStatus = "Error"
'''''        Case Else
'''''            wsStatus = "UNKNOWN"
'''''    End Select
'''''End Function

''''Public Function GetFileName(ByVal FullFileName As String) As String
''''
''''    Dim iFileStart, iFileEnd As Integer
''''    Dim i As Integer
''''    For i = Len(FullFileName) To 1 Step -1
''''        Select Case Mid(FullFileName, i, 1)
''''            Case "."
''''                If iFileEnd = 0 Then
''''                    iFileEnd = i - 1
''''                End If
''''            Case "\"
''''                iFileStart = i + 1
''''                Exit For
''''        End Select
''''    Next i
''''    If iFileStart = 0 Then
''''        iFileStart = 1
''''    End If
''''    GetFileName = Mid(FullFileName, iFileStart, iFileEnd - iFileStart + 1)
''''End Function


'''Public Function GetFilePath(ByVal FullFileName As String) As String
'''    Dim i As Integer
'''    Dim lenStr As Integer
'''    Dim CurrChar As String
'''    lenStr = Len(FullFileName)
'''    For i = lenStr To 1 Step -1
'''        CurrChar = Mid(FullFileName, i, 1)
'''        Select Case CurrChar
'''            Case "\"
'''                Exit For
'''            Case ":"
'''                Exit For
'''        End Select
'''    Next i
'''    If i > 1 Then
'''        GetFilePath = Left(FullFileName, i)
'''    Else
'''        GetFilePath = ""
'''    End If
'''End Function


'''''''Public Function PixelsX(ByVal TwipsIn As Long) As Integer
'''''''    PixelsX = TwipsIn / Screen.TwipsPerPixelX
'''''''End Function
'''''''
'''''''Public Function PixelsY(ByVal TwipsIn As Long) As Integer
'''''''    PixelsY = TwipsIn / Screen.TwipsPerPixelY
'''''''End Function



Public Function NullCrypt(ByRef ShortEncryptedPassword As String) As String
    Dim outString As String
    Dim PwdLen As Integer
    Dim CurrChar As Integer
    outString = ""
    PwdLen = Len(ShortEncryptedPassword)
    For CurrChar = 1 To PwdLen
        outString = outString & Mid(ShortEncryptedPassword, CurrChar, 1) & Chr(0)
    Next CurrChar
    NullCrypt = outString
End Function

Public Function CleanCrypt(ByRef LongEncryptedPassword As String) As String
    
    Dim CurrChar As Integer
    Dim PwdLen As Integer
    Dim OddString As String
    Dim EvenString As String
    
    PwdLen = Len(LongEncryptedPassword)
    OddString = ""
    EvenString = ""
    
    CurrChar = 1
    Do While CurrChar <= PwdLen
        OddString = OddString & Mid(LongEncryptedPassword, CurrChar, 1)
        CurrChar = CurrChar + 1
        If CurrChar <= PwdLen Then
            EvenString = EvenString & Mid(LongEncryptedPassword, CurrChar, 1)
            CurrChar = CurrChar + 1
        End If
    Loop
    
    If Len(OddString) = Len(EvenString) And EvenString = sReplicate(Chr(0), Len(EvenString)) Then
        CleanCrypt = OddString
    Else
        CleanCrypt = LongEncryptedPassword
    End If
    
End Function


Public Function GetBinaryField(InField As Object) As String
    
    Dim PassString As String
    Dim CurrByte As Integer
    
    ' This logic must be included to read the password out of back end databases
    ' Default 32 bit behavior with the binary fields is put a bunch of nulls in it.
    PassString = ""
    CurrByte = 0
    Select Case True
        Case InField.Type = 9  ' BINARY FIELD DATA
            PassString = ""
            On Error Resume Next
            PassString = StrConv(InField, vbUnicode)
            On Error GoTo 0
            
            'On Error Resume Next
            'Do While True
            '    PassString = PassString & Chr(InField(CurrByte))
            '    If Err Then
            '        Exit Do
            '    End If
            '    CurrByte = CurrByte + 1
            'Loop
            'On Error GoTo 0
            'If Right(PassString, 1) = Chr(0) Then
            '    PassString = Left(PassString, Len(PassString) - 1)
            'End If
        Case Else
            PassString = CStr(InField.Value & "") & ""
    End Select
    
    GetBinaryField = PassString
    
End Function



Public Function InvertQuotes(ByVal StringToFix As String) As String
    StringToFix = Replace(StringToFix, "'", Chr(250))
    StringToFix = Replace(StringToFix, Chr(34), "'")
    StringToFix = Replace(StringToFix, Chr(250), Chr(34))
    InvertQuotes = StringToFix
End Function



'''''Public Function CharTrim(ByVal StringToTrim As String, ByRef CharacterToTrimOff As String) As String
'''''    While Left(StringToTrim, 1) = CharacterToTrimOff
'''''        StringToTrim = Mid(StringToTrim, 2)
'''''    Wend
'''''    While Right(StringToTrim, 1) = CharacterToTrimOff
'''''        StringToTrim = Left(StringToTrim, Len(StringToTrim) - 1)
'''''    Wend
'''''    CharTrim = StringToTrim
'''''End Function


Public Function sFormat(ByVal String2Format As String, TypeofFormat As String) As String
    Select Case UCase(TypeofFormat)
        Case "SHORT TIME"
            sFormat = Format(Left(String2Format, 2), "00") & ":" & Format(Right(String2Format, 2), "00")
        Case "/YYYYMMDD"
            sFormat = Mid(String2Format, 5, 2) & "/" & Mid(String2Format, 7, 2) & "/" & Left(String2Format, 4)
        Case Else
            MsgBox "sFormat() error!  UNSUPPORTED FORMAT!!"
            sFormat = String2Format
    End Select
End Function


Public Function SQLUpCase(ByRef StringToFix As String) As String
    
    Dim ReturnString As String
    Dim iCurrChar As Integer
    Dim CurrChar As String
    Dim StrLen As Integer
    Dim bQuotesOnSingle As Boolean
    Dim bQuotesOnDouble As Boolean
            
    StrLen = Len(StringToFix)
    ReturnString = ""
    
    For iCurrChar = 1 To StrLen
        CurrChar = Mid(StringToFix, iCurrChar, 1)
        Select Case CurrChar
            Case "'"
                Select Case True
                    Case bQuotesOnDouble
                    Case Else
                        bQuotesOnSingle = Not bQuotesOnSingle
                End Select
            Case Chr(34)
                Select Case True
                    Case bQuotesOnSingle
                    Case Else
                        bQuotesOnDouble = bQuotesOnDouble
                End Select
        End Select
        ReturnString = ReturnString & IIf(bQuotesOnDouble Or bQuotesOnSingle, CurrChar, UCase(CurrChar))
    Next iCurrChar
    
    SQLUpCase = ReturnString
    
End Function


Public Function SQLQuotes(ByVal sql As String) As String

    Dim i As Integer
    Dim bQuotesOn As Boolean
    Dim iFirstQuote As Integer
    Dim iSQLLen As Integer
    
    If Left$(sql, 7) = "INSERT " Then
        'Don't do it
        SQLQuotes = sql
    Else
        bQuotesOn = False
        sql = " " & sql & " "
        iSQLLen = Len(sql)
        iFirstQuote = InStr(sql, "'")
        If iFirstQuote = 0 Then
            iFirstQuote = iSQLLen
        End If
    
        For i = iSQLLen To iFirstQuote Step -1
            If InStr("'" & Chr(34), Mid(sql, i, 1)) > 0 Then
    
                If bQuotesOn Then
                    If InStr(" <>=(", Mid(sql, i - 1, 1)) > 0 Then
                        bQuotesOn = False
                        Mid(sql, i, 1) = """"
                    Else
                        Mid(sql, i, 1) = "'"
                    End If
                Else
                    bQuotesOn = True
                    Mid(sql, i, 1) = """"
                End If
            End If
        Next i
       
        SQLQuotes = Mid(sql, 2, Len(sql) - 2)
    End If
    
End Function

''''Public Function StrIsSame(String1 As Variant, String2 As Variant) As Boolean
''''    StrIsSame = Trim(UCase(CStr(String1))) = Trim(UCase(CStr(String2)))
''''End Function





'''''Public Function DateMonthAdd(ByVal Date2AddMonthTo As Date, ByVal Months2Add As Integer) As Date
'''''
'''''    Dim CurrYear As Integer
'''''    Dim CurrMonth As Integer
'''''    Dim CurrDay As Integer
'''''    Dim CurrTime As String
'''''
'''''    CurrYear = Year(Date2AddMonthTo)
'''''    CurrDay = Day(Date2AddMonthTo)
'''''    CurrMonth = Month(Date2AddMonthTo)
'''''    CurrTime = Mid(CStr(Date2AddMonthTo), 11)
'''''
'''''    CurrMonth = CurrMonth + Months2Add
'''''
'''''    While CurrMonth > 12
'''''        CurrMonth = CurrMonth - 12
'''''        CurrYear = CurrYear + 1
'''''    Wend
'''''
'''''    While CurrMonth < 1
'''''        CurrMonth = CurrMonth + 12
'''''        CurrYear = CurrYear - 1
'''''    Wend
'''''
'''''    DateMonthAdd = CDate(Format(CurrMonth, "00") & "/" & Format(CurrDay, "00") & "/" & Format(CurrYear, "0000") & CurrTime)
'''''
'''''End Function


'''''Public Function DateYearAdd(ByVal Date2AddYearTo As Date, ByVal Years2Add As Integer) As Date
'''''
'''''    Dim CurrYear As Integer
'''''    Dim CurrMonth As Integer
'''''    Dim CurrDay As Integer
'''''    Dim CurrTime As String
'''''
'''''    CurrYear = Year(Date2AddYearTo)
'''''    CurrDay = Day(Date2AddYearTo)
'''''    CurrMonth = Month(Date2AddYearTo)
'''''    CurrTime = Mid(CStr(Date2AddYearTo), 11)
'''''
'''''    CurrYear = CurrYear + Years2Add
'''''
'''''    DateYearAdd = CDate(Format(CurrMonth, "00") & "/" & Format(CurrDay, "00") & "/" & Format(CurrYear, "0000") & CurrTime)
'''''
'''''End Function

Public Function CenterFormInScreen(ByRef FormToCenter As Form)
    AutoCenterToScreen FormToCenter
End Function

Public Function CenterObjInForm(ByRef ObjectToCenter As Object, ByRef FormToCenterIn As Object, Optional ByVal HorVerOption As AlignmentCenteringConstants)
    '  HorVerOption = 1  -- Center Horizontally ONLY
    '  HorVerOption = 2  -- Center Vertically ONLY
    '  HorVerOption = 3  -- Center Both
    If IsMissing(HorVerOption) Then
        HorVerOption = 0
    End If
    If HorVerOption = 0 Then
        HorVerOption = Center
    End If
    
    Dim iNewLeft, iNewTop As Integer
    iNewLeft = ObjectToCenter.Left
    iNewTop = ObjectToCenter.Top
    If HorVerOption And CenterHorizontally Then
        iNewLeft = (FormToCenterIn.ScaleWidth / 2) - (ObjectToCenter.Width / 2)
    End If
    If HorVerOption And CenterVertically Then
        iNewTop = (FormToCenterIn.ScaleHeight / 2) - (ObjectToCenter.Height / 2)
    End If
    ObjectToCenter.Move iNewLeft, iNewTop
End Function

'''''Public Function CenterChildFormInParent(ByRef ChildFormToCenter As Form, ByRef ParentFormToCenterIn As MDIForm)
'''''    ChildFormToCenter.Move (ParentFormToCenterIn.ScaleWidth / 2) - (ChildFormToCenter.Width / 2), (ParentFormToCenterIn.ScaleHeight / 2) - (ChildFormToCenter.Height / 2)
'''''End Function

Public Function CenterChildObjInParent(ByRef ChildObjectToCenter As Object, ByRef ParentObjectToCenterIn As Object, Optional ByVal CenteringOptions As AlignmentCenteringConstants)
    If CenteringOptions = 0 Then
        CenteringOptions = Center
    End If
    Dim iNewLeft, iNewTop As Integer
    iNewLeft = ChildObjectToCenter.Left
    iNewTop = ChildObjectToCenter.Top
    If CenteringOptions And CenterHorizontally Then
        iNewLeft = (ParentObjectToCenterIn.Width / 2) - (ChildObjectToCenter.Width / 2)
    End If
    If CenteringOptions And CenterVertically Then
        iNewTop = (ParentObjectToCenterIn.Height / 2) - (ChildObjectToCenter.Height / 2)
    End If
    ChildObjectToCenter.Move iNewLeft, iNewTop
End Function


Public Function CenterObjInObj(ByRef ObjectToCenter As Object, ByRef ObjectToCenterIn As Object, Optional ByVal iHorizontalVertical As AlignmentCenteringConstants)
    If IsMissing(iHorizontalVertical) Then
        iHorizontalVertical = Center
    End If
    If iHorizontalVertical = 0 Then
        iHorizontalVertical = Center
    End If
    Dim iNewLeft, iNewTop As Integer
    iNewLeft = ObjectToCenter.Left
    iNewTop = ObjectToCenter.Top
    If iHorizontalVertical And CenterHorizontally Then
        On Error Resume Next
        iNewLeft = ObjectToCenterIn.Left + ((ObjectToCenterIn.Width / 2) - (ObjectToCenter.Width / 2))
        On Error GoTo 0
    End If
    If iHorizontalVertical And CenterVertically Then
        On Error Resume Next
        iNewTop = ObjectToCenterIn.Top + ((ObjectToCenterIn.Height / 2) - (ObjectToCenter.Height / 2))
        On Error GoTo 0
    End If
    ObjectToCenter.Move iNewLeft, iNewTop
End Function

Public Function CenterObjectInObject(ByRef ObjectToCenter As Object, ByRef ObjectToCenterIn As Object, Optional ByVal iHorizontalVertical As AlignmentCenteringConstants)
    
    Dim iNewLeft As Integer
    Dim iNewTop As Integer
    
    ' Default To Center
    If IsMissing(iHorizontalVertical) Then
        iHorizontalVertical = 3
    End If
    If iHorizontalVertical = 0 Then
        iHorizontalVertical = 3
    End If
    
    On Error Resume Next
    
    iNewLeft = ObjectToCenter.Left
    iNewTop = ObjectToCenter.Top
    
    If iHorizontalVertical And vbHorizontal Then
        iNewLeft = ObjectToCenterIn.Left + ((ObjectToCenterIn.Width / 2) - (ObjectToCenter.Width / 2))
    End If
    If iHorizontalVertical And vbVertical Then
        iNewTop = ObjectToCenterIn.Top + ((ObjectToCenterIn.Height / 2) - (ObjectToCenter.Height / 2))
    End If
    
    ObjectToCenter.Move iNewLeft, iNewTop
        
    On Error GoTo 0
    
End Function





Public Function Century(ByRef Date2GetCenturyFrom As Date) As Integer
    Century = Val(Left(Format(Date2GetCenturyFrom, "yyyy"), 2))
End Function

Public Function Y2Kdate(ByVal Date2Fix As Date) As Date
    
    Dim wkMonth As Integer
    Dim wkDay As Integer
    Dim wkYear As Integer
    Dim wkCentury As Integer
    Dim wkThreshold As Integer
    wkCentury = Century(Date2Fix)
    wkYear = Val(Right(Format(Date2Fix, "yyyy"), 2))
    If Tools.Y2Kthreshold = 0 Then
        wkThreshold = 40
    Else
        wkThreshold = Tools.Y2Kthreshold
    End If
    Select Case True
        Case wkYear > wkThreshold And wkCentury = 19
        Case wkYear > wkThreshold And wkCentury = 20
            Date2Fix = DateAdd("yyyy", -100, Date2Fix)
        Case wkYear <= wkThreshold And wkCentury = 20
        Case wkYear <= wkThreshold And wkCentury = 19
            Date2Fix = DateAdd("yyyy", 100, Date2Fix)
    End Select
    Y2Kdate = Date2Fix
    
End Function

Public Function Y2Kstring(ByVal Date2Fix As String, Optional ByVal DateFormat As String) As String
    If Len(DateFormat) = 0 Then
        DateFormat = "mm/dd/yyyy"
    End If
    Y2Kstring = Format(Y2Kdate(CDate(Date2Fix)), DateFormat)
End Function


'''Public Function RInstr(ByVal String2Search As String, ByVal String2Look4 As String) As Long
'''    Debug.Print vbCrLf & vbCrLf & "!!!>>> Defunct Function:  Replace with InStrRev() <<<!!!" & vbCrLf & vbCrLf
'''    RInstr = InStrRev(String2Search, String2Look4)
''''''    Dim lLoc As Long
''''''    Dim lMatchLen As Long
''''''    lMatchLen = Len(String2Look4)
''''''    Dim lStringLen As Long
''''''    lStringLen = Len(String2Search)
''''''    For lLoc = lStringLen - lMatchLen To 1 Step -1
''''''        If Mid(String2Search, lLoc, lMatchLen) = String2Look4 Then
''''''            RInstr = lLoc
''''''            Exit Function
''''''        End If
''''''    Next
''''''    RInstr = 0
''''    Dim iLocation As Integer
''''    iLocation = InStr(1, InvertString(String2Search), InvertString(String2Look4))
''''    If iLocation = 0 Then
''''        RInstr = 0
''''    Else
''''        RInstr = (Len(String2Search) + 1) - iLocation
''''    End If
'''End Function


Public Function InvertString(ByVal String2Invert As String) As String
    Dim i As Long
    Dim wkLen As Long
    Dim OutPutString As String
    OutPutString = ""
    wkLen = Len(String2Invert)
    For i = wkLen To 1 Step -1
        OutPutString = OutPutString & Mid(String2Invert, i, 1)
    Next i
    InvertString = OutPutString
End Function

Public Function arrayCount(ByRef WorkArray As Variant) As Integer
    Dim ReturnVal As Integer
    On Error Resume Next
    ReturnVal = UBound(WorkArray)
    On Error GoTo 0
    arrayCount = ReturnVal
End Function

Public Function arrayAddTo(ByRef WorkArray As Variant, ByVal NewValue2Add As Variant) As Integer
    Dim iArrCnt As Integer
    iArrCnt = arrayCount(WorkArray)
    iArrCnt = iArrCnt + 1
    ReDim Preserve WorkArray(iArrCnt)
    WorkArray(iArrCnt) = NewValue2Add
    arrayAddTo = iArrCnt
End Function

Public Function arrayTakeFrom(ByRef WorkArray As Variant, Optional ByVal Value2Remove As Variant) As Integer
    Dim iArrCnt As Integer
    iArrCnt = arrayCount(WorkArray)
    If Len(Value2Remove) <> 0 Then
        Dim i As Integer
        For i = LBound(WorkArray) To iArrCnt
            If WorkArray(i) = Value2Remove Then
                Dim t As Integer
                For t = i To iArrCnt - 1
                    WorkArray(t) = WorkArray(t + 1)
                Next t
                Exit For
            End If
        Next i
    End If
    iArrCnt = iArrCnt - 1
    ReDim Preserve WorkArray(iArrCnt)
    arrayTakeFrom = iArrCnt
End Function


''''Public Function arrayRippleSort(ByRef ArrayToSort As Variant)
''''
''''  Dim NumOfEntries As Integer, NumOfTimes As Integer
''''  Dim i As Integer, j As Integer
''''  Dim temp As Variant
''''
''''
''''  NumOfEntries = UBound(ArrayToSort)
''''  NumOfTimes = NumOfEntries - 1
''''
''''  For i = LBound(ArrayToSort) To NumOfTimes
''''    For j = i + 1 To NumOfEntries
''''      If ArrayToSort(j) < ArrayToSort(i) Then     ' swap array items
''''        temp = ArrayToSort(j)
''''        ArrayToSort(j) = ArrayToSort(i)
''''        ArrayToSort(i) = temp
''''      End If
''''    Next j
''''  Next i
''''
''''End Function

''''''Public Function arrayShellSort(a() As Variant)
''''''    Dim i As Integer, j As Integer
''''''    Dim Low As Integer, hi As Integer
''''''    Dim PushPop As Variant
''''''   Low = LBound(a)
''''''   hi = UBound(a)
''''''   j = (hi - Low + 1) \ 2
''''''   Do While j > 0
''''''     For i = Low To hi - j
''''''      If a(i) > a(i + j) Then
''''''         PushPop = a(i)
''''''         a(i) = a(i + j)
''''''         a(i + j) = PushPop
''''''      End If
''''''     Next i
''''''     For i = hi - j To Low Step -1
''''''      If a(i) > a(i + j) Then
''''''         PushPop = a(i)
''''''         a(i) = a(i + j)
''''''         a(i + j) = PushPop
''''''      End If
''''''     Next i
''''''     j = j \ 2
''''''   Loop
''''''
''''''End Function


'''''''Public Function arraySortNPack(ByRef WorkArray As Variant)
'''''''    arrayRippleSort WorkArray
'''''''    Dim iArrTop As Integer
'''''''    iArrTop = UBound(WorkArray)
'''''''    Dim i As Integer
'''''''    Dim CurrArrCell As Integer
'''''''    CurrArrCell = LBound(WorkArray)
'''''''    For i = LBound(WorkArray) + 1 To iArrTop
'''''''        If WorkArray(i) = WorkArray(CurrArrCell) Then
'''''''        Else
'''''''            CurrArrCell = CurrArrCell + 1
'''''''            WorkArray(CurrArrCell) = WorkArray(i)
'''''''        End If
'''''''    Next i
'''''''    ReDim Preserve WorkArray(CurrArrCell)
'''''''End Function



'''''''Public Function Wait(ByVal SecondsToWait As Single)
'''''''    Dim StartTime As Single
'''''''    StartTime = Timer
'''''''    Do While Abs(Timer - StartTime) < SecondsToWait
'''''''        DoEvents
'''''''    Loop
'''''''
'''''''End Function


Public Sub DisableAll(ByRef Form2DisableControlsOn As Form)
    
    On Error Resume Next
    Dim X As Integer
    For X = 0 To Form2DisableControlsOn.Controls.Count - 1
        Form2DisableControlsOn.Controls(X).Enabled = False
        If Err.Number <> 438 And Err.Number <> 0 Then
            Err.Raise Err.Number, "Tools.DisableAll", Err.Description
        End If
    Next X
    On Error GoTo 0
    
End Sub
''Public Sub DisableAllExceptMenus(ByRef Form2DisableControlsOn As Form)
''
''    On Error Resume Next
''    Dim X As Integer
''    For X = 0 To Form2DisableControlsOn.Controls.Count - 1
''        If Not ObjectIsMenu(Form2DisableControlsOn.Controls(X)) Then
''            Form2DisableControlsOn.Controls(X).Enabled = False
''            If Err.Number <> 438 And Err.Number <> 0 Then
''                Err.Raise Err.Number, "Tools.DisableAll", Err.Description
''            End If
''        End If
''    Next X
''    On Error GoTo 0
''
''End Sub

Public Sub EnableAll(ByRef Form2EnableControlsOn As Form)
    
    On Error Resume Next
    Dim X As Integer
    For X = 0 To Form2EnableControlsOn.Controls.Count - 1
        Form2EnableControlsOn.Controls(X).Enabled = True
        If Err.Number <> 438 And Err.Number <> 0 Then
            Err.Raise Err.Number, "Tools.EnableAll", Err.Description
        End If
    Next X
    On Error GoTo 0
    
End Sub
'''Public Sub EnableAllExceptMenus(ByRef Form2EnableControlsOn As Form)
'''
'''    On Error Resume Next
'''    Dim X As Integer
'''    For X = 0 To Form2EnableControlsOn.Controls.Count - 1
'''        If Not ObjectIsMenu(Form2EnableControlsOn.Controls(X)) Then
'''            Form2EnableControlsOn.Controls(X).Enabled = True
'''            If Err.Number <> 438 And Err.Number <> 0 Then
'''                Err.Raise Err.Number, "Tools.EnableAll", Err.Description
'''            End If
'''        End If
'''    Next X
'''    On Error GoTo 0
'''
'''End Sub

''''''Public Function PrintToFile(ByVal FileNameToWrite2 As String, ByVal String2WriteOut As String, Optional ByVal bOverWrite As Boolean) As Boolean
''''''
''''''    If bOverWrite Then
''''''        On Error Resume Next
''''''        Kill FileNameToWrite2
''''''        On Error GoTo 0
''''''    End If
''''''    If Right(String2WriteOut, 2) = vbCrLf Then
''''''        String2WriteOut = Left(String2WriteOut, Len(String2WriteOut) - 2)
''''''    End If
''''''    On Error GoTo didntwork
''''''    Dim FileHandle As Integer
''''''    FileHandle = FreeFile
''''''    Open FileNameToWrite2 For Append Access Write Lock Read As #FileHandle
''''''    Print #FileHandle, String2WriteOut
''''''    Close #FileHandle
''''''    PrintToFile = True
''''''    Exit Function
''''''
''''''didntwork:
''''''    PrintToFile = False
''''''
''''''End Function




''''''Public Function AppendCharacter(ByVal String2AppendTo As String, ByVal Char2Append As String) As String
''''''
''''''    If Right(String2AppendTo, Len(Char2Append)) <> Char2Append Then
''''''        AppendCharacter = String2AppendTo & Char2Append
''''''    Else
''''''        AppendCharacter = String2AppendTo
''''''    End If
''''''
''''''End Function



''''''Public Function fileIsReadOnly(ByRef FileNameToCheck As String) As Boolean
''''''    fileIsReadOnly = CBool(GetAttr(FileNameToCheck) And vbReadOnly)
''''''End Function
''''''
''''''
''''''Public Function fileIsHidden(ByRef FileNameToCheck As String) As Boolean
''''''    fileIsHidden = CBool(GetAttr(FileNameToCheck) And vbHidden)
''''''End Function

Public Function AlignToolBarButtonToObject(ByRef ObjectToAlign As Object, ByRef ObjectToAlignTo As Object, ByVal Alignment As AlignmentConstantsNEW)
    Dim iBtnCnt As Integer
    iBtnCnt = ObjectToAlign.Buttons.Count
        
    If Alignment And vbLeft Then     ' LEFT
        ObjectToAlign.Left = ObjectToAlignTo.Left - (ObjectToAlign.Buttons(iBtnCnt).Left)
    End If
    If Alignment And vbTop Then     ' TOP
        ObjectToAlign.Top = ObjectToAlignTo.Top - (ObjectToAlign.Buttons(iBtnCnt).Top)
    End If
    If Alignment And vbright Then     ' RIGHT
        ObjectToAlign.Left = ObjectToAlignTo.Left + ObjectToAlignTo.Width - ObjectToAlign.Buttons(iBtnCnt).Width
    End If
    If Alignment And vbRightFlush Then  ' RIGHTFLUSH
        ObjectToAlign.Left = ObjectToAlignTo.Left - (ObjectToAlign.Width)
    End If
    If Alignment And vbLeftFlush Then   ' LEFTFLUSH
        ObjectToAlign.Left = ObjectToAlignTo.Left + ObjectToAlignTo.Width
    End If
    
End Function


Public Function AlignObjectToToolBarButton(ByRef ObjectToAlign As Object, ByRef ObjectToAlignTo As Object, ByVal Alignment As AlignmentConstantsNEW)
    
    Dim iBtnCnt As Integer
    iBtnCnt = ObjectToAlignTo.Buttons.Count
    
'    If Alignment And vbLeft Then     ' LEFT
'        ObjectToAlign.Left = ObjectToAlignTo.Left - (ObjectToAlign.Buttons(iBtnCnt).Left)
'    End If
    If Alignment And vbTop Then     ' TOP
        ObjectToAlign.Top = ObjectToAlignTo.Top + ObjectToAlignTo.Buttons(iBtnCnt).Top
    End If
'    If Alignment And vbRight Then     ' RIGHT
'    End If
    If Alignment And vbRightFlush Then  ' RIGHTFLUSH
        ObjectToAlign.Left = (ObjectToAlignTo.Left + ObjectToAlignTo.Buttons(iBtnCnt).Left) - ObjectToAlign.Width
    End If
    If Alignment And vbLeftFlush Then   ' LEFTFLUSH
        ObjectToAlign.Left = ObjectToAlignTo.Left + ObjectToAlignTo.Width
    End If
    
    If Alignment And vbMiddleVertical Then
        ObjectToAlign.Top = ObjectToAlignTo.Top + (ObjectToAlignTo.Height / 2) - (ObjectToAlign.Height / 2)
    End If
    If Alignment And vbMiddleHorizontal Then
        ObjectToAlign.Left = ObjectToAlignTo.Left + (ObjectToAlignTo.Width / 2) - (ObjectToAlign.Width / 2)
    End If
    
End Function

Public Function AlignObjectToObject(ByRef ObjectToAlign As Object, ByRef ObjectToAlignTo As Object, ByVal Alignment As AlignmentConstantsNEW)
    
    If Alignment And vbLeft Then     ' LEFT
        ObjectToAlign.Left = ObjectToAlignTo.Left
    End If
    If Alignment And vbTop Then     ' TOP
        ObjectToAlign.Top = ObjectToAlignTo.Top
    End If
    If Alignment And vbright Then     ' RIGHT
        ObjectToAlign.Left = (ObjectToAlignTo.Left + ObjectToAlignTo.Width) - ObjectToAlign.Width
    End If
    If Alignment And vbRightFlush Then  ' RIGHTFLUSH
        ObjectToAlign.Left = (ObjectToAlignTo.Left + ObjectToAlignTo.Width)
    End If
    If Alignment And vbLeftFlush Then   ' LEFTFLUSH
        ObjectToAlign.Left = ObjectToAlignTo.Left + ObjectToAlignTo.Width
    End If
    
    If Alignment And vbMiddleVertical Then
        ObjectToAlign.Top = ObjectToAlignTo.Top + (ObjectToAlignTo.Height / 2) - (ObjectToAlign.Height / 2)
    End If
    If Alignment And vbMiddleHorizontal Then
        ObjectToAlign.Left = ObjectToAlignTo.Left + (ObjectToAlignTo.Width / 2) - (ObjectToAlign.Width / 2)
    End If
    
End Function
Public Function FnStripExtraTableNames(ByVal sSQL As String, ByVal sTableName As String) As String

Dim iSqlLenLoop  As Integer
Dim iWhereloop  As Integer
Dim iStartPos  As Integer
Dim sWhereClause As String
Dim sSelectClause As String
Dim bNotJunkString As Boolean

    
    

'This function removes the non-needed tablenames if the origional SQL Statement contains
'non-needed tablenames in the where part of the SQL statement
'For instance if the sql statment has 'SELECT * FROM APD_BASE WHERE APD_BASE.NUMBER_KEY = "B98"'
'It should really read Select * from APD_BASE WHERE NUMBER_KEY = "B98"

'First make sure that the tablename is alone
    
    bNotJunkString = True 'Assume that the string getting entered is not formatted correctly
    For iSqlLenLoop = 1 To Len(sSQL)
    'Look for the string that says "FROM {TABLE.NAME} WHERE
        If Trim$(UCase$(Mid$(sSQL, iSqlLenLoop, Len(sTableName) + 12))) = Trim$("FROM " & UCase$(sTableName) & " WHERE") Then
            iStartPos = iSqlLenLoop + 20 'found at this location
            sSelectClause = UCase$(Left$(sSQL, iStartPos - 1)) 'The Select clause = "Select * from tablename"
            sWhereClause = Mid$(sSQL, iStartPos)  'the where clause is where element_key = ....
            sTableName = UCase$(sTableName)
            sWhereClause = Replace(sWhereClause, sTableName & ".", "") 'replace {table_name.field} with {field}
            FnStripExtraTableNames = sSelectClause & sWhereClause 'and this is the result
            bNotJunkString = False 'the data isn't junk
            Exit For
        End If
    Next iSqlLenLoop
    If bNotJunkString = True Then FnStripExtraTableNames = sSQL  'if the data is junk just return the string sent


End Function

'''''Public Function TrimToolBar(ByRef ToolBarToTrim As Object)
'''''
'''''    ToolBarToTrim.ButtonWidth = 20
'''''    ToolBarToTrim.ButtonHeight = 20
'''''    Dim iMaxButton As Integer
'''''    iMaxButton = ToolBarToTrim.Buttons.Count
'''''    On Error Resume Next
'''''        ToolBarToTrim.Width = ToolBarToTrim.Buttons(iMaxButton).Left + ToolBarToTrim.Buttons(iMaxButton).Width
'''''    On Error GoTo 0
'''''
'''''    ToolBarToTrim.Height = ToolBarToTrim.Buttons(iMaxButton).Top + ToolBarToTrim.Buttons(iMaxButton).Height - ts.twipsY(1)
'''''
'''''End Function




'''''Public Function TimeToSeconds(ByVal sTime As String) As Long
'''''
'''''    Dim lOutSecs As Long
'''''    lOutSecs = 0
'''''
'''''    Dim iCurrColon
'''''    iCurrColon = InStr(sTime, ":")
'''''
'''''    If iCurrColon > 0 Then
'''''        lOutSecs = (val(sTime) * 60) * 60
'''''        sTime = Mid(sTime, iCurrColon + 1)
'''''        iCurrColon = InStr(sTime, ":")
'''''        If iCurrColon > 0 Then
'''''            lOutSecs = lOutSecs + (val(sTime) * 60)
'''''            sTime = Mid(sTime, iCurrColon + 1)
'''''            lOutSecs = lOutSecs + val(sTime)
'''''
'''''        Else
'''''            lOutSecs = lOutSecs + (val(sTime) * 60)
'''''        End If
'''''
'''''    Else
'''''        lOutSecs = (val(sTime) * 60) * 60
'''''    End If
'''''
'''''    TimeToSeconds = lOutSecs
'''''
'''''End Function


'''''''Public Function SecondsToTime(ByVal lSeconds As Long, Optional ByVal sFormat As String) As String
'''''''
'''''''    Dim sHour As String
'''''''    Dim sMinute As String
'''''''    Dim sSeconds As String
'''''''
'''''''    sHour = Format(Int(lSeconds / 60 / 60), "00")
'''''''    lSeconds = lSeconds - (val(sHour) * 60 * 60)
'''''''    sMinute = Format(Int(lSeconds / 60), "00")
'''''''    lSeconds = lSeconds - (val(sMinute) * 60)
'''''''    sSeconds = Format(lSeconds, "00")
'''''''
'''''''    Select Case True
'''''''        Case sOccurs(sFormat, ":") = 2
'''''''            SecondsToTime = sHour & ":" & sMinute & ":" & sSeconds
'''''''        Case UCase$(sFormat) = "MM:SS"
'''''''            SecondsToTime = sMinute & ":" & sSeconds
'''''''        Case Else
'''''''            SecondsToTime = sHour & ":" & sMinute
'''''''    End Select
'''''''
'''''''End Function

Public Function HidePassword(ByVal sString As String) As String
    HidePassword = FNHidePassword(sString)
End Function
Public Function FNHidePassword(ByVal sString As String) As String
    
    Dim sTempVar
    Dim iPos As Integer
    Dim zed As Integer
    Dim iEndPos As Integer
    Dim iBegPos As Integer
     
    
        
'This rather boring function hides the password from an ODBC string. It
'will probably only be used in the "WHO AM I" FORM

    sTempVar = Trim$(UCase$(sString))
    iBegPos = InStr(sTempVar, ";PWD=")
    
    If iBegPos <> 0 Then
    
        For zed = iBegPos + 1 To Len(sTempVar)
        
            If Mid$(sTempVar, zed, 1) = ";" Then
                iEndPos = zed
                Exit For
            End If
            
        Next zed
    
        If iEndPos <> 0 Then
            FNHidePassword = Left$(sTempVar, iBegPos + 4) & "********" & Mid$(sTempVar, iEndPos)
        Else
            FNHidePassword = Left$(sTempVar, iBegPos + 4) & "********;"
        End If
    
    Else
        iBegPos = InStr(sTempVar, ";pwd=")
    
        If iBegPos <> 0 Then
    
            For zed = iBegPos + 1 To Len(sTempVar)
    
                If Mid$(sTempVar, zed, 1) = ";" Then
                    iEndPos = zed
                    Exit For
                End If
    
            Next zed
        
        Else
    
            FNHidePassword = sString
        End If
    
    End If
    
End Function


''''Public Function Set_Focus(ByRef ObjToSetFocusTo As Object)
''''    On Error Resume Next
''''    If ObjToSetFocusTo.Enabled Then
''''        ObjToSetFocusTo.SetFocus
''''    End If
''''    On Error GoTo 0
''''End Function


'''''Public Function TrimAll(ByVal sString2Trim As String, Optional ByVal sOptionalSpaceCharacter As String) As String
'''''
'''''    If Trim(sOptionalSpaceCharacter & "") = "" Then
'''''        sOptionalSpaceCharacter = " "
'''''    End If
'''''
'''''    Do While Left(sString2Trim, Len(sOptionalSpaceCharacter)) = sOptionalSpaceCharacter
'''''        sString2Trim = Mid(sString2Trim, Len(sOptionalSpaceCharacter) + 1)
'''''    Loop
'''''
'''''    Do While Right(sString2Trim, Len(sOptionalSpaceCharacter)) = sOptionalSpaceCharacter
'''''        sString2Trim = Left(sString2Trim, Len(sString2Trim) - Len(sOptionalSpaceCharacter))
'''''    Loop
'''''
'''''    TrimAll = sString2Trim
'''''
'''''End Function



'''''Public Function RenameFile(ByVal sFrom As String, ByVal sTO As String) As Boolean
'''''
'''''    On Error Resume Next
'''''    Name sFrom As sTO
'''''    RenameFile = (Err.Number = 0)
'''''    On Error GoTo 0
'''''
'''''End Function


''''''Public Function FileWithoutExtension(ByVal sFileNameAndPath As String) As String
''''''
''''''    Dim i As Integer
''''''    Dim sReturn As String
''''''    sReturn = sFileNameAndPath
''''''
''''''    For i = Len(sFileNameAndPath) To 1 Step -1
''''''        If Mid(sFileNameAndPath, i, 1) = "." Then
''''''            sReturn = Left(sFileNameAndPath, i - 1)
''''''            Exit For
''''''        End If
''''''    Next i
''''''
''''''    FileWithoutExtension = sReturn
''''''
''''''End Function


Public Function arrayIsSet(Optional ByRef aWork As Variant) As Boolean
    Err.Clear
    Dim iTest As Integer
    On Error Resume Next
    iTest = UBound(aWork)
    arrayIsSet = (Err.Number = 0)
    On Error GoTo 0
End Function





'''Public Function GetNewFileName(ByVal sStartingPathName As String)
'''
'''    Dim sFilePath, sFileName, sFileExt As String
'''    sFilePath = Tools.GetFilePath(sStartingPathName)
'''    sFileName = Tools.GetFileName(sStartingPathName)
'''    sFileExt = Tools.GetFileExt(sStartingPathName)
'''
'''    If Not ts.fileExists(sStartingPathName) Then
'''        GetNewFileName = sStartingPathName
'''        Exit Function
'''    End If
'''
'''    Dim iNewNumber As Integer
'''    iNewNumber = 2
'''
'''    Do While ts.fileExists(sFilePath & sFileName & " (" & Trim(Str(iNewNumber)) & ")." & sFileExt)
'''        iNewNumber = iNewNumber + 1
'''    Loop
'''
'''    GetNewFileName = sFilePath & sFileName & " (" & Trim(Str(iNewNumber)) & ")." & sFileExt
'''
'''End Function


''''Public Function SetCurDir(ByVal sPathName As String) As Boolean
''''    On Error Resume Next
''''    ChDrive sPathName
''''    ChDir sPathName
''''    SetCurDir = Err.Number = 0
''''    On Error GoTo 0
''''End Function






' 99-08-07
'=-=-=-=-=-

'''''Public Function DeleteDirectory(ByVal sDirectory As String, Optional ByRef StatusObject As Object) As Boolean
'''''
'''''    Dim sFileName As String
'''''    Dim Files As Collection
'''''    Dim i As Integer
'''''
'''''    On Error GoTo DeleteError
'''''
'''''    ' Get a list of files it contains.
'''''    Set Files = New Collection
'''''    sFileName = Dir$(sDirectory & "\*.*", vbReadOnly + vbHidden + vbSystem + vbDirectory)
'''''    Do While Len(sFileName) > 0
'''''        If (sFileName <> "..") And (sFileName <> ".") Then
'''''            Files.Add sDirectory & "\" & sFileName
'''''        End If
'''''        sFileName = Dir$()
'''''    Loop
'''''
'''''    ' Delete the files.
'''''    For i = 1 To Files.Count
'''''
'''''        DoEvents
'''''
'''''        sFileName = Files(i)
'''''        ' See if it is a directory.
'''''        If GetAttr(sFileName) And vbDirectory Then
'''''            ' It is a directory. Delete it.
'''''            DeleteDirectory sFileName
'''''        Else
'''''            ' It's a file. Delete it.
'''''            Select Case True
'''''                Case Tools.ObjectHas(StatusObject, Caption)
'''''                    StatusObject.Caption = sFileName
'''''                    StatusObject.Refresh
'''''
'''''            End Select
'''''            SetAttr sFileName, vbNormal
'''''            Kill sFileName
'''''        End If
'''''    Next i
'''''
'''''    ' The directory is now empty. Delete it.
'''''    Select Case True
'''''        Case Tools.ObjectHas(StatusObject, Caption)
'''''            StatusObject.Caption = sDirectory
'''''            DoEvents
'''''    End Select
'''''    RmDir sDirectory
'''''
'''''    DeleteDirectory = True
'''''    Set Files = Nothing
'''''    Exit Function
'''''
'''''DeleteError:
'''''
'''''    DeleteDirectory = False
'''''    On Error Resume Next
'''''    Set Files = Nothing
'''''    On Error GoTo 0
'''''
'''''End Function


Public Function ObjectHas(ByRef ObjectToCheck As Object, ByVal PropertyToCheckFor As PotentialControlProperties) As Boolean
    ObjectHas = Tools.ItemHasProperty(CVar(ObjectToCheck), PropertyToCheckFor)
End Function


'''''Public Sub GradientColorBackground(ByRef frmWork As Form, ByVal FromColor As Long, ByVal ToColor As Long)
'''''
'''''    If Not Tools.bDoingGradientBackGroundFill And frmWork.ScaleHeight > 0 And frmWork.WindowState <> vbMinimized Then
'''''        Tools.bDoingGradientBackGroundFill = True
'''''
'''''        Dim intLoop As Integer
'''''        Dim bSaveAutoRedraw As Boolean
'''''        Dim iDrawMode As Integer
'''''        Dim iDrawStyle As Integer
'''''        Dim lDrawWidth As Long
'''''        Dim lScaleHeight As Long
'''''        Dim iScaleMode As Integer
'''''
'''''        With frmWork
'''''            bSaveAutoRedraw = .AutoRedraw
'''''            iDrawStyle = .DrawStyle
'''''            iDrawMode = .DrawMode
'''''            iScaleMode = .ScaleMode
'''''            lDrawWidth = .DrawWidth
'''''            lScaleHeight = .ScaleHeight
'''''
'''''            .AutoRedraw = True
'''''            .DrawStyle = vbInsideSolid
'''''            .DrawMode = vbCopyPen
'''''            .ScaleMode = vbPixels
'''''            .DrawWidth = 2
'''''            .ScaleHeight = 256
'''''        End With
'''''
'''''        Dim iFromRed, iFromBlue, iFromGreen As Integer
'''''        Dim iToRed, iToBlue, iToGreen As Integer
'''''        Dim iCurrRed, iCurrBlue, iCurrGreen As Integer
'''''        iFromRed = GetRGBValue(FromColor, RGBRed)
'''''        iFromBlue = GetRGBValue(FromColor, RGBBlue)
'''''        iFromGreen = GetRGBValue(FromColor, RGBGreen)
'''''        iToRed = GetRGBValue(ToColor, RGBRed)
'''''        iToBlue = GetRGBValue(ToColor, RGBBlue)
'''''        iToGreen = GetRGBValue(ToColor, RGBGreen)
'''''
'''''        For intLoop = 0 To 255
'''''            iCurrRed = iFromRed - ((iFromRed - iToRed) * Fraction(intLoop, 255))
'''''            iCurrGreen = iFromGreen - ((iFromGreen - iToGreen) * Fraction(intLoop, 255))
'''''            iCurrBlue = iFromBlue - ((iFromBlue - iToBlue) * Fraction(intLoop, 255))
''''''            On Error Resume Next
'''''            frmWork.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(iCurrRed, iCurrGreen, iCurrBlue), BF
''''''            On Error GoTo 0
'''''            'DoEvents
'''''        Next
'''''
'''''        frmWork.Refresh
'''''        'DoEvents
'''''
'''''        With frmWork
'''''            .AutoRedraw = bSaveAutoRedraw
'''''            .DrawStyle = iDrawStyle
'''''            .DrawMode = iDrawMode
'''''            .DrawWidth = lDrawWidth
'''''            '.ScaleHeight = lScaleHeight
'''''            .ScaleMode = iScaleMode
'''''        End With
'''''
'''''        Tools.bDoingGradientBackGroundFill = False
'''''    End If
'''''
'''''End Sub


Public Function NumberToBits(ByVal lNumberToConvert As Long, Optional ByVal iNumberOfBits As Integer) As String
    
    Dim d As Double
    d = 1
    Dim sReturn As String
    
    Do While d < lNumberToConvert
        sReturn = IIf(d And lNumberToConvert, "1", "0") & sReturn
        d = d * 2
    Loop
    
    If lNumberToConvert = 1 Then
        sReturn = "1"
    End If
    
    If iNumberOfBits = 0 Then
        NumberToBits = sReturn
    Else
        NumberToBits = IIf(Len(sReturn) > iNumberOfBits, Right(sReturn, iNumberOfBits), String(iNumberOfBits - Len(sReturn), "0") & sReturn)
    End If
    
End Function


'''''Public Function BitsToNumber(ByVal sBitString As String) As Long
'''''
'''''    Dim lReturn As Long
'''''    Dim i As Integer
'''''    Dim iStrLen As Integer
'''''    iStrLen = Len(sBitString)
'''''    For i = iStrLen To 1 Step -1
'''''        If Mid(sBitString, iStrLen - i + 1, 1) = "1" Then
'''''            lReturn = lReturn + BitValue(i)
'''''        End If
'''''    Next i
'''''    BitsToNumber = lReturn
'''''End Function

'''Public Function BitValue(ByVal iBitPosition As Integer) As Long
'''    Dim lReturn As Long
'''    lReturn = 1
'''    Dim i As Integer
'''    For i = 2 To iBitPosition
'''        lReturn = lReturn * 2
'''    Next i
'''    BitValue = lReturn
'''End Function

'''''Public Function GetRGBValue(ByVal lColorToGetFrom As Long, ByVal iValueToGet As RGBBaseColors)
'''''
'''''    Dim sBits As String
'''''    sBits = NumberToBits(lColorToGetFrom, 24)
'''''    Select Case iValueToGet
'''''        Case RGBBlue
'''''            GetRGBValue = BitsToNumber(Left(sBits, 8))
'''''        Case RGBGreen
'''''            GetRGBValue = BitsToNumber(Mid(sBits, 9, 8))
'''''        Case RGBRed
'''''            GetRGBValue = BitsToNumber(Right(sBits, 8))
'''''    End Select
'''''End Function

'''''Public Function Fraction(ByRef ThisNumber As Variant, ByRef OfThisNumber As Variant) As Double
'''''    Fraction = ThisNumber / OfThisNumber
'''''End Function

''''Public Function Percent(ByRef ThisNumber As Long, ByRef OfThisNumber As Long) As Integer
''''    Percent = (ThisNumber / OfThisNumber) * 100
''''End Function

''''Public Function SelectAllTextInControl(ByRef TextControl As Object)
''''
''''    On Error Resume Next
''''    TextControl.SelStart = 0
''''    TextControl.SelLength = Len(TextControl.Text)
''''    On Error GoTo 0
''''End Function


''''''Public Function KeyAsciiAllow(ByVal KeyAscii As Integer, ByVal TypeToAllow As enumKeyPressAllowTypes) As Integer
''''''    KeyAsciiAllow = KeyAscii
''''''    ' NumbersOnly
''''''    If (TypeToAllow And NumbersOnly) > 0 Then
''''''        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
''''''            KeyAscii = 8 Or _
''''''            KeyAscii = 27 Or _
''''''            (AllowNegative And KeyAscii) Or _
''''''            (AllowDecimal And KeyAscii) Then
''''''            KeyAsciiAllow = KeyAscii
''''''        Else
''''''            KeyAsciiAllow = 0
''''''        End If
''''''    End If
''''''
''''''    ' UpperCase
''''''    If (TypeToAllow And Uppercase) > 0 Then
''''''        KeyAsciiAllow = Asc(UCase(Chr(KeyAscii)))
''''''    End If
''''''
''''''    ' No Spaces
''''''    If (TypeToAllow And NoSpaces) > 0 And KeyAscii = 32 Then
''''''        KeyAsciiAllow = 0
''''''    End If
''''''
''''''    ' No Single Quotes
''''''    If (TypeToAllow And NoSingleQuotes) > 0 And KeyAscii = Asc("'") Then
''''''        KeyAsciiAllow = 0
''''''    End If
''''''
''''''    ' No Double Quotes
''''''    If (TypeToAllow And NoDoubleQuotes) > 0 And KeyAscii = Asc("""") Then
''''''        KeyAsciiAllow = 0
''''''    End If
''''''End Function




''Public Function fnDecimalFix(dblValue As Double, intDecimals As Integer) As Double
''    Dim sWrk As String
''    sWrk = Format(dblValue, "0." & String$(intDecimals + 1, "0"))
''    If intDecimals = 0 Then
''        fnDecimalFix = CDbl(Left$(sWrk, Len(sWrk) - 2))
''    Else
''        fnDecimalFix = CDbl(Left$(sWrk, Len(sWrk) - 1))
''    End If
''End Function


''''''Public Function ObjectIsMenu(ByRef ObjectToCheck As Object) As Boolean
''''''    Dim Test As Variant
''''''    On Error Resume Next
''''''    Err.Clear
''''''    Test = ObjectToCheck.WindowList
''''''    ObjectIsMenu = (Err.Number = 0)
''''''    On Error GoTo 0
''''''
''''''End Function




''''''Public Function FileWithOutPath(ByVal sFullFileName As String) As String
''''''    Dim i As Integer
''''''    Dim sReturn As String
''''''    sReturn = sFullFileName
''''''    For i = Len(sFullFileName) To 1 Step -1
''''''        If Mid(sFullFileName, i, 1) = "\" Then
''''''            sReturn = Mid(sFullFileName, i + 1)
''''''            Exit For
''''''        End If
''''''    Next i
''''''    FileWithOutPath = sReturn
''''''End Function



Public Function FormHide(ByRef FormToHide As Form)
    FormToHide.Tag = FormToHide.Tag & "HIDEINFO~" & FormToHide.Left & "~" & FormToHide.Top
    FormToHide.Enabled = False
    FormToHide.Move -(FormToHide.Width) - 10, -(FormToHide.Height) - 10
    
End Function

Public Function FormUnhide(ByRef FormToUnhide As Form)
    Dim iNewLeft, iNewTop As Integer
    Dim sTag As String
    sTag = FormToUnhide.Tag
    
    iNewLeft = Val(sWord(sTag, sWordCnt(sTag, "~") - 1, "~"))
    iNewTop = Val(sWord(sTag, sWordCnt(sTag, "~"), "~"))
    
    FormToUnhide.Tag = Left(sTag, InStr(sTag, "HIDEINFO~") - 1)
    FormToUnhide.Enabled = True
    FormToUnhide.Move iNewLeft, iNewTop
        
End Function




Public Function ObjectIndexFor(ByRef WorkObject As Object, ByVal WorkString As String) As Integer
    ' Made to return the index value for a given string for a Combo Box.
    ' If no matches, returns the current ListIndex value.
    
    ObjectIndexFor = WorkObject.ListIndex
    Dim i As Integer
    For i = 0 To WorkObject.ListCount - 1
        If UCase(WorkObject.List(i)) = UCase(WorkString) Then
            ObjectIndexFor = i
            Exit Function
        End If
    Next i
    
End Function



Public Function Do_Events(Optional ByVal NumberOfTimes As Integer)
    If NumberOfTimes = 0 Then
        NumberOfTimes = 1
    End If
    Do While NumberOfTimes > 0
        If GetInputState() <> 0 Then
            DoEvents
        End If
        NumberOfTimes = NumberOfTimes - 1
    Loop
End Function





'''''Public Function formByHWnd(ByVal hwnd As Long) As Form
'''''    Dim fCurr As Form
'''''    For Each fCurr In Forms
'''''        If fCurr.hwnd = hwnd Then
'''''            Set formByHWnd = fCurr
'''''            Exit Function
'''''        End If
'''''    Next fCurr
'''''End Function



'''''Sub subTextBoxEdit(pCtrl As Control, Pmode As Integer, pIndex As Integer, pASCii As Integer, pShift As Integer)
'''''    Dim X As Integer
'''''    Dim w As Variant
'''''    Dim stInvalidKeys As String
'''''    stInvalidKeys = Chr$(34) + Chr$(39)
'''''    Select Case Pmode
'''''        Case 0 ' Text Box Got Focus
'''''            If pIndex > -1 Then
'''''                pCtrl(pIndex).SelStart = 0
'''''                pCtrl(pIndex).SelLength = Len(pCtrl(pIndex).Text) + pASCii
'''''            Else
'''''                pCtrl.SelStart = 0
'''''                pCtrl.SelLength = Len(pCtrl.Text) + pASCii
'''''            End If
'''''        Case 1 ' Key Pressed Key stroke Edit
'''''            If InStr(1, stInvalidKeys, Chr$(pASCii)) > 0 Then
'''''                pASCii = 0
'''''            End If
'''''        Case 2 ' Date Box got Focus
'''''            If pIndex > -1 Then
'''''                pCtrl(pIndex).SelStart = 0
'''''                pCtrl(pIndex).SelLength = Len(pCtrl(pIndex).RawData) + 2
'''''            Else
'''''                pCtrl.SelStart = 0
'''''                pCtrl.SelLength = Len(pCtrl.RawData) + 2
'''''            End If
'''''        Case 3
'''''            If InStr(1, Chr$(8) & "0123456789", Chr$(pASCii)) = 0 Then
'''''                pASCii = 0
'''''            End If
'''''        Case 4
'''''            If InStr(1, Chr$(8) & "0123456789", Chr$(pASCii)) = 0 Then
'''''                pASCii = 0
'''''            End If
'''''        Case 5
'''''            If IsDate(pCtrl) Then
'''''                w = CVDate(pCtrl)
'''''                If pASCii = 187 Then ' + or =
'''''                    pCtrl = DateAdd("d", 1, w)
'''''                ElseIf pASCii = 189 Then ' the minus key -
'''''                    pCtrl = DateAdd("d", -1, w)
'''''                ElseIf pASCii = 84 Then ' the T key
'''''                    pCtrl = Now
'''''                End If
'''''            Else
'''''                If pASCii = 187 Or pASCii = 189 Or pASCii = 84 Then
'''''                    pCtrl = Now
'''''                End If
'''''            End If
'''''    End Select
'''''End Sub


Public Function FormTileBackGroundImage(ByRef FormToDoBackGroundOn As Form, ByRef ImageToUse As Variant, Optional ByVal RefreshForm As Boolean, Optional ByVal Do_Events As Boolean) As Boolean
    
    Dim Row As Integer
    Dim Column As Integer
    Dim CurrCell As Integer
    Dim bHasPicture As Boolean
    Dim bHasLoad As Boolean
    Dim tileHeight As Long
    Dim tileWidth As Long
    
    CurrCell = 0
    On Error Resume Next
    
    bHasPicture = (ImageToUse > 0)
    If Not bHasPicture Then
        Set ImageToUse = LoadPicture(ImageToUse)
        bHasLoad = (ImageToUse > 0)
    End If
    On Error GoTo 0
    Select Case True
        Case bHasPicture Or bHasLoad
            DoEvents
            
            tileHeight = FormToDoBackGroundOn.ScaleY(ImageToUse.Height, vbHimetric, vbTwips)
            tileWidth = FormToDoBackGroundOn.ScaleX(ImageToUse.Width, vbHimetric, vbTwips)
            
            For Row = 0 To FormToDoBackGroundOn.Height + tileHeight Step tileHeight
                For Column = 0 To FormToDoBackGroundOn.Width + tileWidth Step tileWidth
                    FormToDoBackGroundOn.PaintPicture ImageToUse, Column, Row
                    If Do_Events Then
                        DoEvents
                    End If
                Next Column
            Next Row
            If RefreshForm Then
                FormToDoBackGroundOn.Refresh
            End If
            FormTileBackGroundImage = True
        Case Else
            FormToDoBackGroundOn.Cls
            If RefreshForm Then
                FormToDoBackGroundOn.Refresh
            End If
            FormTileBackGroundImage = False
    End Select
    
End Function







'''''Public Function PCase(ByVal txt As String) As String
'''''
'''''    Dim need_cap As Boolean
'''''    Dim i As Integer
'''''    Dim ch As String
'''''
'''''    txt = LCase(txt)
'''''    need_cap = True
'''''    For i = 1 To Len(txt)
'''''        ch = Mid$(txt, i, 1)
'''''        If (ch >= "a" And ch <= "z") Or ch = "'" Then
'''''            If need_cap Then
'''''                Mid$(txt, i, 1) = UCase(ch)
'''''                need_cap = False
'''''            End If
'''''        Else
'''''            need_cap = True
'''''        End If
'''''    Next i
'''''    PCase = txt
'''''
'''''End Function

''''Public Function Proper(ByVal sText As String) As String
''''    Proper = PCase(sText)
''''End Function



Public Function bcd2dbl(ByVal StringIn As String, DecPos As Integer) As Double
    
    Dim hexstr As String
    Dim L As Integer
    Dim X As Integer
    Dim neg As String
    Dim wk As String
    
    StringIn = Trim$(StringIn)
    X = Len(StringIn)
    L = (X * 2) - 1
    If L < DecPos Then
        hexstr = "0"
        neg = "0"
    Else
        hexstr = ""
        
        While X > 0
            wk = Format$(Hex$(Asc(Mid$(StringIn, X, 1))), "00")
            If Len(wk) < 2 Then
                wk = "0" & wk
            End If
            hexstr = wk + hexstr
            X = X - 1
        Wend
        
        If Right$(hexstr, 1) = "D" Then
            neg = "-"
        Else
            neg = "0"
        End If
        
        hexstr = Left$(hexstr, L)
        
        If L > DecPos Then
            hexstr = Left$(hexstr, L - DecPos) & "." & Right$(hexstr, DecPos)
        Else
            hexstr = "." & hexstr
        End If
    End If
    
    bcd2dbl = Val(neg + hexstr)
    
End Function

Function dbl2bcd(Double_Value As Double, DecPos As Integer, Storage_Width As Integer) As String
    
    Dim wkstr As String
    Dim Mask  As String
    Dim HexOut As String
    Dim X As Integer
    Dim Y As Integer
    Dim n As Integer
    If Double_Value < 0 Then
        n = 1
        Double_Value = 0 - Double_Value
    Else
        n = 0
    End If
    
    Mask = Format$(Double_Value, "000000000000000000.000000000000000000")
    wkstr = Right$(Mask, 18)
    wkstr = Left$(wkstr, DecPos)
    wkstr = Left$(Mask, 18) + wkstr
    If n Then
        wkstr = wkstr & "D"
    Else
        wkstr = wkstr & "C"
    End If
    
    X = 0
    Y = Len(wkstr) - 1
    HexOut = ""
    While X < Storage_Width
        Mask = "&H" & Mid$(wkstr, Y, 2)
        HexOut = Chr$(Val(Mask)) + HexOut
        Y = Y - 2
        X = X + 1
    Wend
    dbl2bcd = HexOut
    
End Function



Public Function ValueIsInCollection(ByVal vValue2Find As Variant, ByRef col2Check As Variant, Optional ByRef lIndex As Long, Optional ByVal bOneBased As Boolean = False) As Boolean
    
    On Error GoTo ErrorInSearch
    ValueIsInCollection = False
    Dim i As Integer
    For i = IIf(bOneBased, 1, 0) To IIf(bOneBased, col2Check.Count, col2Check.Count - 1)
        If col2Check(i) = vValue2Find Then
            lIndex = i
            ValueIsInCollection = True
            Exit For
        End If
    Next i
    Exit Function
    
ErrorInSearch:
    ValueIsInCollection = False
    
End Function


Public Function ItemIsInCollectionSpin(ByVal vItem As Variant, ByRef col2Check As Variant) As Boolean
    
    Dim vTest As Variant
    ItemIsInCollectionSpin = False
    Dim i As Integer
    If col2Check.Count > 0 Then
        Dim bZeroBased As Boolean
        On Error Resume Next
        vTest = col2Check(0)
        bZeroBased = (Err.Number = 0)
        On Error GoTo 0
        For i = IIf(bZeroBased, 0, 1) To IIf(bZeroBased, col2Check.Count - 1, col2Check.Count)
            If col2Check(i) = vItem Then
                ItemIsInCollectionSpin = True
                Exit Function
            End If
        Next i
    End If
    
End Function


Public Function ItemIsInCollection(ByVal vItem As Variant, ByRef col2Check As Variant) As Boolean
    If col2Check.Count > 0 Then
        
        Dim vTest As Variant
        On Error Resume Next
'        vTest = col2Check(vItem)
        ItemIsInCollection = (Not col2Check(vItem) Is Nothing)
        ItemIsInCollection = (Err.Number = 0)
        On Error GoTo 0
    Else
        ItemIsInCollection = False
    End If
    
'
'
'
'
'
'    Dim vTest As Variant
'    On Error Resume Next
'    vTest = col2Check(vItem)
'    ItemIsInCollection = Err.Number = 0
'    On Error GoTo 0
'
End Function



Public Function ItemHasProperty(ByRef vItem As Object, ByVal Property As PotentialControlProperties) As Boolean
    
    Dim vTest As Variant
    On Error GoTo NotAProperty
    Select Case True
        Case Property = Justification
            vTest = vItem.Justification
        Case Property = MaxCharacters
            vTest = vItem.MaxCharacters
        Case Property = MaxLength
            vTest = vItem.MaxLength
        Case Property = Text
            vTest = vItem.Text
        Case Property = SelBookmarks
            vTest = vItem.SelBookmarks.Count
        Case Property = Connection
            vTest = vItem.Connection.Connect
        Case Property = LockEdits
            vTest = vItem.LockEdits
        Case Property = Caption
            vTest = vItem.Caption
        Case Property = Count
            vTest = vItem.Count
        Case Property = Connect
            vTest = vItem.Connect
        Case Property = Name
            vTest = vItem.Name
        Case Property = DateCreated
            vTest = vItem.DateCreated
        Case Property = LastModified
            vTest = vItem.LastModified
        Case Property = LastUpdated
            vTest = vItem.LastUpdated
        Case Property = Enabled
            vTest = vItem.Enabled
        Case Property = List
            vTest = vItem.List(0)
        Case Property = ListCount
            vTest = vItem.ListCount
        Case Property = ListIndex
            vTest = vItem.ListIndex
        Case Property = SimpleText
            vTest = vItem.SimpleText
        Case Property = Style
            vTest = vItem.Style
        Case Property = Panels
            vTest = vItem.Panels.Count
        Case Else
            GoTo NotAProperty
    End Select
    On Error GoTo 0
    ItemHasProperty = True
    Exit Function
    
NotAProperty:
    ItemHasProperty = False
    
End Function

Public Function GetProperty(ByRef vItem As Object, ByVal Property As PotentialControlProperties, Optional ByVal FormatString As String) As Variant
    On Error GoTo HadError
    Select Case True
        Case Property = Connection
            Set GetProperty = vItem.Connection
        Case Property = LockEdits
            If Tools.ItemHasProperty(vItem, LockEdits) Then
                GetProperty = CBool(vItem.LockEdits)
            Else
                GetProperty = False
            End If
        Case Property = DateCreated
            If Tools.ItemHasProperty(vItem, DateCreated) Then
                If IsDate(vItem.DateCreated) Then
                    GetProperty = PropFormat(vItem.DateCreated, FormatString)
                End If
            End If
        Case Property = LastModified
            If Tools.ItemHasProperty(vItem, LastModified) Then
                If IsDate(vItem.LastModified) Then
                    GetProperty = PropFormat(vItem.LastModified, FormatString)
                End If
            End If
        Case Property = LastUpdated
            If Tools.ItemHasProperty(vItem, LastUpdated) Then
                If IsDate(vItem.LastUpdated) Then
                    GetProperty = PropFormat(vItem.LastUpdated, FormatString)
                End If
            End If
        Case Property = ListCount
            Tools.Do_Events 5
            GetProperty = 0
            On Error Resume Next
            GetProperty = vItem.ListCount
            On Error GoTo 0
    End Select
    On Error GoTo 0
    Exit Function
HadError:
    GetProperty = vbNull
End Function
Private Function PropFormat(ByRef ItemToFormat As Variant, Optional ByVal FormatString As String)
    If Len(Trim(FormatString)) > 0 Then
        On Error Resume Next
        ItemToFormat = Format(ItemToFormat, FormatString)
        On Error GoTo 0
    End If
                
End Function

''''''Public Function SelectCurrentGridRecord(ByRef SheridanGrid As Object, Optional ByVal bPreserveCurrentSelections As Boolean)
''''''
''''''    If Tools.ItemHasProperty(SheridanGrid, SelBookmarks) Then
''''''        If Not bPreserveCurrentSelections Then
''''''            SheridanGrid.SelBookmarks.RemoveAll
''''''        End If
''''''
''''''        On Error Resume Next
''''''        SheridanGrid.SelBookmarks.Add SheridanGrid.RowBookmark(0)
''''''        On Error GoTo 0
''''''
''''''        DoEvents
''''''    End If
''''''
''''''End Function


Public Function sqlBuildSelect(ByRef oSQLobject As SQLSelectObject) As String
    
    Dim sSQL As String
    sSQL = "SELECT " & oSQLobject.Fields & " FROM " & oSQLobject.From
    If Trim(oSQLobject.Where) <> "" Then
        sSQL = sSQL & " WHERE " & oSQLobject.Where
    End If
    If Trim(oSQLobject.OrderBy) <> "" Then
        sSQL = sSQL & " ORDER BY " & oSQLobject.OrderBy
    End If
    sqlBuildSelect = sSQL
    
End Function




Public Function sqlParseSelect(ByVal sSQL As String, ByRef oSQLobject As SQLSelectObject) As Boolean
    
    Dim iFIELDS As Integer
    Dim iFROM As Integer
    Dim iWHERE As Integer
    Dim iORDERBY As Integer
    Dim iHAVING As Integer
    Dim iGROUPBY As Integer
    Dim sUSQL As String
    
    sUSQL = UCase(sSQL)
    
    oSQLobject.Fields = ""
    oSQLobject.From = ""
    oSQLobject.GroupBy = ""
    oSQLobject.Having = ""
    oSQLobject.OrderBy = ""
    oSQLobject.Where = ""
    
    ' Abort If Not A SQL Statement
    If Left(Trim(sUSQL), 7) <> "SELECT " Then
        sqlParseSelect = False
        Exit Function
    End If
    
    iFIELDS = InStr(sUSQL, "SELECT ")
    iFROM = InStr(sUSQL, " FROM ")
    iWHERE = InStr(sUSQL, " WHERE ")
    iORDERBY = InStr(sUSQL, " ORDER BY ")
    'iHAVING = InStr(sUSQL, " HAVING ")
    'iGROUPBY = InStr(sUSQL, " GROUP BY ")
    
    oSQLobject.Fields = Mid(sSQL, iFIELDS + 7, iFROM - (iFIELDS + 6))
    
    Dim iNext As Integer
    iNext = Len(sSQL)
    If iWHERE > 0 Then
        iNext = iWHERE
    End If
    If iNext > iORDERBY And iORDERBY > 0 Then
        iNext = iORDERBY
    End If
    oSQLobject.From = Mid(sSQL, iFROM + 6, iNext - (iFROM + 5))
    
    If iWHERE > 0 Then
        iNext = Len(sSQL)
        If iORDERBY > 0 Then
            iNext = iORDERBY
        End If
        
        oSQLobject.Where = Mid(sSQL, iWHERE + 7, iNext - (iWHERE + 6))
    End If
    
    If iORDERBY > 0 Then
        oSQLobject.OrderBy = Mid(sSQL, iORDERBY + 9)
    End If
    
    sqlParseSelect = True
    
End Function



'''''Public Function StringIs(ByVal sStringToCheck As String, ByVal WhatToCheck As StringIsTypes) As Boolean
'''''
'''''    Dim sCheck As String
'''''    Dim i As Integer
'''''    Select Case True
'''''        Case WhatToCheck = OnlyNumbers
'''''            sCheck = " 0123456789"
'''''        Case WhatToCheck = MathematicalCharacters
'''''            sCheck = " 0123456789./-+=*"
'''''        Case WhatToCheck = NumbersAndMoneyPunctuation
'''''            sCheck = " 0123456789.$,-"
'''''        Case WhatToCheck = NumbersAndNumericPunctuation
'''''            sCheck = " 0123456789.-"
'''''    End Select
'''''
'''''    StringIs = True
'''''    For i = 1 To Len(sStringToCheck)
'''''        If Not InStr(sCheck, Mid(sStringToCheck, i, 1)) > 0 Then
'''''            StringIs = False
'''''            Exit For
'''''        End If
'''''    Next i
'''''End Function


''''Public Function LowestVal(ByVal lValue1 As Long, ByVal lValue2 As Long) As Long
''''    If lValue1 < lValue2 Then
''''        LowestVal = lValue1
''''    Else
''''        LowestVal = lValue2
''''    End If
''''End Function


'''''Public Function bGetSetting(ByVal appname As String, ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As Boolean) As Boolean
'''''    bGetSetting = ForceBool(GetSetting(appname, Section, Key, CStr(DefaultValue)))
'''''End Function
'''''Public Function iGetSetting(ByVal appname As String, ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As Integer) As Integer
'''''    iGetSetting = ForceInt(GetSetting(appname, Section, Key, CStr(DefaultValue)))
'''''End Function
'''''Public Function lGetSetting(ByVal appname As String, ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As Long) As Long
'''''    lGetSetting = ForceLong(GetSetting(appname, Section, Key, CStr(DefaultValue)))
'''''End Function
'''''
'''''Public Function bSaveSetting(ByVal appname As String, ByVal Section As String, ByVal Key As String, ByVal Value As Boolean) As Boolean
'''''    SaveSetting appname, Section, Key, CStr(Value)
'''''End Function
'''''Public Function iSaveSetting(ByVal appname As String, ByVal Section As String, ByVal Key As String, ByVal Value As Integer) As Integer
'''''    SaveSetting appname, Section, Key, CStr(Value)
'''''End Function
'''''Public Function lSaveSetting(ByVal appname As String, ByVal Section As String, ByVal Key As String, ByVal Value As Long) As Long
'''''    SaveSetting appname, Section, Key, CStr(Value)
'''''End Function


Public Function ItemIsInList(ByVal sItem As String, ByRef vList As Object, Optional ByVal bOneBased As Boolean) As Boolean
    ItemIsInList = False
    If Tools.ItemHasProperty(vList, ListCount) And Tools.ItemHasProperty(vList, List) Then
        Dim i As Integer
        For i = IIf(bOneBased, 1, 0) To Tools.GetProperty(vList, ListCount) - IIf(bOneBased, 0, 1)
            If sEquals(sItem, vList.List(i)) Then
                ItemIsInList = True
                Exit For
            End If
        Next i
    End If
    
End Function

'''''Public Function SaveObjectCoordinates(ByRef oWorkObject As Object, ByRef oCoords As ObjectCoordinates)
'''''    oCoords.objHeight = oWorkObject.Height
'''''    oCoords.objWidth = oWorkObject.Width
'''''    oCoords.objLeft = oWorkObject.Left
'''''    oCoords.objTop = oWorkObject.Top
'''''
'''''End Function
'''''
'''''Public Function RestoreObjectCoordinates(ByRef oWorkObject As Object, ByRef oCoords As ObjectCoordinates)
'''''    On Error Resume Next
'''''    oWorkObject.Height = oCoords.objHeight
'''''    oWorkObject.Width = oCoords.objWidth
'''''    oWorkObject.Left = oCoords.objLeft
'''''    oWorkObject.Top = oCoords.objTop
'''''    On Error GoTo 0
'''''
'''''End Function


'''''Public Function ApplyMinMaxToProgressBar(ByRef oProgressBar As Object, ByVal lMin As Long, ByVal lMax As Long)
'''''
'''''    Dim lCurMin As Long
'''''    Dim lCurMax As Long
'''''    lCurMin = oProgressBar.Min
'''''    lCurMax = oProgressBar.Max
'''''
'''''    If lMin > lCurMax Then
'''''        oProgressBar.Max = lMax
'''''        oProgressBar.Min = lMin
'''''    Else
'''''        oProgressBar.Min = lMin
'''''        oProgressBar.Max = lMax
'''''
'''''    End If
'''''
'''''End Function


'''''Public Function SortCollection(ByRef colWork As Collection) As Boolean
'''''    MsgBox "This function hasn't been coded yet.  =o)"
'''''End Function






''''''Public Function WaitForFileToExist(ByVal sFileToLookFor As String, Optional ByVal iTimeout As Integer = 15) As Boolean
''''''    Dim iCounter As Integer
''''''    Do While Not fileExists(sFileToLookFor) And iCounter < iTimeout
''''''        DoEvents
''''''        iCounter = iCounter + 1
''''''
''''''    Loop
''''''    WaitForFileToExist = fileExists(sFileToLookFor)
''''''End Function


'''''Public Function CreateUniqueFileName(ByVal sCurrFile As String) As String
'''''
'''''    Dim i As Integer
'''''    Dim iExt As Integer
'''''    Dim sExt As String
'''''    Dim sBase As String
'''''    i = 1
'''''    iExt = RInstr(sCurrFile, ".")
'''''    If iExt = 0 Then
'''''        sExt = ""
'''''        sBase = sCurrFile
'''''    Else
'''''        sExt = Mid(sCurrFile, iExt)
'''''        sBase = Left(sCurrFile, iExt - 1)
'''''    End If
'''''    Do While fileExists(sCurrFile)
'''''        i = i + 1
'''''        sCurrFile = sBase & "(" & Trim(i) & ")" & sExt
'''''
'''''    Loop
'''''    CreateUniqueFileName = sCurrFile
'''''End Function

Sub FilesSearch(DrivePath As String, Ext As String)
       Dim XDir() As String
       Dim TmpDir As String
       Dim FFound As String
       Dim DirCount As Integer
       Dim X As Integer
       '     'Initialises Variables
       DirCount = 0
       ReDim XDir(0) As String
       XDir(DirCount) = ""

              If Right(DrivePath, 1) <> "\" Then
                     DrivePath = DrivePath & "\"
              End If

       '     'Enter here the code for showing the path being
       '     'search. Example: Form1.label2 = DrivePath
       '     'Search for all directories and store in the
       '     'XDir() variable

 DoEvents
  TmpDir = Dir(DrivePath, vbDirectory)

   Do While TmpDir <> ""

      If TmpDir <> "." And TmpDir <> ".." Then

 If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
         XDir(DirCount) = DrivePath & TmpDir & "\"
         DirCount = DirCount + 1
         ReDim Preserve XDir(DirCount) As String
 End If


     End If

        TmpDir = Dir
  Loop

   '     'Searches for the files given by extension Ext
   FFound = Dir(DrivePath & Ext)

       Do Until FFound = ""
   '     'Code in here for the actions of the files found.
   '     'Files found stored in the variable FFound.
   '     'Example: Form1.list1.AddItem DrivePath & FFound
          FFound = Dir
     Loop


   '     'Recursive searches through all sub directories

    For X = 0 To (UBound(XDir) - 1)
          FilesSearch XDir(X), Ext
   Next X

 End Sub



''''Public Function NRTrim(ByVal vValue As Variant) As String
''''    NRTrim = RTrim(vValue & "")
''''End Function

'''''Public Function DeQuote(ByVal sStringToUnquote As String) As String
'''''    ' This function will remove quotes that are wrapped around a text string you want.
'''''
'''''    Dim sQuoteType As String
'''''    Select Case Left(Trim(sStringToUnquote), 1)
'''''        Case """"
'''''            sQuoteType = """"
'''''        Case "'"
'''''            sQuoteType = "'"
'''''    End Select
'''''    If sQuoteType <> "" Then
'''''        sStringToUnquote = Trim(sStringToUnquote)
'''''        Do While Left(sStringToUnquote, 1) = sQuoteType
'''''            sStringToUnquote = Mid(sStringToUnquote, 2)
'''''        Loop
'''''        Do While Right(sStringToUnquote, 1) = sQuoteType
'''''            sStringToUnquote = Left(sStringToUnquote, Len(sStringToUnquote) - 1)
'''''        Loop
'''''    End If
'''''
'''''    DeQuote = sStringToUnquote
'''''
'''''End Function

'''''Public Function DeQuote(ByVal sStringToUnquote As String) As String
'''''    ' This function will remove quotes that are wrapped around a text string you want.
'''''
'''''    Dim sQuoteType As String
'''''    Select Case Left(Trim(sStringToUnquote), 1)
'''''        Case """"
'''''            sQuoteType = """"
'''''        Case "'"
'''''            sQuoteType = "'"
'''''    End Select
'''''    If sQuoteType <> "" Then
'''''        sStringToUnquote = Trim(sStringToUnquote)
'''''        Do While Left(sStringToUnquote, 1) = sQuoteType
'''''            sStringToUnquote = Mid(sStringToUnquote, 2)
'''''        Loop
'''''        Do While Right(sStringToUnquote, 1) = sQuoteType
'''''            sStringToUnquote = Left(sStringToUnquote, Len(sStringToUnquote) - 1)
'''''        Loop
'''''    End If
'''''
'''''    DeQuote = sStringToUnquote
'''''
'''''End Function


' Should be moved to Tools
''''Public Function WordVal(ByVal sWord As String) As Variant
''''    ' Converts a byte WORD into it's corresping value
''''    '  i.e.  chr(0) & chr(67) = 67
''''    '        chr(3) & chr(31) = 799
''''    Dim vReturn As Variant
''''    Dim dBase As Double
''''    Dim iCharVal As Integer
''''    dBase = 1
''''    Dim i As Integer
''''    For i = Len(sWord) To 1 Step -1
''''        iCharVal = Asc(Mid(sWord, i, 1))
''''        vReturn = vReturn + (iCharVal * dBase)
''''        dBase = dBase * 256
''''    Next i
''''    WordVal = vReturn
''''
''''End Function

Public Function StringToBits(ByVal sSource As String) As String
    ' Converts a string of characters into it's corresping 0/1 bit value string
    ' i.e. StringToBits("FF")
    Dim sReturn As String
    Dim i As Integer
    For i = 1 To Len(sSource)
        sReturn = sReturn & Tools.NumberToBits(Asc(Mid(sSource, i, 1)), 8)
    Next i
    StringToBits = sReturn
End Function


''''Public Function HexVal(ByVal sHexString As String) As Variant
''''    ' Converts a hex string into it's corresponding value
''''    ' i.e.  HexVal( "FF") = 255
''''
''''    Dim vReturn As String
''''    Dim vBase As Variant
''''
''''    vReturn = 0
''''    vBase = 1
''''
''''    Dim i As Integer
''''    For i = Len(sHexString) To 1 Step -1
''''        Dim sCurrChar
''''        Dim iCurrVal
''''        sCurrChar = UCase(Mid(sHexString, i, 1))
''''        Select Case True
''''            Case InStr("0123456789", sCurrChar) > 0
''''                iCurrVal = Val(sCurrChar)
''''            Case InStr("ABCDEF", sCurrChar) > 0
''''                iCurrVal = 9 + InStr("ABCDEF", sCurrChar)
''''            Case Else
''''                iCurrVal = 0
''''        End Select
''''        vReturn = vReturn + (vBase * iCurrVal)
''''        vBase = vBase * 16
''''    Next i
''''    HexVal = vReturn
''''End Function




Public Function EmptyToSpace(ByRef sWork As String)
    If Trim(sWork) = "" Then
        sWork = " "
    End If
End Function





''''Public Function SelectedOptionIndex(ByRef oOptions As Object, Optional ByVal iDefault As Integer) As Integer
''''
''''    If Not IsMissing(iDefault) Then
''''        SelectedOptionIndex = iDefault
''''    Else
''''        SelectedOptionIndex = -1
''''    End If
''''    Dim i As Integer
''''    For i = 0 To oOptions.Count - 1
''''        If oOptions(i).Value <> 0 Then
''''            SelectedOptionIndex = i
''''            Exit Function
''''        End If
''''    Next i
''''
''''End Function


Public Function EditMask(iKeyAscii As Integer, sType As enumType) As Integer
'Purpose:       Accepts an ASCII key code (iKeyAscii) and returns 0
'               if the key code is not of the type passed as sType.
'sType values:  etDATE_SLASH: 0 to 9, /, Backspace
'               etPOS_NUM: 0 to 9, Backspace
'               etLETTER: A to Z, a to z, Backspace
'               etTIME: 0 to 9, :, A, M, P, [Space]
    EditMask = iKeyAscii
    If iKeyAscii <> vbKeyBack Then
        Select Case sType
        Case etDATE_SLASH
            Select Case iKeyAscii
            Case vbKey0 To vbKey9, 47 '/
            Case Else
                EditMask = 0
            End Select
        Case etPOS_NUM
            Select Case iKeyAscii
            Case vbKey0 To vbKey9
            Case Else
                EditMask = 0
            End Select
        Case etLETTER
            Select Case iKeyAscii
            Case vbKeyA To vbKeyZ, vbKeyA To vbKeyZ
            Case Else
                EditMask = 0
            End Select
        Case etTIME
            Select Case iKeyAscii
            Case vbKeyA, vbKeyM, vbKeyP, vbKey0 To vbKey9, vbKeySpace, 58 ':
            Case Else
                EditMask = 0
            End Select
        End Select
    End If

End Function



''Public Function CreateDir(ByVal sPath As String) As Boolean
''    If DirectoryExists(sPath) Then
''        CreateDir = True
''        Exit Function
''    End If
''
''    Dim sCurrPath As String
''    sPath = Tools.AppendCharacter(sPath, "\")
''    Dim i As Integer
''    For i = 1 To WordCnt(sPath, "\")
''        sCurrPath = sCurrPath & Word(sPath, i, "\") & "\"
''        If Not DirectoryExists(sCurrPath) Then
''            On Error Resume Next
''            MkDir sCurrPath
''            If Err <> 0 Then
''                On Error GoTo 0
''                CreateDir = False
''                Exit Function
''            End If
''            On Error GoTo 0
''        End If
''    Next i
''    CreateDir = True
''
''End Function


'''''Public Function R_Trim(ByRef sTrimMe As String, Optional ByVal sChars2Trim As String) As String
'''''    Do While Right(sTrimMe, Len(sChars2Trim)) = sChars2Trim
'''''        sTrimMe = Left(sTrimMe, Len(sTrimMe) - Len(sChars2Trim))
'''''    Loop
'''''    R_Trim = sTrimMe
'''''End Function




' ==================================================
' ListProps()
'   Returns a string containing a vbCrLf delimited
'   listing of property names from an object.  Good
'   for spying on "Dynamic Properties".
' ==================================================
Public Function ListProps(ByRef oObjectToList4 As Object) As String
    
    Dim sReturn As String
    Dim i As Integer
    For i = 0 To oObjectToList4.Properties.Count - 1
        sReturn = sReturn & oObjectToList4.Properties(i).Name & " = " & oObjectToList4.Properties(i).Value & vbCrLf
    Next i
    ListProps = sReturn
    
End Function









Public Function arrayDebugDisplay(ByRef aVar As Variant)
    
    Dim iRows As Integer
    Dim iCols As Integer
    Dim bMultiColumn As Boolean
    
    iRows = UBound(aVar, 1)
    On Error Resume Next
    iCols = UBound(aVar, 2)
    bMultiColumn = (Err.Number = 0)
    On Error GoTo 0
    
    Dim sCurrLine As String
    Dim r As Integer
    Dim c As Integer
    For r = 0 To iRows
        sCurrLine = ""
        If bMultiColumn Then
            For c = 0 To iCols
                sCurrLine = sCurrLine & sPadR(CStr(aVar(r, c)), 15) & " | "
                
            Next c
        Else
            sCurrLine = aVar(r)
        End If
        Debug.Print sCurrLine
        
    Next r
    
End Function


Public Function arrayAddRow(aVar() As Variant, rVar() As Variant, Optional ByVal bStartNew As Boolean = False)
    
    Static iCurRow As Integer
    
    If bStartNew Then
        iCurRow = 0
    End If
    
    Dim iNewCols As Integer
    iNewCols = UBound(rVar, 1)
    
    Dim i As Integer
    For i = 0 To iNewCols
        aVar(iCurRow, i) = rVar(iNewCols)
        
    Next i
    iCurRow = iCurRow + 1
    
End Function
