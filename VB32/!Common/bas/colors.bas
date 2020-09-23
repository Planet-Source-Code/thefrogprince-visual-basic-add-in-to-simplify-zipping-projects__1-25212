Attribute VB_Name = "colors"
Option Explicit

' System Colors
Public Const TitleBarBack = &H80000002
Public Const TitleBarText = &H80000009
Public Const ScrollBars = &H80000000
Public Const ApplicationWorkspace = &H8000000C
Public Const ButtonFace = &H8000000F
Public Const ButtonText = &H80000012
Public Const ButtonHighlight = &H80000014
Public Const BackColorMenuBar = &H80000004

Public Const DisabledButtonForeColor = &H80000011
Public Const WindowBackground = &H80000005

' TOTALLY DIFFERENT COLOR SET TO BE USED WITH COMPOSER  =o|
#If IsPermComp Then
    Public Const BackColorREHEARSAL = &HFFFFFF
    Public Const BackColorGradientREHEARSAL = &HFFFFFF
    Public Const BackColorLIVE = &HE0FFFF
    Public Const BackColorGradientLIVE = 8454143
    Public Const BackColorPRACTICE = &HFFFFC0
    Public Const BackColorGradientPRACTICE = 16744448
    Public Const RealBackColorLIVE = &HE0E0E0
    Public Const RealBackColorGradientLIVE = 8421504
    Public Const BackColorModeling = &HC0&
#Else
    Public Const BackColorREHEARSAL = &HE0FFFF
    Public Const BackColorGradientREHEARSAL = 8454143
    Public Const BackColorLIVE = &HE0E0E0
    Public Const BackColorGradientLIVE = 8421504
    Public Const BackColorPRACTICE = &HFFFFC0
    Public Const BackColorGradientPRACTICE = 16744448
    Public Const BackColorModeling = &HC0&
#End If

Public Const ppRequiredField = &HE0FFFF

' Flat Colors
Public Const RED = &HFF&
Public Const Orange = &H80FF&
Public Const YELLOW = &HFFFF&
Public Const GREEN = &HFF00&
Public Const CYAN = &HFFFF00
Public Const BLUE = &HFF0000
Public Const BrightPurple = &HFF00FF
Public Const BLACK = &H0&
Public Const Purple = &HC000C0
Public Const DarkPurple = &H800080
Public Const WHITE = &HFFFFFF

