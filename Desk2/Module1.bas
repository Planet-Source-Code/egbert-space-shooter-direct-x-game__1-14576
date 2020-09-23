Attribute VB_Name = "Module1"
Option Explicit
Dim Respons As Boolean
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const MAX_PATH = 260
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NODEFAULT = &H2
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328


#If Win32 Then
Public Const HWND_TOPMOST& = -1
#Else
Public Const HWND_TOPMOST& = -1
#End If 'WIN32

#If Win32 Then
 Const SWP_NOMOVE& = &H2
 Const SWP_NOSIZE& = &H1
#Else
 Const SWP_NOMOVE& = &H2
 Const SWP_NOSIZE& = &H1
#End If 'WIN32

#If Win32 Then
 Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#Else
 Declare Sub SetWindowPos Lib "user" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#End If 'WIN32


Function StayOnTop(Form As Form) 'EX: Call StayOnTop(Me)
Dim lFlags As Long
Dim lStay As Long

lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hWnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Private Function HiWord(ByVal l As Long) As Integer
    l = l \ &H10000
    
    HiWord = Val("&H" & Hex$(l))
End Function

Private Function LoWord(ByVal l As Long) As Integer
    l = l And &HFFFF&
    
    LoWord = Val("&H" & Hex$(l))
End Function

Function Msg2(Prompt As String, Number As Long) As Long  ''1 = vbokonly , 2 = vbyesno
Dim A
'' 1 = Yes
'' 2 = No
'' 3 = Ok
If Number = 1 Then
MsgBox Prompt, vbOKOnly + vbInformation
Msg2 = 3
Else
A = MsgBox(Prompt, vbYesNo + vbInformation)
If A = vbYes Then
Msg2 = 1
Else
Msg2 = 2
End If
End If
End Function
 
Function SetRespons(Res As Long)
Respons = Res
End Function

