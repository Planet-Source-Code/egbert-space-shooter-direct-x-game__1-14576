VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Laucher"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Launcher.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Res2 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Text            =   "0"
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Res1 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   3
      Left            =   1350
      TabIndex        =   3
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hi-score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   2
      Left            =   1170
      TabIndex        =   2
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   1
      Left            =   1305
      TabIndex        =   1
      Top             =   2040
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   0
      Left            =   1005
      TabIndex        =   0
      Top             =   840
      Width           =   1395
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim Wait, Fool
Dim I
Dim j

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then End
Dim strBuffer As String
Dim lngReturn As Long
Dim strWindowsSystemDirectory As String
Dim FileName As String
Dim Z As Long
For Z = 2 To 4
Select Case Z
Case 2
FileName = App.Path & "\RD.ocx"
Case 3
FileName = App.Path & "\OLEAUT32.DLL"
Case 4
FileName = App.Path & "\MSVBVM50.DLL"
End Select
If Dir(FileName) = "" Then
MsgBox "File not found : " & FileName, vbExclamation + vbOKOnly, "error"
End
End If
Next Z
Wait = 100
Form5.Show
Form4.Hide
Form5.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For I = 0 To 4
Label1(I).ForeColor = vbYellow
Next I
End Sub

Private Sub Label1_Click(Index As Integer)
Dim FileName
On Error Resume Next
If Index < 4 Then
FileName = App.Path & "\Sound\" & "Main2.wav"
j = sndPlaySound(FileName, &H1)
Else
FileName = App.Path & "\Sound\" & "Main2.wav"
j = sndPlaySound(FileName, &H0)
End If
Form5.Hide
Me.Hide
Select Case Index
Case 0
Dim intWidth As Integer
Dim intHeight As Integer
Res1 = Screen.Width \ Screen.TwipsPerPixelX
Res2 = Screen.Height \ Screen.TwipsPerPixelY
Dim typDevM As typDevMODE
Dim lngResult As Long
Dim intAns    As Integer
lngResult = EnumDisplaySettings(0, 0, typDevM)
With typDevM
.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
.dmPelsWidth = 800  'ScreenWidth (640,800,1024, etc)
.dmPelsHeight = 600 'ScreenHeight (480,600,768, etc)
End With
lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
Select Case lngResult
Case DISP_CHANGE_RESTART
intAns = MsgBox("You must restart your computer to apply these changes." & vbCrLf & vbCrLf & "Do you want to restart now?", vbYesNo + vbSystemModal, "Screen Resolution")
If intAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
Case DISP_CHANGE_SUCCESSFUL
Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
Form1.Load
Case Else
MsgBox "Display mode error can not switch to 800x600 game can not be loaded.", vbSystemModal, "Error"
End Select
Case 1
Form5.Show
Case 3
Form3.Show
Case 2
ScoreList.Load
ShowScoreList.Load
Case 4
End
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim FileName
If Wait = Index Then GoTo 1
FileName = App.Path & "\Sound\" & "Main.wav"
j = sndPlaySound(FileName, &H1)
Wait = Index
1:
For I = 0 To 4
Label1(I).ForeColor = vbYellow
Next I
Label1(Index).ForeColor = vbRed
End Sub

