VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Option"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Option.frx":0000
   ScaleHeight     =   3750
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check5 
      BackColor       =   &H00000000&
      Caption         =   "Game Speed enabled"
      Height          =   195
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   840
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "Game Speed enabled"
      Height          =   195
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   9
      Top             =   1200
      Width           =   210
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Music"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Value           =   1  'Checked
      Width           =   225
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.HScrollBar S2 
      Height          =   255
      LargeChange     =   100
      Left            =   240
      Max             =   4000
      SmallChange     =   10
      TabIndex        =   1
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Game Speed enabled"
      Height          =   195
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flip Stereo"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic Speed control enable"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   2265
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed control enable"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sound fx"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Music"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   150
      Width           =   420
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j

Private Sub Form_Load()
Load
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed
Label4.ForeColor = vbRed
End Sub

Private Sub Label3_Click()
Dim Filename As String
Dim j
Filename = App.Path & "\Sound\" & "Main2.wav"
j = sndPlaySound(Filename, &H1)
On Error Resume Next
j = sndPlaySound(vbNullString, &H1)
Me.Hide
Form2.Show
Save
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Filename As String
Dim j
If Label3.ForeColor = vbRed Then
Filename = App.Path & "\Sound\" & "Main.wav"
j = sndPlaySound(Filename, &H1)
End If
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
Dim Wait As Long
Dim Filename
Filename = App.Path & "\Sound\" & "Left.wav"
j = sndPlaySound(Filename, &H0)
Filename = App.Path & "\Sound\" & "Right.wav"
j = sndPlaySound(Filename, &H0)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Filename As String
Dim j
If Label4.ForeColor = vbRed Then
Filename = App.Path & "\Sound\" & "Main.wav"
j = sndPlaySound(Filename, &H1)
End If
Label4.ForeColor = vbBlue
End Sub

Function Save()
On Error Resume Next
Close #1
Dim Filename As String
Filename = App.Path & "\Data\Option.dat"
Open Filename For Output As #1
Write #1, Check1.Value, Check2.Value, Check3.Value, Check4.Value, Check5.Value, S2.Value
Close #1
End Function

Function Load()
Dim Filename As String
Dim A, B, C, D, E, F
Filename = App.Path & "\Data\Option.dat"
2:
On Error GoTo 1
Open Filename For Input As #1
Input #1, A, B, C, D, E, F
Close #1
Check1.Value = A
Check2.Value = B
Check3.Value = C
Check4.Value = D
S2.Value = E
Check5.Value = F
Exit Function
1:
Save
GoTo 2
End Function

