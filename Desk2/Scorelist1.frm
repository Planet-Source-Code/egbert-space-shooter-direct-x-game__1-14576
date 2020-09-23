VERSION 5.00
Begin VB.Form ShowScoreList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Score list"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   30
      ScaleHeight     =   5175
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "ShowScoreList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Command1.Visible = False
Dim Y1 As Single
Picture1.Picture = ScoreList.Image
Y1 = 0
Me.Show
Do
DoEvents
Picture1.Top = Y1
'Me.PaintPicture ScoreList.Image, 10, Y1
Y1 = Y1 + 0.5
Loop Until Y1 > Me.Width - 600
Me.Hide
Form2.Show
End Sub

Function Load()
On Error Resume Next
Dim A As Long
Dim Y1 As Single
Picture1.Picture = ScoreList.Image
Y1 = Me.Height - 600
Me.Show
Do
DoEvents
Picture1.Top = Y1
'Me.PaintPicture ScoreList.Image, 10, Y1
Y1 = Y1 - 0.5
Loop Until Y1 < 0
Command1.Visible = True
End Function

