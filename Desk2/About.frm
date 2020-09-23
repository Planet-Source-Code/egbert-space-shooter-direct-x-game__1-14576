VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      X1              =   2160
      X2              =   2160
      Y1              =   1440
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      X1              =   2160
      X2              =   480
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   480
      X2              =   480
      Y1              =   2640
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   480
      X2              =   2160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Check for updates at http://212.120.70.98/eb"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Egberttheone@hotmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Buid by Egbert Boer from Holland"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "About Spaceshooter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   1000
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   1000
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This game is easy. Just shoot the rocks and the other planes. But be sure NOTHING hits you !"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   4920
      Left            =   0
      Picture         =   "About.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4920
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Label7.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Show
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbYellow
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlue

End Sub

