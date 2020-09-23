VERSION 5.00
Object = "{58635701-4313-11D1-9D7F-CD6975009A1F}#1.0#0"; "RD.OCX"
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Shop.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   8160
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   9
      Top             =   7200
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   11
         Left            =   0
         Picture         =   "Shop.frx":17E8AE
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   37
         Top             =   0
         Width           =   1920
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   0
         Picture         =   "Shop.frx":1818F0
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   27
         Top             =   480
         Width           =   1920
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit Shop"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   7200
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
      Begin REALDIGITSLib.RD Score 
         Height          =   225
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         ThreeDView      =   0   'False
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   10
         Left            =   0
         Picture         =   "Shop.frx":184932
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   39
         Top             =   0
         Width           =   1920
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   0
         Picture         =   "Shop.frx":187974
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   26
         Top             =   480
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   3360
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   7
      Top             =   7200
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   0
         Picture         =   "Shop.frx":18A9B6
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   25
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Here can you buy ammo for your extra cannons."
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   480
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   6
      Top             =   7200
      Width           =   1935
      Begin REALDIGITSLib.RD Price 
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "3000"
         Length          =   7
         ThreeDView      =   0   'False
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   0
         Picture         =   "Shop.frx":18D9F8
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   21
         Top             =   0
         Width           =   1920
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buy"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   34
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   480
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   5
      Top             =   5280
      Width           =   1935
      Begin REALDIGITSLib.RD Price 
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "200000"
         Length          =   7
         ThreeDView      =   0   'False
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   0
         Picture         =   "Shop.frx":190A3A
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   20
         Top             =   0
         Width           =   1920
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buy"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   32
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   3360
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   4
      Top             =   5280
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   0
         Picture         =   "Shop.frx":193A7C
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   24
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Here can you buy new cannons for your ship."
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3360
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   0
         Picture         =   "Shop.frx":196ABE
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   23
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Here can you repair your  shields from your ship."
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3360
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   0
         Picture         =   "Shop.frx":199B00
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   22
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Here can you up grade your shooting speed."
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   480
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
      Begin REALDIGITSLib.RD Price 
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "6000"
         Length          =   7
         ThreeDView      =   0   'False
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   0
         Picture         =   "Shop.frx":19CB42
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   19
         Top             =   0
         Width           =   1920
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buy"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   30
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   480
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
      Begin REALDIGITSLib.RD Price 
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "8000"
         Length          =   7
         ThreeDView      =   0   'False
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   0
         Picture         =   "Shop.frx":19FB84
         ScaleHeight     =   480
         ScaleWidth      =   1920
         TabIndex        =   18
         Top             =   0
         Width           =   1920
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buy"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   28
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Sheeld 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Speed 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Points available"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   7200
      TabIndex        =   40
      Top             =   1200
      Width           =   1905
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cannon Ammo"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   13
      Top             =   6840
      Width           =   1980
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Cannons"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   12
      Top             =   4920
      Width           =   1980
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shield"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   11
      Top             =   3000
      Width           =   1860
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shooting Speed"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   1080
      Width           =   1905
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Speed2 As Long
Dim CSpeed As Long
Dim Level As Long
Dim j

Function Reset()
On Error Resume Next
Price(0) = 8000
CSpeed = 0
Speed2 = 0
Level = 0
End Function

Function Load()
Dim I As Long
For I = 0 To 3
Price(I).Visible = False
Next I
Score.Visible = False
For I = 0 To Picture2.Count - 1
Picture2(I).Top = 0
Next I
Me.Show
''On Error Resume Next
Label11.Caption = Form1.GetCannons
Label12.Caption = Form1.GetCondition - 15
Score.Digits = Form1.Score.Caption
Level = Form1.Getlevel
Price(1).Digits = Level * 3000
Me.Show
Sheeld = Form1.Getsheeld
Speed = CSpeed
Speed2 = Form1.GetSpeed
Me.Enabled = False
OpenWindow
Command6(0).ToolTipText = "Current shooting speed " & Speed.Caption & "    Max 5"
Command6(1).ToolTipText = "Current shield " & Sheeld.Caption & "    Max 2655"
Command6(2).ToolTipText = "Current cannons " & Label11.Caption & "    Max 3"
Command6(3).ToolTipText = "Current ammo " & Label12.Caption & "    Max 9999"
End Function

Private Sub Command2_Click()
On Error Resume Next
CSpeed = Speed
Me.Hide
Form1.SetCondition Label12.Caption + 15
Form1.Setsheeld Sheeld.Caption
Form1.Score.Caption = Score.Digits
Form1.SetPause False
Form1.SetSpeed Speed2
Form1.Label3.Visible = False
Form1.SetCannons Label11.Caption
Form1.Show
End Sub


Private Sub Command6_Click(Index As Integer)
On Error Resume Next
Select Case Index

Case 0
If Speed2 < 1 Then
Msg2 "Shooting Speed is on maximum!", 1
Exit Sub
End If
If Int(Score.Digits) < Int(Price(0).Digits) Then
Msg2 "You can not pay this!", 1
Exit Sub
End If
Price(0).Digits = Price(0).Digits * 2
Score.Digits = Score.Digits - Price(0).Digits
Speed = Speed + 1
Speed2 = Speed2 - 1

Case 1
If Sheeld.Caption > 2654 Then
Msg2 "Sheeld is on maximum!", 1
Exit Sub
End If
If Int(Score.Digits) < Int(Price(1).Digits) Then
Msg2 "You can not pay this!", 1
Exit Sub
End If
If Sheeld.Caption + 500 > 2654 Then
Sheeld.Caption = 2655
Else
Sheeld.Caption = Sheeld.Caption + 500
End If
Score.Digits = Score.Digits - Price(1).Digits

Case 3
If Label11.Caption < 2 Then
Msg2 "No cannons how are useing te ammo.", 1
Exit Sub
End If
If Label12.Caption > 9998 Then
Msg2 "ammo is on maximum!", 1
Exit Sub
End If
If Int(Score.Digits) < Int(Price(3).Digits) Then
Msg2 "You can not pay this!", 1
Exit Sub
End If
If Label12.Caption + 200 > 9998 Then
Label12.Caption = 9999
Else
Label12.Caption = Label12.Caption + 200
End If
Score.Digits = Score.Digits - Price(3).Digits

Case 2
If Label11.Caption > 2 Then
Msg2 "Cannons are on maximum!", 1
Exit Sub
End If
If Int(Score.Digits) < Int(Price(2).Digits) Then
Msg2 "You can not pay this!", 1
Exit Sub
End If
Label12.Caption = 1200
Label11.Caption = Label11.Caption + 1
Score.Digits = Score.Digits - Price(2).Digits

End Select

Command6(0).ToolTipText = "Current shooting speed " & Speed.Caption & "    Max 5"
Command6(1).ToolTipText = "Current shield " & Sheeld.Caption & "    Max 2655"
Command6(2).ToolTipText = "Current cannons " & Label11.Caption & "    Max 3"
Command6(3).ToolTipText = "Current ammo " & Label12.Caption & "    Max 9999"
End Sub

Function OpenWindow()
Dim E1, I As Long
Do Until E1 = Picture2.Count
For I = 1 To 10000
Next I
DoEvents
Picture2(E1).Top = Picture2(E1).Top + 1
If Picture2(E1).Top > 479 Then
If E1 = 0 Then Price(0).Visible = True
If E1 = 1 Then Price(1).Visible = True
If E1 = 2 Then Price(2).Visible = True
If E1 = 3 Then Price(3).Visible = True
If E1 = 11 Then Score.Visible = True
E1 = E1 + 1
End If
Loop
Me.Enabled = True
End Function

