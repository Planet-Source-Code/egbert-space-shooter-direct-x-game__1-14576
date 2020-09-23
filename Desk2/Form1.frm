VERSION 5.00
Object = "{58635701-4313-11D1-9D7F-CD6975009A1F}#1.0#0"; "RD.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Space Driver"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0080FFFF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":030A
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Nukemask2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   960
      Picture         =   "Form1.frx":4374
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   72
      Top             =   6120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Nuke2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1080
      Picture         =   "Form1.frx":475E
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   71
      Top             =   6120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Nukemask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   1320
      Picture         =   "Form1.frx":4B48
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   70
      Top             =   5880
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Nuke1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   1200
      Picture         =   "Form1.frx":509A
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   69
      Top             =   5880
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2640
      Picture         =   "Form1.frx":55EC
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   68
      Top             =   7200
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   2640
      Picture         =   "Form1.frx":6422
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   67
      Top             =   6840
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2640
      Picture         =   "Form1.frx":7258
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   63
      Top             =   6480
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonusmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      Picture         =   "Form1.frx":808E
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   62
      Top             =   5040
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2640
      Picture         =   "Form1.frx":8EC4
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   61
      Top             =   6120
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2640
      Picture         =   "Form1.frx":9CFA
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   60
      Top             =   5760
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2640
      Picture         =   "Form1.frx":AB30
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   59
      Top             =   5400
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2640
      Picture         =   "Form1.frx":B966
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   58
      Top             =   5040
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2640
      Picture         =   "Form1.frx":C79C
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   57
      Top             =   4680
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   7080
   End
   Begin VB.PictureBox Laser 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   45
      Picture         =   "Form1.frx":D5D2
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   49
      Top             =   2280
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Laser 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      Picture         =   "Form1.frx":DAB8
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   48
      Top             =   2280
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Laser 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      Picture         =   "Form1.frx":DF9E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   47
      Top             =   2280
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   0
      Left            =   -105
      Picture         =   "Form1.frx":E484
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   1
      Left            =   1080
      Picture         =   "Form1.frx":12186
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   2
      Left            =   2280
      Picture         =   "Form1.frx":15E88
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   3
      Left            =   3480
      Picture         =   "Form1.frx":19B8A
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   4
      Left            =   4680
      Picture         =   "Form1.frx":1D88C
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   5
      Left            =   5880
      Picture         =   "Form1.frx":2158E
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   6
      Left            =   7080
      Picture         =   "Form1.frx":25290
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   7
      Left            =   8280
      Picture         =   "Form1.frx":28F92
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   8
      Left            =   9480
      Picture         =   "Form1.frx":2CC94
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   9
      Left            =   9480
      Picture         =   "Form1.frx":30996
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":34698
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   2
      Left            =   1200
      Picture         =   "Form1.frx":3839A
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   3
      Left            =   2400
      Picture         =   "Form1.frx":3C09C
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   4
      Left            =   3600
      Picture         =   "Form1.frx":3FD9E
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   5
      Left            =   4800
      Picture         =   "Form1.frx":43AA0
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   6
      Left            =   6000
      Picture         =   "Form1.frx":477A2
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   7
      Left            =   7200
      Picture         =   "Form1.frx":4B4A4
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   8
      Left            =   8400
      Picture         =   "Form1.frx":4F1A6
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Shipam 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      Picture         =   "Form1.frx":52EA8
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox Shipa 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      Picture         =   "Form1.frx":53E26
      ScaleHeight     =   975
      ScaleWidth      =   285
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox Exp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3360
      Picture         =   "Form1.frx":54DA4
      ScaleHeight     =   255
      ScaleWidth      =   345
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox Exp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3360
      Picture         =   "Form1.frx":552AE
      ScaleHeight     =   285
      ScaleWidth      =   345
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox Exp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3000
      Picture         =   "Form1.frx":55848
      ScaleHeight     =   285
      ScaleWidth      =   390
      TabIndex        =   24
      Top             =   4200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox Exp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3000
      Picture         =   "Form1.frx":55E7A
      ScaleHeight     =   255
      ScaleWidth      =   390
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox ExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3360
      Picture         =   "Form1.frx":5640C
      ScaleHeight     =   255
      ScaleWidth      =   345
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox ExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3360
      Picture         =   "Form1.frx":56916
      ScaleHeight     =   285
      ScaleWidth      =   345
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox ExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3000
      Picture         =   "Form1.frx":56EB0
      ScaleHeight     =   285
      ScaleWidth      =   390
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox ExpM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3000
      Picture         =   "Form1.frx":574E2
      ScaleHeight     =   255
      ScaleWidth      =   390
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   8040
      Width           =   11985
      Begin VB.CommandButton Command1 
         Caption         =   "Shop"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9855
         TabIndex        =   66
         Top             =   315
         Width           =   975
      End
      Begin REALDIGITSLib.RD Shops 
         Height          =   225
         Left            =   8880
         TabIndex        =   64
         Top             =   360
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "1"
         Length          =   3
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Lasers 
         Height          =   225
         Left            =   7080
         TabIndex        =   55
         Top             =   360
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         Length          =   5
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Frame11 
         Height          =   225
         Left            =   6480
         TabIndex        =   54
         Top             =   720
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "000"
         Length          =   3
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Score1 
         Height          =   225
         Left            =   4680
         TabIndex        =   52
         Top             =   360
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         ThreeDView      =   0   'False
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         ScaleHeight     =   225
         ScaleWidth      =   2625
         TabIndex        =   14
         Top             =   360
         Width           =   2655
         Begin VB.Label Shield 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Shield : 100%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   2655
         End
         Begin VB.Shape Healtm 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shops left :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   7920
         TabIndex        =   65
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ammo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   56
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Frame's sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   53
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Score 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4200
         TabIndex        =   51
         Top             =   240
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Shield :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   0
         Picture         =   "Form1.frx":57A74
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11985
      End
   End
   Begin VB.PictureBox LaserMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Picture         =   "Form1.frx":688B6
      ScaleHeight     =   495
      ScaleWidth      =   165
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Laser 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   780
      Picture         =   "Form1.frx":68D9C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox AsteroidMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   1620
      Picture         =   "Form1.frx":69282
      ScaleHeight     =   555
      ScaleWidth      =   750
      TabIndex        =   10
      Top             =   5100
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox AsteroidMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   1620
      Picture         =   "Form1.frx":6A8BC
      ScaleHeight     =   555
      ScaleWidth      =   750
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox AsteroidMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   1620
      Picture         =   "Form1.frx":6BEF6
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox AsteroidMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   1620
      Picture         =   "Form1.frx":6D530
      ScaleHeight     =   555
      ScaleWidth      =   750
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox Asteroid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   840
      Picture         =   "Form1.frx":6EB6A
      ScaleHeight     =   555
      ScaleWidth      =   750
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox Asteroid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   840
      Picture         =   "Form1.frx":701A4
      ScaleHeight     =   555
      ScaleWidth      =   750
      TabIndex        =   5
      Top             =   4380
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox Asteroid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   840
      Picture         =   "Form1.frx":717DE
      ScaleHeight     =   555
      ScaleWidth      =   750
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox Asteroid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   840
      Picture         =   "Form1.frx":72E18
      ScaleHeight     =   555
      ScaleWidth      =   750
      TabIndex        =   3
      Top             =   3060
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox ShipMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   8040
      Picture         =   "Form1.frx":74452
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   3615
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox LeftTurnMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Index           =   5
      Left            =   1680
      Picture         =   "Form1.frx":7C324
      ScaleHeight     =   1350
      ScaleWidth      =   1800
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox Ship 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   6120
      Picture         =   "Form1.frx":841F6
      ScaleHeight     =   1350
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   11895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   2820
      Visible         =   0   'False
      Width           =   11925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cheata As Boolean
Dim Exitting As Boolean
Dim Shoping As Boolean
Dim Pause As Boolean
Dim XDest As Integer
Dim YDest As Integer
Dim Healthlost, Shooting As Boolean
Dim j
Dim X, Y As Integer
Dim MError As Boolean
Dim Speed1
Dim Counter As Long
Dim Level As Long
Dim Cheat
Dim Cannon As Long
Dim Frame As Integer
Dim KeyLeft As Boolean '' 37
Dim KeyRight As Boolean '' 39
Dim KeyUp As Boolean '' 38
Dim KeyDown As Boolean '' 40
Dim NukeLaunch As Boolean
Dim ASpeed

Private Sub Command1_Click()
If Shops.Digits < 1 Then
Msg2 "No shoping credits left", 1
Exit Sub
End If
Shops.Digits = Shops.Digits - 1
Shoping = True
Me.Hide
Pause = True
Form4.Load
Do Until Pause = False
DoEvents
Loop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
NukeLaunch = True
End If
If KeyCode = 37 Then
KeyLeft = True
End If
If KeyCode = 39 Then
KeyRight = True
End If
If KeyCode = 38 Then
KeyUp = True
End If
If KeyCode = 40 Then
KeyDown = True
End If
If KeyCode = 17 Then
Shooting = True
End If
Cheat = Cheat & Chr(KeyCode)
If Format(Cheat, "<") = "egbert" Then
Score = Score + Int(Rnd * 3000000)
Cheat = ""
End If
If Format(Cheat, "<") = "newlive" Then
Cheata = True
End If
Select Case KeyCode
Case vbKeyDelete
Cheat = ""
Case vbKeyEscape
j = Msg2("Are you sure you want to leave this game?", 2)
If j = 1 Then
j = sndPlaySound(vbNullString, SND_ASYNC)
Form2.Show
Me.Hide
Dim typDevM As typDevMODE
Dim lngResult As Long
Dim intAns    As Integer
lngResult = EnumDisplaySettings(0, 0, typDevM)
With typDevM
.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
.dmPelsWidth = Form2.Res1  'ScreenWidth (640,800,1024, etc)
.dmPelsHeight = Form2.Res2 'ScreenHeight (480,600,768, etc)
End With
lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
Select Case lngResult
Case DISP_CHANGE_RESTART
intAns = MsgBox("You must restart your computer to apply these changes." & vbCrLf & vbCrLf & "Do you want to restart now?", vbYesNo + vbSystemModal, "Screen Resolution")
If intAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
Case DISP_CHANGE_SUCCESSFUL
Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
Case Else
MsgBox "Display mode error can not switch to 800x600 game can not be loaded.", vbSystemModal, "Error"
End Select
Exitting = True
End If
Case vbKeyPause
If Pause = False Then
Pause = True
Label3.Visible = True
Else
Pause = False
Label3.Visible = False
End If
Do Until Pause = False
DoEvents
Loop
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
KeyLeft = False
End If
If KeyCode = 39 Then
KeyRight = False
End If
If KeyCode = 38 Then
KeyUp = False
End If
If KeyCode = 40 Then
KeyDown = False
End If
If KeyCode = 17 Then
Shooting = False
End If
End Sub

Private Sub Form_Load()
Me.Hide
End Sub

Function Load()
On Error Resume Next
Lasers.Digits = 0
ASpeed = 0
Me.Shops.Digits = 1
Me.Score = 0
Form4.Price(0).Digits = 8000
Form4.Price(1).Digits = 6000
Form4.Price(2).Digits = 200000
Form4.Price(3).Digits = 3000
KeyLeft = False
KeyUp = False
KeyLeft = False
KeyRight = False
Me.Refresh
Frame1.Width = Me.ScaleWidth - 1
Frame1.Top = Form1.ScaleHeight - 64
Image1.Width = Me.Width - 15
Me.Refresh
Label3.Visible = False
Shooting = False
Form4.Reset
Cannon = 1
Score = 0
Speed1 = 5
Me.Show
Cheata = False
Healtm.Width = 2655
Dim I, A, B, S
Dim XX, YY, XXX, YYY, WW, WWW As Long
Dim Counter4 As Long
Dim Rockvisible(1 To 10) As Boolean
Dim LaserX(1 To 100) As Integer
Dim LaserY(1 To 100) As Integer
Dim LaserX2(1 To 100) As Integer
Dim LaserY2(1 To 100) As Integer
Dim CountExp(1 To 100) As Long
Dim Shootcount As Long
Dim Hitcountrock(1 To 10) As Long
Dim ShootLaser(1 To 100) As Boolean
Dim ShootLaser2(1 To 100) As Boolean
Dim RockNum(1 To 10) As Integer
Dim RockXPos(1 To 10) As Integer
Dim RockYPos(1 To 10) As Integer
Dim ShipXPos(1 To 10) As Integer
Dim Laser1(1 To 100) As Long
Dim ShipYPos(1 To 10) As Integer
Dim ShipExplode(1 To 10) As Long
Dim ShipVisible(1 To 10) As Boolean
Dim Counter1(1 To 10) As Long
Dim Explode1(1 To 100) As Boolean
Dim ExpplaceX1(1 To 100) As Long
Dim ExpplaceY1(1 To 100) As Long
Dim ExpplaceX2(1 To 100) As Long
Dim ExpplaceY2(1 To 100) As Long
Dim ExpplaceX3(1 To 100) As Long
Dim ExpplaceY3(1 To 100) As Long
Dim ExpplaceX4(1 To 100) As Long
Dim ExpplaceY4(1 To 100) As Long
Dim BonusVisible(1 To 5) As Boolean
Dim BonusX(1 To 5) As Integer
Dim BonusY(1 To 5) As Integer
Dim BonusHitCount(1 To 5) As Long
Dim BonusPicture(1 To 5) As Long
Dim Color1 As Integer
Dim Wait As Long
Dim XLoader As Integer
Dim Wait2 As Long
Dim YLoader As Integer
Dim Startcolor(1 To 6000) As Integer
Dim StarXPos(1 To 6000)
Dim StarYPos(1 To 6000)
Dim StarPosSpeed(1 To 6000) As Long
Dim RandX As Integer
Dim currentchannel
Dim Many As Long
Dim typDevM As typDevMODE
Dim lngResult As Long
Dim intAns As Integer
Dim JK As Long
Dim K As Integer
Dim SpeedUp As Integer
Dim Speed3 As Long
Dim Nuke As Long
Dim NukeX(1 To 2) As Integer
Dim NukeY(1 To 2) As Integer
Dim NukeVisible(1 To 2) As Boolean
Dim NukeSpeed(1 To 2) As Integer
Dim NukeExplode(1 To 2) As Boolean
Dim Nukeexplode2(1 To 2) As Integer
SpeedUp = 8 ''Speed from moving
NukeLaunch = False
Nuke = 0
For I = 1 To 600
StarYPos(I) = I
Color1 = Int(Rnd * 270)
Startcolor(I) = Color1
StarXPos(I) = Int(Rnd * Me.ScaleWidth + 1)
StarPosSpeed(I) = Int((10 * Rnd) + 3)
Next I

Exitting = False
Counter = 990
Wait = 0
Level = 1
Counter4 = 2000

If Form5.Check2.Value = 1 Then
If SetupSound(Me.hWnd, 302) = False Then End   ' never mind...
End If

Dim Filename As String
If Form5.Check1.Value = 1 Then
Filename = App.Path & "\Sound\" & "Needle.wav"
j = sndPlaySound(Filename, &H1 Or &H8)
End If
Show
For I = 1 To 10
RockYPos(I) = Int(Rnd * 425 + 1)
Rockvisible(I) = True
Next I
X = Me.ScaleWidth / 2
Y = Me.ScaleHeight / 2
Label4.Top = Me.ScaleHeight + 30
Do

For I = 1 To ASpeed
Next I

If Shoping = False Then
If KeyLeft = True Then
If X > 40 Then X = X - SpeedUp
End If
If KeyRight = True Then
If X < Me.ScaleWidth - 40 Then X = X + SpeedUp
End If

If KeyUp = True Then
If Y > 40 Then Y = Y - SpeedUp
End If

If KeyDown = True Then
If Y < Me.ScaleHeight - 90 Then Y = Y + SpeedUp
End If

If Y > Me.ScaleHeight - 90 Then Y = Me.ScaleHeight - 90
YDest = Y - Ship.Height / 2
End If

If X > 10 And X < Me.ScaleWidth Then
XDest = X - Ship.Width / 2
End If


Frame = Frame + 1
DoEvents
If Form5.Check3.Value = 1 Then
Do Until Wait2 > Form5.S2.Value
Wait2 = Wait2 + 1
DoEvents
Loop
Wait2 = 0
End If
Do Until Counter < 999
DoEvents
If Counter > 1000 Then
Label4 = "Level " & Level
Label4.Visible = True
Label4.Top = Label4.Top - 0.5
If Val(Label4.Top) < Val(Me.ScaleHeight / 2 - 100) Then
Counter = 0
Level = Level + 1
If Level > 2 Then
For I = 1 To 10
If ShipVisible(I) = False Then
ShipVisible(I) = True
ShipXPos(I) = Int(Rnd * Me.ScaleWidth - 60)
ShipYPos(I) = Me.ScaleHeight
End If
Next I
End If
End If
''Counter = 0
''152 stop
''400 start
'' -100 invisible
Else
Counter = Counter + 1
End If
Loop
Counter = Counter + 1

Do Until Wait < 90
DoEvents
If Label4.Visible = False Then Exit Do: GoTo 15
If Wait > 90 Then
If Label4.Top < -10 Then
Label4.Visible = False
Label4.Top = Me.ScaleHeight
Wait = 0
Else
Label4.Top = Label4.Top - 0.5
End If
Else
Wait = Wait + 1
End If
Loop
15:
If Label4.Visible = True Then
If Wait < 90 Then
Wait = Wait + 1
End If
End If

Cls

For I = 1 To Nuke '' nuke  launcher
If NukeVisible(I) = False Then
If I = 1 Then
j = BitBlt(Me.hDC, XDest + 18, YDest + 15, 150, 112, Nukemask2.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, XDest + 18, YDest + 15, 150, 112, Nuke2.hDC, 0, 0, vbSrcPaint)
NukeX(I) = XDest + 18
NukeY(I) = YDest + 15
Else
j = BitBlt(Me.hDC, XDest + Ship.Width - 23, YDest + 15, 150, 112, Nukemask2.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, XDest + Ship.Width - 23, YDest + 15, 150, 112, Nuke2.hDC, 0, 0, vbSrcPaint)
NukeX(I) = XDest + Ship.Width - 23
NukeY(I) = YDest + 15
End If
End If
Next I

If NukeLaunch = True Then
If Nuke = 2 Then
NukeVisible(2) = True
NukeLaunch = False
End If
If Nuke = 1 Then
NukeVisible(1) = True
NukeLaunch = False
End If
Nuke = Nuke - 1
End If

For I = 1 To 2
If NukeExplode(I) = True Then
Me.BackColor = RGB(Nukeexplode2(I), Nukeexplode2(I), Nukeexplode2(I))
Nukeexplode2(I) = Nukeexplode2(I) - 5
If Nukeexplode2(I) < 90 Then
Nukeexplode2(I) = 300
NukeExplode(I) = False
NukeVisible(I) = False
Me.BackColor = vbBlack
For K = 1 To 10
Rockvisible(K) = False
For S = 1 To 100
If Explode1(S) = False Then
Explode1(S) = True
ExpplaceX1(S) = RockXPos(K)
ExpplaceY1(S) = RockYPos(K)
ExpplaceY2(S) = RockYPos(K) + 16
ExpplaceX2(S) = RockXPos(K)
ExpplaceX3(S) = RockXPos(K) + 24
ExpplaceY3(S) = RockYPos(K) + 16
ExpplaceX4(S) = RockXPos(K) + 24
ExpplaceY4(S) = RockYPos(K)
Exit For
End If
Next S
PlaySound 2, ExpplaceX1(K)
ShipExplode(K) = 1
ShipVisible(K) = False
Next K
End If
PlaySound 2, Me.ScaleWidth / 2
Score = Score + 4000
End If
If NukeVisible(I) = True Then
j = BitBlt(Me.hDC, NukeX(I), NukeY(I), 150, 112, Nukemask.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, NukeX(I), NukeY(I), 150, 112, Nuke1.hDC, 0, 0, vbSrcPaint)
NukeY(I) = NukeY(I) + NukeSpeed(I)
If NukeSpeed(I) < 5 Then NukeSpeed(I) = NukeSpeed(I) + 1
If NukeY(I) > 800 Then NukeVisible(I) = False: NukeExplode(I) = True: Nukeexplode2(I) = 300

End If
Next I

For I = 1 To 5
If BonusVisible(I) = True Then
j = BitBlt(Me.hDC, BonusX(I), BonusY(I), 150, 112, Bonusmask.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, BonusX(I), BonusY(I), 150, 112, Bonus(BonusPicture(I)).hDC, 0, 0, vbSrcPaint)
BonusY(I) = BonusY(I) - 3
If BonusY(I) < 0 Then BonusVisible(I) = False: Exit For
XX = BonusX(I)
YY = BonusY(I)
YYY = YDest + 20 - Ship.Height / 2
XXX = XDest + 35 - Ship.Width / 2
WW = Ship.Width - 10
WWW = Ship.Height
If XX > XXX And XX < XXX + WW And YY > YYY And YY < YYY + WWW Then
BonusVisible(I) = False

If BonusPicture(I) = 0 Then
Score = Score + 10000 ''Plus 10000 score
PlaySound 8, Me.ScaleWidth / 2
End If

If BonusPicture(I) = 6 Then
If Nuke < 2 Then Nuke = Nuke + 1 ''Plus 1 Nuke
PlaySound 13, Me.ScaleWidth / 2
End If

If BonusPicture(I) = 7 Then
SpeedUp = SpeedUp + 2 ''Plus 1 speed
PlaySound 12, Me.ScaleWidth / 2
End If

If BonusPicture(I) = 1 Then
Lasers.Digits = Lasers.Digits + 600 ''plus 600 ammo
PlaySound 9, Me.ScaleWidth / 2
End If

If BonusPicture(I) = 1 Then
Lasers.Digits = Lasers.Digits + 600 ''plus 600 ammo
PlaySound 9, Me.ScaleWidth / 2
End If

If BonusPicture(I) = 2 Then
Healtm.Width = 2655 ''Full shield
PlaySound 10, Me.ScaleWidth / 2
End If

If BonusPicture(I) = 3 Then
PlaySound 8, Me.ScaleWidth / 2
Score = Score + 20000 ''plus 20000 score
End If

If BonusPicture(I) = 4 Then
Shops.Digits = Shops.Digits + 1 ''goto shop
PlaySound 11, Me.ScaleWidth / 2
End If

If BonusPicture(I) = 5 Then
Score = Score + 30000 ''plus 30000 score
PlaySound 8, Me.ScaleWidth / 2
End If

BonusPicture(I) = 0
End If
End If
Next I


If Shooting = True Then
If Speed3 > Speed1 Then Else Speed3 = Speed3 + 1
If Speed3 > Speed1 Then
Speed3 = 0
Shootcount = Shootcount + 1
For I = 1 To 100
If ShootLaser(I) = False Then
ShootLaser(I) = True
Laser1(I) = 0
LaserX(I) = XDest
LaserY(I) = YDest + Ship.Height - 20
PlaySound 1, XDest
Exit For
End If
Next I
If Cannon > 1 Then
For I = 1 To 100
If ShootLaser(I) = False Then
If Lasers.Digits = 0 Then Exit For Else Lasers.Digits = Lasers.Digits - 1
ShootLaser(I) = True
Laser1(I) = 2
LaserX(I) = XDest - 24
LaserY(I) = YDest + Ship.Height - 45
PlaySound 1, LaserY(I)
Exit For
End If 'klkllklk
Next I
End If
If Cannon > 2 Then
For I = 1 To 100
If ShootLaser(I) = False Then
If Lasers.Digits = 0 Then Exit For Else Lasers.Digits = Lasers.Digits - 1
ShootLaser(I) = True
Laser1(I) = 1
LaserX(I) = XDest + 24
LaserY(I) = YDest + Ship.Height - 45
PlaySound 1, LaserY(I)
Exit For
End If
Next I
End If
End If
End If

If Healthlost = True Then
A = A - 5
B = B + 5
Healtm.FillColor = RGB(A, B, 0)
If B > 220 Then Healthlost = False
End If

If Exitting = True Then
GoTo 13
End If

For I = 1 To 600
StarYPos(I) = StarYPos(I) - StarPosSpeed(I)
If StarYPos(I) < 0 Then StarYPos(I) = 600: StarPosSpeed(I) = Int((10 * Rnd) + 3)
Color1 = Startcolor(I)
PSet (StarXPos(I), StarYPos(I)), RGB(Color1, Color1, Color1)
Next I


j = BitBlt(Me.hDC, XDest, YDest, 150, 112, ShipMask.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, XDest, YDest, 150, 112, Ship.hDC, 0, 0, vbSrcPaint)

If Level = 1 Then JK = 100
If Level = 2 Then JK = 80
If Level = 3 Then JK = 60
If Level = 4 Then JK = 40
If Level = 5 Then JK = 20
If Level = 6 Then JK = 10
If Level > 7 Then JK = 5

If Int(Rnd * JK) = 3 Then
For I = 1 To 10
If ShipVisible(I) = False Then
ShipVisible(I) = True
ShipXPos(I) = Int(Rnd * Me.ScaleWidth - 60)
ShipYPos(I) = Me.ScaleHeight
Exit For
End If
Next I
End If

For I = 1 To 100
If ShootLaser2(I) = True Then 'If thehas been activated then:
j = BitBlt(Me.hDC, LaserX2(I) + 55, LaserY2(I), 150, 112, LaserMask.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, LaserX2(I) + 55, LaserY2(I), 150, 112, Laser(3).hDC, 0, 0, vbSrcPaint)
LaserY2(I) = LaserY2(I) - 20
If LaserY2(I) < 0 Then
ShootLaser2(I) = False
End If
End If
If ShootLaser(I) = True Then 'If the laser has been activated then:
j = BitBlt(Me.hDC, LaserX(I) + 55, LaserY(I), 150, 112, LaserMask.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, LaserX(I) + 55, LaserY(I), 150, 112, Laser(Laser1(I)).hDC, 0, 0, vbSrcPaint)
LaserY(I) = LaserY(I) + 23
If LaserY(I) >= Me.ScaleHeight Then
ShootLaser(I) = False
End If
End If
If Explode1(I) = True Then
CountExp(I) = CountExp(I) + 1
If CountExp(I) > Level + 5 Then
Explode1(I) = False
CountExp(I) = 0
End If
ExpplaceX1(I) = ExpplaceX1(I) - 10
ExpplaceY1(I) = ExpplaceY1(I) - 10
ExpplaceY2(I) = ExpplaceY2(I) + 10
ExpplaceX2(I) = ExpplaceX2(I) - 10
ExpplaceX3(I) = ExpplaceX3(I) + 10
ExpplaceY3(I) = ExpplaceY3(I) + 10
ExpplaceX4(I) = ExpplaceX4(I) + 10
ExpplaceY4(I) = ExpplaceY4(I) - 10
j = BitBlt(Me.hDC, ExpplaceX1(I), ExpplaceY1(I), 150, 112, ExpM(0).hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, ExpplaceX1(I), ExpplaceY1(I), 150, 112, Exp(0).hDC, 0, 0, vbSrcPaint)
j = BitBlt(Me.hDC, ExpplaceX2(I), ExpplaceY2(I), 150, 112, ExpM(1).hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, ExpplaceX2(I), ExpplaceY2(I), 150, 112, Exp(1).hDC, 0, 0, vbSrcPaint)
j = BitBlt(Me.hDC, ExpplaceX3(I), ExpplaceY3(I), 150, 112, ExpM(2).hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, ExpplaceX3(I), ExpplaceY3(I), 150, 112, Exp(2).hDC, 0, 0, vbSrcPaint)
j = BitBlt(Me.hDC, ExpplaceX4(I), ExpplaceY4(I), 150, 112, ExpM(3).hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, ExpplaceX4(I), ExpplaceY4(I), 150, 112, Exp(3).hDC, 0, 0, vbSrcPaint)
End If

If Explode1(I) = True Then
If ExpplaceX1(I) > XDest And ExpplaceX1(I) < XDest + ShipMask.ScaleHeight And ExpplaceY1(I) > YDest And ExpplaceY1(I) < YDest + ShipMask.ScaleWidth Then GoSub Healta
If ExpplaceX2(I) > XDest And ExpplaceX2(I) < XDest + ShipMask.ScaleHeight And ExpplaceY2(I) > YDest And ExpplaceY2(I) < YDest + ShipMask.ScaleWidth Then GoSub Healta
If ExpplaceX3(I) > XDest And ExpplaceX3(I) < XDest + ShipMask.ScaleHeight And ExpplaceY3(I) > YDest And ExpplaceY3(I) < YDest + ShipMask.ScaleWidth Then GoSub Healta
If ExpplaceX4(I) > XDest And ExpplaceX4(I) < XDest + ShipMask.ScaleHeight And ExpplaceY4(I) > YDest And ExpplaceY4(I) < YDest + ShipMask.ScaleWidth Then GoSub Healta
End If

Next I

For I = 1 To 10
If ShipExplode(I) = 0 Then GoTo 98
A = ShipExplode(I)
j = BitBlt(Me.hDC, ShipXPos(I), ShipYPos(I), 150, 112, Picture1(A).hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, ShipXPos(I), ShipYPos(I), 150, 112, Picture2(A).hDC, 0, 0, vbSrcPaint)
If ShipExplode(I) > 7 Then
ShipExplode(I) = 0
Else
Counter1(I) = Counter1(I) + 1
If Counter1(I) > 2 Then
Counter1(I) = 0
ShipExplode(I) = ShipExplode(I) + 1
End If
End If
98:
If ShipVisible(I) = True Then
ShipYPos(I) = ShipYPos(I) - 5
If XDest > ShipXPos(I) - 55 Then
ShipXPos(I) = ShipXPos(I) + 3
End If
If XDest < ShipXPos(I) - 55 Then
ShipXPos(I) = ShipXPos(I) - 3
End If
If ShipYPos(I) < -100 Then ShipVisible(I) = False
If Level = 1 Then JK = 100
If Level = 2 Then JK = 80
If Level = 3 Then JK = 60
If Level = 4 Then JK = 40
If Level = 5 Then JK = 20
If Level = 6 Then JK = 10
If Level > 7 Then JK = 5
If Int(Rnd * JK) = 3 Then ''jk
For S = 1 To 100
If ShootLaser2(S) = False Then
LaserX2(S) = ShipXPos(I) - 50
LaserY2(S) = ShipYPos(I) - 35 'Set the right pos
PlaySound 5, LaserY2(S)
ShootLaser2(S) = True
Exit For
End If
Next S
End If
j = BitBlt(Me.hDC, ShipXPos(I), ShipYPos(I), 150, 112, Shipam.hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, ShipXPos(I), ShipYPos(I), 150, 112, Shipa.hDC, 0, 0, vbSrcPaint)
End If
j = BitBlt(Me.hDC, RockXPos(I), RockYPos(I), 150, 112, AsteroidMask(RockNum(I)).hDC, 0, 0, vbSrcAnd)
j = BitBlt(Me.hDC, RockXPos(I), RockYPos(I), 150, 112, Asteroid(RockNum(I)).hDC, 0, 0, vbSrcPaint)
If Rockvisible(I) = False Then GoTo 1
If RockYPos(I) + 100 <= 0 Then
1:
RockYPos(I) = Me.ScaleHeight
RockXPos(I) = Int(Rnd * Me.ScaleWidth)
RockNum(I) = Int(Rnd * 4)
Rockvisible(I) = True
End If
2:
If Rockvisible(I) = True Then
RockYPos(I) = RockYPos(I) - 5
XX = RockXPos(I)
YY = RockYPos(I)
YYY = YDest + 20 - Ship.Height / 2
XXX = XDest + 35 - Ship.Width / 2
WW = Ship.Width - 10 'Laser(0).Heightdfsdsffdsifdsfhfuihdsiushiusghsduysgsuysgsdusdgsiusdfddgudsdfgusdsdgkjdsgs
WWW = Ship.Height  'Laser(0).Width
If XX > XXX And XX < XXX + WW And YY > YYY And YY < YYY + WWW Then
Rockvisible(I) = False
For S = 1 To 100
If Explode1(S) = False Then
Explode1(S) = True
ExpplaceX1(S) = RockXPos(I)
ExpplaceY1(S) = RockYPos(I)
ExpplaceY2(S) = RockYPos(I) + 16
ExpplaceX2(S) = RockXPos(I)
ExpplaceX3(S) = RockXPos(I) + 24
ExpplaceY3(S) = RockYPos(I) + 16
ExpplaceX4(S) = RockXPos(I) + 24
ExpplaceY4(S) = RockYPos(I)
Exit For
End If
Next S
PlaySound 2, XDest
GoSub Health
End If
End If
If ShipVisible(I) = True Then
XX = ShipXPos(I)
YY = ShipYPos(I)
YYY = YDest + 20 - Ship.Height / 2
XXX = XDest + 35 - Ship.Width / 2
WW = Ship.Width - 10 'Laser(0).Heightdfsdsffdsifdsfhfuihdsiushiusghsduysgsuysgsdusdgsiusdfddgudsdfgusdsdgkjdsgs
WWW = Ship.Height  'Laser(0).Width
If XX > XXX And XX < XXX + WW And YY > YYY And YY < YYY + WWW Then
GoSub Healtc
ShipVisible(I) = False
ShipExplode(I) = 1
PlaySound 4, ShipYPos(I)
End If
End If
For A = 1 To 100
If ShootLaser2(A) = True Then
XX = LaserX2(A)
YY = LaserY2(A)
YYY = YDest + 20 - Ship.Height / 2
XXX = XDest + 35 - Ship.Width / 2
WW = Ship.Width - 10 'Laser(0).Heightdfsdsffdsifdsfhfuihdsiushiusghsduysgsuysgsdusdgsiusdfddgudsdfgusdsdgkjdsgs
WWW = Ship.Height  'Laser(0).Width
If XX > XXX And XX < XXX + WW And YY > YYY And YY < YYY + WWW Then
PlaySound 6, LaserY2(A)
ShootLaser2(A) = False: GoSub Healtb
End If
End If
If ShootLaser(A) = True Then

If I < 6 Then
If BonusVisible(I) = False Then GoTo jKl
If BonusX(I) > LaserX(A) And BonusX(I) < LaserX(A) + Bonusmask.ScaleWidth And BonusY(I) > LaserY(A) And BonusY(I) < LaserY(A) + Bonusmask.ScaleHeight + 30 Then Else GoTo jKl
If BonusHitCount(I) > Level - 1 * Cannon Then
BonusHitCount(I) = 0
If SpeedUp > 19 Then K = 6 Else K = 7
If BonusPicture(I) = K Then Else BonusPicture(I) = BonusPicture(I) + 1
Else
BonusHitCount(I) = BonusHitCount(I) + 1
End If
ShootLaser(A) = False
PlaySound 3, BonusX(I)
End If

jKl:

If ShipVisible(I) = True Then
If ShipXPos(I) > LaserX(A) And ShipXPos(I) < LaserX(A) + ShipMask.ScaleWidth And ShipYPos(I) > LaserY(A) And ShipYPos(I) < LaserY(A) + ShipMask.ScaleHeight Then
If Int((7 * Rnd) + 1) = 5 Then
For K = 1 To 5
If BonusVisible(K) = False Then
BonusVisible(K) = True
BonusX(K) = ShipXPos(I)
BonusY(K) = ShipYPos(I)
Exit For
End If
Next K
End If

ShipExplode(I) = 1
ShipVisible(I) = False
ShootLaser(A) = False
PlaySound 4, LaserX(A)
Many = Level * 2500
Score = Score + Int(Rnd * Many)
End If
End If
If RockXPos(I) > LaserX(A) And RockXPos(I) < LaserX(A) + AsteroidMask(1).ScaleWidth And RockYPos(I) > LaserY(A) And RockYPos(I) < LaserY(A) + AsteroidMask(1).ScaleHeight Then
Hitcountrock(I) = Hitcountrock(I) + 1
ShootLaser(A) = False
Many = Level * 10
Score = Score + Int(Rnd * Many)
If Hitcountrock(I) > Level - 1 Then
Hitcountrock(I) = 0
Rockvisible(I) = False
For S = 1 To 100
If Explode1(S) = False Then
Explode1(S) = True
ExpplaceX1(S) = RockXPos(I)
ExpplaceY1(S) = RockYPos(I)
ExpplaceY2(S) = RockYPos(I) + 16
ExpplaceX2(S) = RockXPos(I)
ExpplaceX3(S) = RockXPos(I) + 24
ExpplaceY3(S) = RockYPos(I) + 16
ExpplaceX4(S) = RockXPos(I) + 24
ExpplaceY4(S) = RockYPos(I)
Exit For
End If
Next S
PlaySound 2, RockYPos(I)
Many = Level * 200
Score = Score + Int(Rnd * Many)
End If
End If
End If
Next A
Next I
Shield.Caption = "Shield : " & Int(Healtm.Width / 2655 * 100) & "%"
Loop

Healta:
If Cheata = True Then Return
Healtm.FillColor = RGB(227, 0, 0)
Healtm.Width = Healtm.Width - 5
If Healtm.Width < 20 Then
MError = True
GoTo 12
End If
Healthlost = True
A = 227
B = 0
Return

Healtb:
If Cheata = True Then Return
Healtm.FillColor = RGB(227, 0, 0)
Healtm.Width = Healtm.Width - 9
If Healtm.Width < 20 Then
MError = True
GoTo 12
End If
Healthlost = True
A = 227
B = 0
Return

Healtc:
If Cheata = True Then Return
Healtm.FillColor = RGB(227, 0, 0)
Healtm.Width = Healtm.Width - 70
If Healtm.Width < 20 Then
MError = True
GoTo 12
End If
Healthlost = True
A = 227
B = 0
Return

Health:
If Cheata = True Then Return
Healtm.FillColor = RGB(227, 0, 0)
Healtm.Width = Healtm.Width - 30
If Healtm.Width < 20 Then
MError = True
GoTo 12
End If
Healthlost = True
A = 227
B = 0
Return
12:
j = sndPlaySound(vbNullString, SND_ASYNC)
lngResult = EnumDisplaySettings(0, 0, typDevM)
With typDevM
.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
.dmPelsWidth = Form2.Res1  'ScreenWidth (640,800,1024, etc)
.dmPelsHeight = Form2.Res2 'ScreenHeight (480,600,768, etc)
End With
lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
Select Case lngResult
Case DISP_CHANGE_RESTART
intAns = MsgBox("You must restart your computer to apply these changes." & vbCrLf & vbCrLf & "Do you want to restart now?", vbYesNo + vbSystemModal, "Screen Resolution")
If intAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
Case DISP_CHANGE_SUCCESSFUL
Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
Case Else
MsgBox "Display mode error can not switch to 800x600 game can not be loaded.", vbSystemModal, "Error"
End Select
ScoreList.Himade Score
13:
j = sndPlaySound(vbNullString, SND_ASYNC)
lngResult = EnumDisplaySettings(0, 0, typDevM)
With typDevM
.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
.dmPelsWidth = Form2.Res1  'ScreenWidth (640,800,1024, etc)
.dmPelsHeight = Form2.Res2 'ScreenHeight (480,600,768, etc)
End With
lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
Select Case lngResult
Case DISP_CHANGE_RESTART
intAns = MsgBox("You must restart your computer to apply these changes." & vbCrLf & vbCrLf & "Do you want to restart now?", vbYesNo + vbSystemModal, "Screen Resolution")
If intAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
Case DISP_CHANGE_SUCCESSFUL
Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
Case Else
MsgBox "Display mode error can not switch to 800x600 game can not be loaded.", vbSystemModal, "Error"
End Select 'Unload Me
End Function

Function PlaySound(Number1 As Long, Place)
If Form5.Check2.Value = 1 Then
Dim Pan As Long
Dim Ret As Integer
Dim Filename As String

If Form5.Check5.Value = 1 Then
Pan = Int(Place / Me.ScaleWidth * 100)
Else
Pan = 50
End If

Select Case Number1
Case 1
Filename = App.Path & "\Sound\" & "Shot.wav"
PlaySoundByBuffer 0, Filename, 95, Pan, 0
Case 2
Filename = App.Path & "\Sound\" & "Explode.wav"
Ret = PlaySoundGiveBuffer(Filename, 95, Pan, 0, 10, 100)
Case 3
Filename = App.Path & "\Sound\" & "hit.wav"
PlaySoundByBuffer 1, Filename, 90, Pan, 0
Case 4
Filename = App.Path & "\Sound\" & "Explode2.wav"
Ret = PlaySoundGiveBuffer(Filename, 95, Pan, 0, 101, 201)
Case 5
Filename = App.Path & "\Sound\" & "Shot2.wav"
PlaySoundByBuffer 2, Filename, 100, Pan, 0
Case 6
Filename = App.Path & "\Sound\" & "Explode2.wav"
Ret = PlaySoundGiveBuffer(Filename, 95, Pan, 0, 202, 302)
Case 7
Filename = App.Path & "\Sound\" & "main.wav"
PlaySoundByBuffer 3, Filename, 100, Pan, 0
Case 8
Filename = App.Path & "\Sound\" & "credits.wav"
PlaySoundByBuffer 4, Filename, 100, Pan, 0
Case 9
Filename = App.Path & "\Sound\" & "ammo.wav"
PlaySoundByBuffer 5, Filename, 100, Pan, 0
Case 10
Filename = App.Path & "\Sound\" & "shield.wav"
PlaySoundByBuffer 6, Filename, 100, Pan, 0
Case 11
Filename = App.Path & "\Sound\" & "shop.wav"
PlaySoundByBuffer 7, Filename, 100, Pan, 0
Case 12
Filename = App.Path & "\Sound\" & "SpeedUp.wav"
PlaySoundByBuffer 8, Filename, 100, Pan, 0
Case 13
Filename = App.Path & "\Sound\" & "nuke.wav"
PlaySoundByBuffer 9, Filename, 100, Pan, 0
End Select
End If
End Function

Function GetSpeed()
GetSpeed = Speed1
End Function

Function Getsheeld() As Long
Getsheeld = Healtm.Width
End Function

Function Setsheeld(Newsheeld As Long)
Healtm.Width = Newsheeld
End Function

Function SetSpeed(Newspeed As Long)
Speed1 = Newspeed
End Function

Function SetPause(Pause1 As Boolean)
Pause = Pause1
End Function

Function Getlevel() As Long
Getlevel = Level
End Function

Function GetCannons() As Long
GetCannons = Cannon
End Function

Function SetCondition(New1 As Long)
NukeLaunch = False
KeyLeft = False
KeyUp = False
KeyLeft = False
KeyRight = False
Shooting = False
Lasers.Digits = New1
End Function

Function GetCondition() As Long
GetCondition = Lasers.Digits
End Function

Function SetCannons(New1 As Long)
Shoping = False
KeyLeft = False
KeyUp = False
KeyLeft = False
KeyRight = False
X = Me.ScaleWidth / 2
Y = Me.ScaleHeight / 2
Cannon = New1
End Function

Private Sub Form_Unload(Cancel As Integer)
j = sndPlaySound(vbNullString, SND_ASYNC)
End Sub

Private Sub Score_Change()
Score1.Digits = Score.Caption
End Sub

Private Sub Timer1_Timer()
Frame11.Digits = Frame
Frame = 0

If Shoping = True Then Exit Sub
If Pause = True Then Exit Sub

If Form5.Check4.Value = 1 Then

If Frame11.Digits > 40 Then
ASpeed = ASpeed + 5000
End If

If Frame11.Digits < 30 Then
If ASpeed > 0 Then ASpeed = ASpeed - 5000
End If

If ASpeed < 0 Then ASpeed = 0
End If
End Sub
