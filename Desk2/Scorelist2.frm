VERSION 5.00
Begin VB.Form ScoreList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "Scorelist2.frx":0000
   ScaleHeight     =   4800
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "ScoreList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Points1(0 To 10) As Long
Dim Playername1(0 To 10) As String
Dim hin(0 To 10) As String
Dim hisc(0 To 10) As String
Dim Oldy As Long
Dim A, B, C, D, T, U, X, R, I

Function Load()
LoadFile App.Path & "\Data\Hiscore.dat"
Me.Cls
Me.ForeColor = vbYellow
Me.CurrentX = Me.Width / 2 - 1000
Me.FontSize = 15
Me.Print "SCORE LIST"
Me.CurrentX = 0
Me.CurrentY = Me.CurrentY - 200
Me.Print "---------------------------------------------------------------------"
Me.ForeColor = vbRed
Me.CurrentX = 10
Me.FontSize = 10
Oldy = Me.CurrentY
A = 270
For I = 0 To 9
Me.ForeColor = RGB(A, 0, 0)
Me.CurrentY = Me.CurrentY + 200
Me.Print I & ". " & Playername1(I)
A = A - 270 / 25
Next I
Me.ForeColor = vbBlue
Me.CurrentY = Oldy
A = 270
For I = 0 To 9
Me.ForeColor = RGB(0, 0, A)
Me.CurrentX = Me.Width - 1700
Me.CurrentY = Me.CurrentY + 200
Me.Print Points1(I)
A = A - 270 / 25
Next I
End Function

Function LoadFile(Filename As String)
On Error GoTo errorhandler
Open Filename For Input As #1
For C = 0 To 9
Input #1, hin(C), hisc(C)
Next C
For T = 0 To 9
Playername1(T) = hin(T)
Points1(T) = Val(hisc(T))
Next T
Close #1
Exit Function
errorhandler:
U = 0
For D = 0 To 9
'U = U - 1
Playername1(D) = "......."
Points1(D) = U
Next D
Close #1
Open Filename For Output As #1
For C = 0 To 9
Write #1, Playername1(C), Points1(C)
Next C
Close #1
End Function

Function Himade(Score As Long)
LoadFile App.Path & "\Data\Hiscore.dat"
Unload Form1
If Val(Score) > Val(Points1(9)) Then
Highscores Score, App.Path & "\Data\Hiscore.dat"
Else
ScoreList.Load
ShowScoreList.Load
End If
End Function

Function Highscores(Score As Long, Filename As String)
A = 0
Dim P1name$
For X = 0 To 9
If Val(Score) < Val(Points1(X)) Then
Else
If X = 0 Then A = 0: GoTo 1
A = X
GoTo 1
End If
Next X
1:
P1name$ = InputBox("Well done, you made the Hi-Score." & Chr(13) & "Please insert name." & Chr(13) & "(max 20 chrs)", "Hi-Score Name Entry", "The Master")
If P1name$ = "" Then GoTo 2
If Len(P1name$) >= 20 Then GoTo 1
R = 10
Do
R = R - 1
If R <= A Then Exit Do
Playername1(R) = Playername1(R - 1)
Playername1(R - 1) = ""
Points1(R) = Points1(R - 1)
Points1(R - 1) = 0
Loop
Playername1(A) = P1name$
Points1(A) = Score
Open Filename For Output As #1
For C = 0 To 9
Write #1, Playername1(C), Points1(C)
Next C
2:
Close #1
Load
ShowScoreList.Load
End Function

