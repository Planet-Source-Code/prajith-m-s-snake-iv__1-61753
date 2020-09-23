VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SNAKE-IV"
   ClientHeight    =   5445
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6390
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   2760
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Points:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   225
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   0
      X2              =   0
      Y1              =   600
      Y2              =   5400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6360
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   6360
      X2              =   6360
      Y1              =   5400
      Y2              =   600
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   480
      Shape           =   3  'Circle
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   25
      Left            =   4680
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   24
      Left            =   960
      Top             =   4680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   23
      Left            =   2880
      Top             =   2040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   22
      Left            =   1920
      Top             =   3000
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   21
      Left            =   5880
      Top             =   2520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   20
      Left            =   5520
      Top             =   3600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   19
      Left            =   240
      Top             =   5040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   18
      Left            =   3480
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   17
      Left            =   2520
      Top             =   3600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   16
      Left            =   5040
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   15
      Left            =   960
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   14
      Left            =   4080
      Top             =   3600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   13
      Left            =   480
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   12
      Left            =   1680
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   11
      Left            =   5160
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   10
      Left            =   1680
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   9
      Left            =   3120
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   8
      Left            =   480
      Top             =   2880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   7
      Left            =   5760
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   6
      Left            =   1200
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   5
      Left            =   1200
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   4
      Left            =   2880
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   3
      Left            =   2760
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   2
      Left            =   2640
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   1
      Left            =   2520
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   135
   End
   Begin VB.Menu he 
      Caption         =   "Help"
      Begin VB.Menu tom 
         Caption         =   "To move up - Press up arrow key"
      End
      Begin VB.Menu ml 
         Caption         =   "To move left- Press left arrow key"
      End
      Begin VB.Menu kl 
         Caption         =   "To move right- press right arrow key"
      End
      Begin VB.Menu mu 
         Caption         =   "To move down - Press down arrow key"
      End
   End
   Begin VB.Menu n 
      Caption         =   "&Level"
      Begin VB.Menu l1 
         Caption         =   "Level-1"
      End
      Begin VB.Menu l2 
         Caption         =   "Level-2"
      End
      Begin VB.Menu l3 
         Caption         =   "Level-3"
      End
      Begin VB.Menu l4 
         Caption         =   "Level-4"
      End
      Begin VB.Menu l5 
         Caption         =   "Level-5"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i, k, j, m, p, t, e, x As Integer
Private Sub Form_Activate()
MsgBox "Select a level"
j = 4
i = j
m = 1
Shape1(0).FillColor = QBColor(12)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
k = 56
ElseIf KeyCode = vbKeyDown Then
k = 50
ElseIf KeyCode = vbKeyRight Then
k = 54
ElseIf KeyCode = vbKeyLeft Then
k = 52
End If
End Sub

Private Sub l1_Click()
Timer1.Interval = 100
Timer2.Interval = 1000
l1.Checked = True
l2.Enabled = False
l3.Enabled = False
l4.Enabled = False
l5.Enabled = False
Label3.Caption = "Level-1"
Timer1.Enabled = True
End Sub

Private Sub l2_Click()
Timer1.Interval = 80
Timer2.Interval = 8000
l2.Checked = True
l1.Enabled = False
l3.Enabled = False
l4.Enabled = False
l5.Enabled = False
Label3.Caption = "Level-2"
Timer1.Enabled = True
End Sub

Private Sub l3_Click()
Timer1.Interval = 60
Timer2.Interval = 600
l3.Checked = True
l2.Enabled = False
l1.Enabled = False
l4.Enabled = False
l5.Enabled = False
Label3.Caption = "Level-3"
Timer1.Enabled = True
End Sub

Private Sub l4_Click()
Timer1.Interval = 40
Timer2.Interval = 400
l4.Checked = True
l2.Enabled = False
l3.Enabled = False
l1.Enabled = False
l5.Enabled = False
Label3.Caption = "Level-4"
Timer1.Enabled = True
End Sub

Private Sub l5_Click()
Timer1.Interval = 20
Timer2.Interval = 200
l5.Checked = True
l2.Enabled = False
l3.Enabled = False
l4.Enabled = False
l1.Enabled = False
Label3.Caption = "Level-5"
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If (Shape1(0).Top > 600 And Shape1(0).Top < 5400) And (Shape1(0).Left > 0 And Shape1(0).Left < 6360) Then
If k = 56 Then
Do While i > 0
Shape1(i).Move Shape1(i - 1).Left, Shape1(i - 1).Top
i = i - 1
Loop
If x <> 50 Then
Shape1(0).Move Shape1(0).Left, Shape1(0).Top - Shape1(0).Height + 15
Else
k = 50
Shape1(0).Move Shape1(0).Left, Shape1(0).Top + Shape1(0).Height - 15
End If
ElseIf k = 52 Then
Do While i > 0
Shape1(i).Move Shape1(i - 1).Left, Shape1(i - 1).Top
i = i - 1
Loop
If x <> 54 Then
Shape1(0).Move Shape1(0).Left - Shape1(0).Width + 15, Shape1(0).Top
Else
k = 54
Shape1(0).Move Shape1(0).Left + Shape1(0).Width - 15, Shape1(0).Top
End If
ElseIf k = 54 Then
Do While i > 0
Shape1(i).Move Shape1(i - 1).Left, Shape1(i - 1).Top
i = i - 1
Loop
If x <> 52 Then
Shape1(0).Move Shape1(0).Left + Shape1(0).Width - 15, Shape1(0).Top
Else
k = 52
Shape1(0).Move Shape1(0).Left - Shape1(0).Width + 15, Shape1(0).Top
End If
ElseIf k = 50 Then
Do While i > 0
Shape1(i).Move Shape1(i - 1).Left, Shape1(i - 1).Top
i = i - 1
Loop
If x <> 56 Then
Shape1(0).Move Shape1(0).Left, Shape1(0).Top + Shape1(0).Height - 15
Else
k = 56
Shape1(0).Move Shape1(0).Left, Shape1(0).Top - Shape1(0).Height + 15
End If
End If
Do While m <= j
If (Shape1(0).Top = Shape1(m).Top And Shape1(0).Left = Shape1(m).Left) Then
MsgBox "game over"
If p <= 50 Then
MsgBox " Poor performance"
ElseIf p <= 150 Then
MsgBox "Average Performance"
ElseIf p <= 250 Then
MsgBox " Good performance"
End If
End
Else
m = m + 1
End If
Loop
If Shape1(0).Top = Shape1(j + 1).Top And Shape1(0).Left = Shape1(j + 1).Left Then
j = j + 1
If j Mod 6 = 0 Then
Shape2(e).Visible = True
Timer2.Enabled = True
End If
p = p + 10
Label2.Caption = p
If j = 25 Then
MsgBox " Game over "
MsgBox " Excellent Performance"
End
End If
Shape1(j + 1).Visible = True
End If
If Shape1(0).Top = Shape2(e).Top And Shape1(0).Left = Shape2(e).Left Then
Shape2(e).Visible = False
e = e + 1
If e = 4 Then
e = 0
Timer2.Enabled = False
End If
p = p + 20
Label2.Caption = p
End If
If i = 0 Then
i = j
m = 1
End If
Else
MsgBox "Game over"
If p <= 50 Then
MsgBox " Poor performance"
ElseIf p <= 150 Then
MsgBox "Average Performance"
ElseIf p <= 250 Then
MsgBox " Good performance"
End If

End
End If
x = k
End Sub

Private Sub Timer2_Timer()
t = t + 1
If t = 5 Then
Shape2(e).Visible = False
Timer2.Enabled = False
End If
End Sub
