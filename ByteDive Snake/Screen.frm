VERSION 5.00
Begin VB.Form Screen 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ByteDive Sanke 1.0 (mariodivece@hotmail.com)"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9900
   ForeColor       =   &H8000000E&
   Icon            =   "Screen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2565
      Top             =   4230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2070
      Top             =   4230
   End
   Begin VB.Shape Fruit 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   1125
      Shape           =   1  'Square
      Top             =   990
      Width           =   195
   End
   Begin VB.Shape ColorLed 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   0
      Left            =   1980
      Top             =   2340
      Width           =   195
   End
End
Attribute VB_Name = "Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As Integer 'meaning current
Dim LastMotion As String
Dim KeyPressed As Boolean
Dim Eaten As Integer
Dim Level As Integer

Private Type Position
    PosX As Integer
    PosY As Integer
End Type
Dim Pos() As Position

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyPressed = True Then
    Exit Sub
End If

Select Case KeyAscii
Case 119
    If LastMotion = "D" Then Exit Sub
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY - ColorLed(0).Height
    KeyPressed = True
    LastMotion = "U"
Case 115
    If LastMotion = "U" Then Exit Sub
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY + ColorLed(0).Height
    KeyPressed = True
    LastMotion = "D"
Case 97
    If LastMotion = "R" Then Exit Sub
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX - ColorLed(0).Width
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY
    KeyPressed = True
    LastMotion = "L"
Case 100
    If LastMotion = "L" Then Exit Sub
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX + ColorLed(0).Width
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY
    KeyPressed = True
    LastMotion = "R"
End Select

End Sub

Private Sub Form_Load()
On Local Error GoTo oops
C = 9
LastMotion = "R"
KeyPressed = False
Eaten = 0

ColorLed(0).Left = 1980
ColorLed(0).Top = 2340

Ask:
Level = InputBox("Enter a difficulty level from 1 to 10", "New Game", "5")

If Not IsNumeric(Level) Or Level > 10 Or Level < 1 Then
oops:
    MsgBox "Wrong diffiulty level.", vbCritical, "Error..."
    GoTo Ask
Else
    Timer1.Interval = 300 / Level
End If

Me.Width = ColorLed(0).Width * 30 - 50
Me.Height = ColorLed(0).Height * 20

Fruit.Left = ColorLed(0).Width * 15 + 30
Fruit.Top = ColorLed(0).Height * 10

For i = 1 To 9
    Load ColorLed(i)
    ColorLed(i).Left = ColorLed(i - 1).Left - ColorLed(0).Width
    ColorLed(i).Top = ColorLed(0).Top
    ColorLed(i).Visible = True
Next i

ReDim Pos(9) As Position
With ColorLed(0)
Pos(9).PosX = .Left
Pos(9).PosY = .Top

For i = 0 To 8
    Pos(8 - i).PosX = ColorLed(i).Left
    Pos(8 - i).PosY = ColorLed(i).Top
Next i
End With

Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Local Error GoTo KeepMoving

ColorLed(0).Left = Pos(C).PosX
ColorLed(0).Top = Pos(C).PosY

For i = (UBound(Pos) - ColorLed.Count) To UBound(Pos) - 1
    If Pos(i).PosX = ColorLed(0).Left And Pos(i).PosY = ColorLed(0).Top Then
        Call Lose
        Exit Sub
    End If
Next i

For i = 1 To ColorLed.Count - 1
    ColorLed(i).Left = Pos(C - (ColorLed.Count - i)).PosX
    ColorLed(i).Top = Pos(C - (ColorLed.Count - i)).PosY
Next i

If Fruit.Left = Pos(C).PosX And Fruit.Top = Pos(C).PosY Then
Refruit
Load ColorLed(ColorLed.Count)
ColorLed(ColorLed.Count - 1).Visible = True
    ColorLed(ColorLed.Count - 1).Left = Pos(C - (ColorLed.Count) + 1).PosX
    ColorLed(ColorLed.Count - 1).Top = Pos(C - (ColorLed.Count) + 1).PosY
End If

C = C + 1

KeyPressed = False

Exit Sub

KeepMoving:
Select Case LastMotion
Case "U"
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY - ColorLed(0).Height
    LastMotion = "U"
Case "D"
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY + ColorLed(0).Height
    LastMotion = "D"
Case "L"
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX - ColorLed(0).Width
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY
    LastMotion = "L"
Case "R"
    ReDim Preserve Pos(UBound(Pos) + 1) As Position
    Pos(UBound(Pos)).PosX = Pos(UBound(Pos) - 1).PosX + ColorLed(0).Width
    Pos(UBound(Pos)).PosY = Pos(UBound(Pos) - 1).PosY
    LastMotion = "R"
End Select

If Fruit.Left = Pos(UBound(Pos)).PosX And Fruit.Top = Pos(UBound(Pos)).PosY Then
Refruit
Load ColorLed(ColorLed.Count)
ColorLed(ColorLed.Count - 1).Visible = True
    ColorLed(ColorLed.Count - 1).Left = Pos(UBound(Pos) - (ColorLed.Count)).PosX
    ColorLed(ColorLed.Count - 1).Top = Pos(UBound(Pos) - (ColorLed.Count)).PosY
End If

ColorLed(0).Left = Pos(UBound(Pos)).PosX
ColorLed(0).Top = Pos(UBound(Pos)).PosY

For i = (UBound(Pos) - ColorLed.Count) To UBound(Pos) - 1
    If Pos(i).PosX = ColorLed(0).Left And Pos(i).PosY = ColorLed(0).Top Then
        Call Lose
        Exit Sub
    End If
Next i

For i = 1 To ColorLed.Count - 1
    ColorLed(i).Left = Pos(UBound(Pos) - (ColorLed.Count - i)).PosX
    ColorLed(i).Top = Pos(UBound(Pos) - (ColorLed.Count - i)).PosY
Next i

'Debug.Print "Fruit: " & Fruit.Left & ", " & Fruit.Top
'Debug.Print "Head:  " & Pos(UBound(Pos)).PosX & ", " & Pos(UBound(Pos)).PosY

C = C + 1

KeyPressed = False

End Sub

Private Sub Refruit()
Eaten = Eaten + 1
Beep
Dim X As Integer
Dim Y As Integer

Repeater:
    Fruit.Visible = False
    Randomize Timer
    X = Rnd * 28
    Y = Rnd * 17
    Debug.Print X & ", " & Y
If X = 0 Or Y = 0 Then
    GoTo Repeater
End If

Fruit.Left = ColorLed(0).Width * X + 30
Fruit.Top = ColorLed(0).Height * Y

For i = (UBound(Pos) - ColorLed.Count) To UBound(Pos)
    If Pos(i).PosX = Fruit.Left And Pos(i).PosY = Fruit.Top Then
        GoTo Repeater
    End If
Next i

Fruit.Visible = True

End Sub

Private Sub Timer2_Timer()

If ColorLed(ColorLed.Count - 1).Left < 0 Then
GoTo Loser
End
End If

If ColorLed(ColorLed.Count - 1).Left > Me.Width - ColorLed(0).Width Then
GoTo Loser
End
End If

If ColorLed(ColorLed.Count - 1).Top < 0 Then
GoTo Loser
End
End If

If ColorLed(ColorLed.Count - 1).Top > Me.Height - ColorLed(0).Height * 3 + 1 Then
GoTo Loser
End
End If

Exit Sub

Loser:
Call Lose
End Sub

Private Sub Lose()
Timer1.Enabled = False
Timer2.Enabled = False
Dim Resp As Integer
Resp = MsgBox("You Lose!" & vbNewLine & "Score: " & Eaten * Level & vbNewLine & "Try Again?", vbYesNo Or vbExclamation, "Try Again?")
If Resp = vbYes Then
    Call Reloader
    Timer1.Enabled = True
    Timer2.Enabled = True
Else
    MsgBox "Thank you for playing ByteDive Sanke. Xmas 2001 by Mario Di Vece (mariodivece@hotmail.com)", vbInformation, "ByteDive Software"
    End
End If
End Sub

Private Sub Reloader()
For i = 1 To ColorLed.Count - 1
    Unload ColorLed(i)
Next i
    Form_Load
End Sub
