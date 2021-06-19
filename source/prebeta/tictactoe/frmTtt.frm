VERSION 5.00
Begin VB.Form frmTtt 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image img11 
      Height          =   735
      Left            =   360
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Null33 
      Height          =   720
      Left            =   2640
      Picture         =   "frmTtt.frx":0000
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Null32 
      Height          =   720
      Left            =   1440
      Picture         =   "frmTtt.frx":0ECA
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Null31 
      Height          =   720
      Left            =   360
      Picture         =   "frmTtt.frx":1D94
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Null23 
      Height          =   720
      Left            =   2640
      Picture         =   "frmTtt.frx":2C5E
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Null22 
      Height          =   720
      Left            =   1440
      Picture         =   "frmTtt.frx":3B28
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Null21 
      Height          =   720
      Left            =   360
      Picture         =   "frmTtt.frx":49F2
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Null13 
      Height          =   720
      Left            =   2640
      Picture         =   "frmTtt.frx":58BC
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Null12 
      Height          =   720
      Left            =   1440
      Picture         =   "frmTtt.frx":6786
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderWidth     =   10
      X1              =   2400
      X2              =   2400
      Y1              =   600
      Y2              =   3360
   End
   Begin VB.Image img33 
      Height          =   720
      Left            =   2640
      Top             =   2640
      Width           =   720
   End
   Begin VB.Image img32 
      Height          =   720
      Left            =   1440
      Top             =   2640
      Width           =   720
   End
   Begin VB.Image img31 
      Height          =   720
      Left            =   360
      Top             =   2640
      Width           =   720
   End
   Begin VB.Image img23 
      Height          =   720
      Left            =   2640
      Top             =   1560
      Width           =   720
   End
   Begin VB.Image img22 
      Height          =   720
      Left            =   1440
      Top             =   1560
      Width           =   720
   End
   Begin VB.Image img21 
      Height          =   720
      Left            =   360
      Top             =   1560
      Width           =   720
   End
   Begin VB.Image img13 
      Height          =   720
      Left            =   2640
      Top             =   600
      Width           =   720
   End
   Begin VB.Image img12 
      Height          =   720
      Left            =   1440
      Top             =   600
      Width           =   720
   End
   Begin VB.Image Null11 
      Height          =   720
      Left            =   360
      Picture         =   "frmTtt.frx":7650
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Line Line4 
      BorderWidth     =   10
      X1              =   240
      X2              =   3480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderWidth     =   10
      X1              =   240
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   1320
      X2              =   1320
      Y1              =   480
      Y2              =   3360
   End
   Begin VB.Image Cross11 
      Height          =   720
      Left            =   360
      Picture         =   "frmTtt.frx":851A
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross12 
      Height          =   720
      Left            =   1440
      Picture         =   "frmTtt.frx":A1E4
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross13 
      Height          =   720
      Left            =   2640
      Picture         =   "frmTtt.frx":BEAE
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross21 
      Height          =   720
      Left            =   360
      Picture         =   "frmTtt.frx":DB78
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross22 
      Height          =   720
      Left            =   1440
      Picture         =   "frmTtt.frx":F842
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross23 
      Height          =   720
      Left            =   2640
      Picture         =   "frmTtt.frx":1150C
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross31 
      Height          =   720
      Left            =   360
      Picture         =   "frmTtt.frx":131D6
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross32 
      Height          =   720
      Left            =   1440
      Picture         =   "frmTtt.frx":14EA0
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Cross33 
      Height          =   720
      Left            =   2640
      Picture         =   "frmTtt.frx":16B6A
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmTtt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Done As Integer
Option Base 1
Dim Field(3, 3) As Integer   ''store 0 for null & 1 for cross & 2 for draw & 6 for blank


Private Sub imgNull_Click()

End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer, upto As Integer
upto = 3
For i = 1 To upto
 For j = 1 To upto
  Field(i, j) = 6
 Next j
Next i

End Sub

Private Sub img11_Click()
Cross11.Visible = True
img11.Visible = False
Null11.Visible = False
Field(1, 1) = 1
Done = Check()
If Done = 1 Then   ''if check returns 1(X)
 Vinneru
ElseIf Done = 0 Then ''if check return 0 (o)
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img12_Click()
Cross12.Visible = True
img12.Visible = False
Null12.Visible = False
Field(1, 2) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img13_Click()
Cross13.Visible = True
img13.Visible = False
Null13.Visible = False
Field(1, 3) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img21_Click()
Cross21.Visible = True
img21.Visible = False
Null21.Visible = False
Field(2, 1) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img22_Click()
Cross22.Visible = True
img22.Visible = False
Null22.Visible = False
Field(2, 2) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img23_Click()
Cross23.Visible = True
img23.Visible = False
Null23.Visible = False
Field(2, 3) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img31_Click()
Cross31.Visible = True
img31.Visible = False
Null31.Visible = False
Field(3, 1) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img32_Click()
Cross32.Visible = True
img32.Visible = False
Null32.Visible = False
Field(3, 2) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Private Sub img33_Click()

Cross33.Visible = True
img33.Visible = False
Null33.Visible = False
Field(3, 3) = 1
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
Else
Get_computers_move
End If
Done = Check()
If Done = 1 Then
 Vinneru
ElseIf Done = 0 Then
 Vinnercom
ElseIf Done = 2 Then
 draw
End If

End Sub

Function Vinneru()
 MsgBox "You are the vinner"
 
End Function
Function Vinnercom()
 MsgBox "Computer is the vinner"
  
End Function
Function draw()
 MsgBox "Game is Draw"
 
End Function
 
Function Check() As Integer
 Dim i As Integer, j As Integer, draw As Integer
 draw = 1
 For i = 1 To 3  'check rows
  If (Field(i, 1) = Field(i, 2) And Field(i, 2) = Field(i, 3)) Then
  Check = Field(i, 1)
  Exit Function
  End If
  Next i
 For i = 1 To 3 'check col
  If (Field(1, i) = Field(2, i) And Field(1, i) = Field(3, i)) Then
  Check = Field(1, i)
  Exit Function
  End If
            ' check digonal
 If (Field(1, 1) = Field(2, 2) And Field(2, 2) = Field(3, 3)) Then
  Check = Field(1, 1)
  Exit Function
 End If
 If (Field(1, 3) = Field(2, 2) And Field(2, 2) = Field(3, 1)) Then
  Check = Field(1, 2)
  Exit Function
 End If
 Next i
          'check draw or not by searching 6
 For i = 1 To 3
  For j = 1 To 3
   If (Field(i, j) = 6) Then
   draw = draw + 1
   End If
  Next j
 Next i
 If (draw = 1) Then 'no 6(blank) found
  Check = 2  ' game is draw
  Exit Function
 End If

 'Return ' ';
Check = 6
End Function

Function Get_computers_move()
   'computer search O for triplets
Dim i As Integer, j As Integer, Played As Integer
Played = 0
  For i = 1 To 3   'check rows for winning
   If (Field(i, 1) = 0 And Field(i, 2) = 0 And Field(i, 3) = 6) Then
     Field(i, 3) = 0
     Played = 1
     If (i = 1) Then
      Null13.Visible = True
     ElseIf (i = 2) Then
      Null23.Visible = True
     Else
      Null33.Visible = True
     End If
   ElseIf (Field(i, 1) = 0 And Field(i, 3) = 0 And Field(i, 2) = 6) Then
     Field(i, 2) = 0
     Played = 1
     If (i = 1) Then
      Null12.Visible = True
     ElseIf (i = 2) Then
      Null22.Visible = True
     Else
      Null32.Visible = True
     End If

   ElseIf (Field(i, 2) = 0 And Field(i, 3) = 0 And Field(i, 1) = 6) Then
     Field(i, 1) = 0
     Played = 1
     If (i = 1) Then
      Null11.Visible = True
     ElseIf (i = 2) Then
      Null21.Visible = True
     Else
      Null31.Visible = True
     End If

   End If
 Next i
  If (Played = 1) Then
   Exit Function 'com turns over so exit from loop
  End If
  
  For j = 1 To 3 ' check cols for winning */
   If (Field(1, j) = 0 And Field(2, j) = 0 And Field(3, j) = 6) Then
    Field(3, j) = 0
    Played = 1
    If (j = 1) Then
      Null31.Visible = True
     ElseIf (j = 2) Then
      Null32.Visible = True
     Else
      Null33.Visible = True
     End If

   ElseIf (Field(1, j) = 0 And Field(3, j) = 0 And Field(2, j) = 6) Then
    Field(2, j) = 0
    Played = 1
    If (j = 1) Then
      Null21.Visible = True
     ElseIf (j = 2) Then
      Null22.Visible = True
     Else
      Null23.Visible = True
     End If

   ElseIf (Field(2, j) = 0 And Field(3, j) = 0 And Field(1, j) = 6) Then
    Field(1, j) = 0
    Played = 1
    If (j = 1) Then
      Null11.Visible = True
     ElseIf (j = 2) Then
      Null12.Visible = True
     Else
      Null13.Visible = True
     End If

   End If
  Next j
   
  If (Played = 1) Then
  Exit Function     'com turns over so exit from loop
  End If
  
    ' check digonal for winning
  If (Field(1, 1) = 0 And Field(2, 2) = 0 And Field(3, 3) = 6) Then
   Field(3, 3) = 0
   Null33.Visible = True
   Played = 1
  ElseIf (Field(1, 1) = 0 And Field(3, 3) = 0 And Field(2, 2) = 6) Then
    Field(2, 2) = 0
    Null22.Visible = True
    Played = 1
  ElseIf (Field(2, 2) = 0 And Field(3, 3) = 0 And Field(1, 1) = 6) Then
    Field(1, 1) = 0
    Null11.Visible = True
    Played = 1
  ElseIf (Field(1, 3) = 0 And Field(2, 2) = 0 And Field(3, 1) = 6) Then
    Field(3, 1) = 0
    Null31.Visible = True
    Played = 1
  ElseIf (Field(1, 3) = 0 And Field(3, 1) = 0 And Field(2, 2) = 6) Then
    Field(2, 2) = 0
    Null22.Visible = True
    Played = 1
  ElseIf (Field(2, 2) = 0 And Field(3, 1) = 0 And Field(1, 3) = 6) Then
    Field(1, 3) = 0
    Null13.Visible = True
    Played = 1
  End If
  If (Played = 1) Then
  Exit Function          'com turns over so exit from loop
  End If

  'computer will search dobled X to block

 For i = 1 To 3   ' check rows for blocking*/
   If (Field(i, 1) = 1 And Field(i, 2) = 1 And Field(i, 3) = 6) Then
      Field(i, 3) = 0
      Played = 1
     If (i = 1) Then
      Null13.Visible = True
     ElseIf (i = 2) Then
      Null23.Visible = True
     Else
      Null33.Visible = True
     End If
   ElseIf (Field(i, 1) = 1 And Field(i, 3) = 1 And Field(i, 2) = 6) Then
      Field(i, 2) = 0
      Played = 1
      If (i = 1) Then
      Null12.Visible = True
     ElseIf (i = 2) Then
      Null22.Visible = True
     Else
      Null32.Visible = True
     End If
   ElseIf (Field(i, 2) = 1 And Field(i, 3) = 1 And Field(i, 1) = 6) Then
      Field(i, 1) = 0
      Played = 1
      If (i = 1) Then
      Null11.Visible = True
     ElseIf (i = 2) Then
      Null21.Visible = True
     Else
      Null31.Visible = True
     End If
   End If
  Next i
  If (Played = 1) Then
   Exit Function 'com turns over so exit from loop
  End If
  
  For j = 1 To 3  ' check cols for blocking */
    
   If (Field(1, j) = 1 And Field(2, j) = 1 And Field(3, j) = 6) Then
     Field(3, j) = 0
     Played = 1
     If (j = 1) Then
      Null31.Visible = True
     ElseIf (j = 2) Then
      Null32.Visible = True
     Else
      Null33.Visible = True
     End If

   ElseIf (Field(1, j) = 1 And Field(3, j) = 1 And Field(2, j) = 6) Then
     Field(2, j) = 0
     Played = 1
     If (j = 1) Then
      Null21.Visible = True
     ElseIf (j = 2) Then
      Null22.Visible = True
     Else
      Null23.Visible = True
     End If

   ElseIf (Field(2, j) = 1 And Field(3, j) = 1 And Field(1, j) = 6) Then
     Field(1, j) = 0
     Played = 1
     If (j = 1) Then
      Null11.Visible = True
     ElseIf (j = 2) Then
      Null12.Visible = True
     Else
      Null13.Visible = True
     End If

   End If
     
  Next j
  If (Played = 1) Then
  Exit Function          'com turns over so exit from loop
  End If
    ' check digonal for blocking
  If (Field(1, 1) = 1 And Field(2, 2) = 1 And Field(3, 3) = 6) Then
    Field(3, 3) = 0
    Played = 1
    Null33.Visible = True
  ElseIf (Field(1, 1) = 1 And Field(3, 3) = 1 And Field(2, 2) = 6) Then
    Field(2, 2) = 0
    Played = 1
    Null22.Visible = True
  ElseIf (Field(2, 2) = 1 And Field(3, 3) = 1 And Field(1, 1) = 6) Then
    Field(1, 1) = 0
    Played = 1
    Null11.Visible = True
  ElseIf (Field(1, 3) = 1 And Field(2, 2) = 1 And Field(3, 1) = 6) Then
    Field(3, 1) = 0
    Played = 1
    Null31.Visible = True
  ElseIf (Field(1, 3) = 1 And Field(3, 1) = 1 And Field(2, 2) = 6) Then
    Field(2, 2) = 0
    Played = 1
    Null22.Visible = True
  ElseIf (Field(2, 2) = 1 And Field(3, 1) = 1 And Field(1, 3) = 6) Then
    Field(1, 3) = 0
    Played = 1
    Null13.Visible = True
  End If
  If (Played = 1) Then
  Exit Function  'com turns over so exit from loop
  End If

          'normal tricks
  If (Field(2, 2) = 6) Then
    Field(2, 2) = 0
    Null22.Visible = True
  ElseIf (Field(3, 3) = 6) Then
    Field(3, 3) = 0
    Null33.Visible = True
  ElseIf (Field(3, 1) = 6) Then
    Field(3, 1) = 0
    Null31.Visible = True
  ElseIf (Field(1, 3) = 6) Then
    Field(1, 3) = 0
    Null13.Visible = True
  ElseIf (Field(1, 1) = 6) Then
    Field(1, 1) = 0
    Null11.Visible = True
  ElseIf (Field(2, 1) = 6) Then
    Field(2, 1) = 0
    Null21.Visible = True
  ElseIf (Field(2, 3) = 6) Then
    Field(2, 3) = 0
    Null23.Visible = True
  ElseIf (Field(3, 2) = 6) Then
    Field(3, 2) = 0
    Null32.Visible = True
  ElseIf (Field(1, 2) = 6) Then
    Field(1, 2) = 0
    Null12.Visible = True
  'Else
  'cout<<"\n vinner"
  'return 0
  End If

End Function

