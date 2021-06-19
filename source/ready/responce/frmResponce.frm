VERSION 5.00
Begin VB.Form frmResponce 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Responce"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   2790
   HasDC           =   0   'False
   Icon            =   "frmResponce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   2640
   End
   Begin VB.CommandButton button 
      Caption         =   "Response"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton button 
      Caption         =   "Start"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox display 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1785
      ScaleWidth      =   2265
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Issued in Public Interest  By: VINOD KOTIYA   http:\\vinodkotiya.tripod.com"
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmResponce.frx":1CFA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
End
Attribute VB_Name = "frmResponce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startTime As Integer
Dim Twosec As Integer    'make a msec loop (2000) for timer
    

Private Sub button_Click(Index As Integer)
If Index = 0 Then
 Timer1.Interval = 1
 Twosec = 1
 startTime = 0
 button(1).Enabled = True
 button(0).Enabled = False
ElseIf Index = 1 Then
If startTime = 0 Then
  MsgBox "Are you in hurry.First Wait to appear the color. " & vbCrLf & _
  "             You Chitter !!! "
  Timer1.Interval = 0
  button(0).Enabled = True
   button(0).SetFocus
 button(1).Enabled = False
   Exit Sub
   End If
 If startTime / 100 < 0.2 Then
  MsgBox "Your responce time is " & startTime / 100 & " Seconds " & vbCrLf & _
   "               UNBELIVABLE"
 ElseIf startTime / 100 > 0.2 And startTime / 100 < 0.5 Then
  MsgBox "Your responce time is " & startTime / 100 & " Seconds " & vbCrLf & _
   "               Keep it up !!!"
  ElseIf startTime / 100 > 0.5 And startTime / 100 < 0.7 Then
  MsgBox "Your responce time is " & startTime / 100 & " Seconds " & vbCrLf & _
   "               You need more practice !!!"
 ElseIf startTime / 100 > 0.7 And startTime / 100 < 2 Then
  MsgBox "Your responce time is " & startTime / 100 & " Seconds " & vbCrLf & _
   "               Are you drunk !!!"
 ElseIf startTime / 100 > 2 Then
  MsgBox "Your responce time is " & startTime / 100 & " Seconds " & vbCrLf & _
   "               It's toooo Much.. !!!"
 End If
  Timer1.Interval = 0
  display.BackColor = &HFFFFFF
  button(0).Enabled = True
  button(0).SetFocus
 button(1).Enabled = False
End If
End Sub



Private Sub Timer1_Timer()
If startTime > 0 Then
  startTime = startTime + 1
  Exit Sub
End If
Twosec = Twosec + 1
If Twosec > 200 Then  ''2 seconds complete after pressing start
 
   If Rnd(20) < 0.5 Then
        display.BackColor = Rnd(456) * 65000
        startTime = 1
        End If
End If
End Sub
