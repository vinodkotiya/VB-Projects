VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Environment Variables"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRet 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2400
      Max             =   26
      Min             =   1
      TabIndex        =   2
      Top             =   240
      Value           =   1
      Width           =   1575
   End
   Begin VB.TextBox txtEn 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Environment Variables"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Integer
Dim ret As String
For i = 1 To 26
 ret = ret & i & "  " & Environ(i) & vbCrLf
Next
MsgBox ret
End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub HScroll1_Change()
txtEn.Text = HScroll1.Value
txtRet.Text = Environ(HScroll1.Value)
End Sub

Private Sub Label1_Click()
MsgBox "Written By : VINOD KOTIYA" & vbCrLf & " http:\\vinodkotiya.tripod.com"
End Sub
