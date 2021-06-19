VERSION 5.00
Begin VB.Form frmPrompt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyChat - Prompt"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmPrompt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Prompt"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtPrompt 
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
If txtPrompt <> "" Then
    Call frmChat.wskClient.SendData("Prompt " & txtPrompt.Text)
    DoEvents
    txtPrompt.Text = ""
End If
End Sub

Private Sub txtPrompt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub
