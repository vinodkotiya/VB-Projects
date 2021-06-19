VERSION 5.00
Begin VB.Form frmPM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyChat - Private Message"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmPM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Private Message"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtPM 
         Height          =   1125
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton cmdSendPM 
         Caption         =   "Send PM"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Whatever you type in this window only the user you selected will see."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Send to:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblUserPM 
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdSendPM_Click()
If txtPM.Text <> "" Then
Dim allText As String
allText = frmChat.txtUser.Text & "|" & txtPM.Text
Call frmChat.wskClient.SendData("PMMessage " & lblUserPM.Caption & ":" & allText)
txtPM.Text = ""
DoEvents
End If
End Sub


Private Sub txtPM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSendPM_Click
    KeyAscii = 0
End If
End Sub
