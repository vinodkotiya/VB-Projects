VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New User"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPass 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "Create User"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPass_Click()
If chkPass.Value Then
  txtPassword.PasswordChar = "*"
Else
 txtPassword.PasswordChar = ""
End If
 
End Sub

Private Sub cmdMake_Click()
Dim i As Integer
If Trim(txtUser.Text) = "" Then
   MsgBox "Please Enter a valid user Name."
   txtUser.SetFocus
   Exit Sub
End If
If Trim(txtPassword.Text) = "" Then
   MsgBox "Please Enter a valid password."
   txtPassword.SetFocus
   Exit Sub
End If
For i = 1 To colUserName.Count
 If Trim(txtUser.Text) = colUserName.Item(i) Then
  MsgBox "User name " & txtUser.Text & " is already Exist.Please use any other user Name."
  txtUser.Text = txtUser.Text & "1"
  Exit Sub
  End If
Next
colUserName.Add Trim(txtUser.Text)
colPassword.Add Trim(txtPassword.Text)
Dim fnum As Integer
fnum = FreeFile    'getting file no for futures referance
 Open App.Path & "\data\" & Trim(txtUser.Text) & ".vin" For Output As fnum     'dont use #1 for multiple file openings
 Print #fnum, "<start>"
 Print #fnum, "</end>"
Close #fnum


MsgBox "The New User " & txtUser.Text & " is successfully Created." & vbCrLf & _
"Now use your username and password to Login to password menager."
 Unload Me
 frmLogin.ReLoadCombo
 frmLogin.Show
End Sub

