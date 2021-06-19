VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSignon 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SignIn"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "frmSignon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " SignIn"
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3495
      Begin VB.CheckBox chkUS 
         Caption         =   "Save Username/Server IP"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   0
         ToolTipText     =   "Enter your name so that other clients may recognize you"
         Top             =   345
         Width           =   2175
      End
      Begin VB.TextBox txtSvrIp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   1
         ToolTipText     =   "Enter the IP address or computer-name of computer where chat server is running"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblVersion 
         Caption         =   "Vesion:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Server IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   975
         Width           =   735
      End
   End
   Begin MSWinsockLib.Winsock wskSignon 
      Left            =   2040
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   2220
      Left            =   0
      Picture         =   "frmSignon.frx":1CFA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmSignon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
If InStr(txtUser.Text, " ") <> 0 Then
    MsgBox "UserID must be in single string", , "MyChat-Error"
ElseIf txtUser.Text <> "" Then
If txtSvrIp.Text <> "" Then
    frmChat.Show
    frmChat.txtUser.Text = txtUser.Text
    Call frmChat.Connect(txtSvrIp.Text)
    frmChat.Caption = "MyChat - (" & txtUser.Text & ")"
    DoEvents
    If chkUS.Value = 1 Then
        Call WriteINI("UserSvr", "User", Trim(txtUser.Text), App.Path & "\Options.ini")
        Call WriteINI("UserSvr", "Svr", Trim(txtSvrIp.Text), App.Path & "\Options.ini")
        DoEvents
    Else
        Call WriteINI("UserSvr", "User", "", App.Path & "\Options.ini")
        Call WriteINI("UserSvr", "Svr", "", App.Path & "\Options.ini")
        DoEvents
    End If
    Unload Me
End If
End If
End Sub

Private Sub Form_Load()
Dim Usr As String, Svr As String
Usr$ = ReadINI("UserSvr", "User", App.Path & "\Options.ini")
Svr$ = ReadINI("UserSvr", "Svr", App.Path & "\Options.ini")

If Usr$ <> "" Then
    txtUser.Text = Usr$
    txtSvrIp.Text = Svr$
    chkUS.Value = 1
End If

lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
