VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Instant Messanger"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3675
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   960
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
      Begin VB.TextBox txtIP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "127.0.0.1"
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optSys 
         Caption         =   "As Host"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSys 
         Caption         =   "As Server"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Text            =   "554"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblIp 
         BackStyle       =   0  'Transparent
         Caption         =   "With Local Port No ="
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock wsNet 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "Status : Not Connected"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblClick 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here to Sign in"
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
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
wsNet.Close
If Len(Trim(txtUser.Text)) < 1 Then
 MsgBox "Please enter a valid username"
 txtUser.SetFocus
 Exit Sub
 End If
 
Dim ret As Integer
If optSys(0).Value Then
 wsNet.LocalPort = Trim(txtPort.Text)
 wsNet.Listen ' listen for others to connect you
ElseIf optSys(1).Value Then
 wsNet.Connect Trim(txtIP.Text), Trim(txtPort.Text)   ' connect to server listening on port 554
End If
UserName = txtUser.Text

 ret = UpdateState
 'If ret = 7 Then
 Load frmOM
 frmOM.Show
 'End If
 frmOM.Caption = UserName & " VIN Instant Messanger "
End Sub

Private Sub Command1_Click()
MsgBox wsNet.LocalPort
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox UserName & " r u sure"
wsNet.Close
Unload frmOM
End Sub

Private Sub optSys_Click(Index As Integer)
If optSys(1).Value Then
  txtIP.Enabled = True
  lblIp.Caption = "Server's Port No is"
 Else
 txtIP.Enabled = False
 lblIp.Caption = "With Local Port No ="
 End If
End Sub

Private Sub Timer1_Timer()
Dim ret As Integer
 ret = UpdateState
End Sub

Private Sub wsNet_ConnectionRequest(ByVal requestID As Long)
If wsNet.State <> sckClosed Then wsNet.Close ' if not closed the connection close it
wsNet.Accept requestID ' accept the requestid

End Sub

Private Sub wsNet_DataArrival(ByVal bytesTotal As Long)
Dim IncomeData As String
wsNet.GetData IncomeData ' get data
'frmOM.rtfRoom.TextRTF = frmOM.rtfRoom.Text & vbCrLf & IncomeData   ' append the data to the textbox

IncomeData = Left(IncomeData, InStrRev(IncomeData, "}") - 1)
IncomeData = Right(IncomeData, Len(IncomeData) - InStr(1, IncomeData, ";}}", vbBinaryCompare) - 3)
'MsgBox IncomeData
frmOM.rtfRoom.TextRTF = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}" & _
              frmOM.rtfRoom.TextRTF & IncomeData & " }"  'rtfMsg.TextRTF & "}"

End Sub

