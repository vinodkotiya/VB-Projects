VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Instant Messanger"
   ClientHeight    =   6795
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7575
   Icon            =   "frmOM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   4800
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   5160
      Width           =   375
   End
   Begin VB.ComboBox cmbSize 
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   7
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ListBox listUsers 
      Height          =   4350
      Left            =   5280
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox rtfMsg 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmOM.frx":0CCA
   End
   Begin RichTextLib.RichTextBox rtfRoom 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmOM.frx":0D4C
   End
   Begin VB.Label lblCol 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "Status : Not Connected"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   7935
   End
   Begin VB.Menu user 
      Caption         =   "&User"
      Begin VB.Menu mnuLog 
         Caption         =   "Logout"
         Index           =   0
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Change User"
         Index           =   1
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelp 
         Caption         =   "Help"
         Index           =   0
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "About"
         Index           =   1
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "About Me"
         Index           =   2
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "http:\\vinodkotiya.tripod.com"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
Dim tmp As Integer
 tmp = UpdateState
End Sub

Private Sub cmbSize_Validate(Cancel As Boolean)

 rtfMsg.SelFontSize = cmbSize.List(cmbSize.ListIndex)

End Sub

Private Sub cmdFormat_Click(Index As Integer)
If Index = 0 Then
 rtfMsg.SelBold = Not rtfMsg.SelBold
ElseIf Index = 1 Then
 rtfMsg.SelItalic = Not rtfMsg.SelItalic
ElseIf Index = 1 Then
 rtfMsg.SelUnderline = Not rtfMsg.SelUnderline
End If
 
End Sub

Private Sub cmdSend_Click()
'
'UserName = "vinod"
Dim filter As String
Dim DataSent As String
'rtfRoom.SelStart = Len(rtfRoom.Text)
'rtfRoom.SelLength = 1
'rtfRoom.SelColor = vbBlack
'rtfRoom.SelBold = False

filter = rtfMsg.TextRTF
'filter = Left(filter, Len(filter) - 4)
'filter = Right(filter, Len(filter) - 86)
Dim pos As Double
Dim upper As String
pos = InStr(InStr(1, filter, ";}}") + 3, filter, ";}")
'pos = 0 if colortbl not exist else return posn of ;} for colortbl end
If pos > 0 Then
   upper = Left(filter, pos + 1)
Else
  upper = Left(filter, InStr(1, filter, ";}}") + 3)
End If
'MsgBox upper
filter = Left(filter, InStrRev(filter, "}") - 4)
filter = Right(filter, Len(filter) - InStr(1, filter, ";}}", vbBinaryCompare) - 3)
'MsgBox filter
'"{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}" &
rtfMsg.TextRTF = upper & _
          " \b " & UserName & " \b0 said, '" & filter & "' }"
 '         MsgBox rtfMsg.TextRTF
rtfMsg.SelStart = 1
rtfMsg.SelLength = Len(UserName)
rtfMsg.SelColor = vbRed
rtfMsg.SelBold = True
DataSent = rtfMsg.TextRTF
'Clipboard.SetText rtfMsg.TextRTF, rtfText

frmLogon.wsNet.SendData DataSent ' sent data

rtfMsg.SelColor = vbBlue

'rtfRoom.SelLength = Len(rtfRoom.Text)

filter = rtfMsg.TextRTF

'filter = Left(filter, Len(filter) - 4)
'filter = Right(filter, Len(filter) - 76)
filter = Left(filter, InStrRev(filter, "}") - 1)
filter = Right(filter, Len(filter) - InStr(1, filter, ";}}", vbBinaryCompare) - 3)
'MsgBox filter
rtfRoom.TextRTF = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}" & _
              rtfRoom.TextRTF & filter & " }"  'rtfMsg.TextRTF & "}"
 ' MsgBox rtfRoom.TextRTF
'rtfRoom.TextRTF = rtfRoom.Text
'rtfMsg.Text = "" ' clear the textbox
'MsgBox rtfMsg.SelRTF
Dim tmp As Integer
 tmp = UpdateState
rtfMsg.SetFocus
End Sub

Private Sub Command1_Click()
MsgBox rtfRoom.SelRTF
End Sub

Private Sub Command2_Click()
MsgBox rtfMsg.SelRTF

End Sub



Private Sub Form_Load()
cmbSize.AddItem "8"
cmbSize.AddItem "10"
cmbSize.AddItem "12"
cmbSize.AddItem "14"
cmbSize.AddItem "16"
cmbSize.AddItem "18"
cmbSize.AddItem "20"
cmbSize.AddItem "24"
cmbSize.AddItem "28"
cmbSize.AddItem "36"
cmbSize.AddItem "48"
cmbSize.ListIndex = 1
End Sub

Private Sub lblCol_Click()
Dim CDFlags As Long
Dim Rang As Long

On Error GoTo ColorError

    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)

    CommonDialog1.Flags = CDFlags
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    Rang& = CommonDialog1.Color     'obtained BGR color
    lblCol.BackColor = Rang&
    rtfMsg.SelColor = Rang&
    Exit Sub
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
    
    Else
        MsgBox "An error occured"
    End If

End Sub

Private Sub mnuLog_Click(Index As Integer)
If Index = 0 Then frmLogon.wsNet.Close
End Sub

Private Sub Text1_Click()

End Sub
