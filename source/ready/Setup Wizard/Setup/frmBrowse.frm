VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBrowse.frx":0000
   ScaleHeight     =   2520
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1440
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   3720
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1920
   End
   Begin VB.Image imgCancel 
      Height          =   300
      Index           =   0
      Left            =   240
      Picture         =   "frmBrowse.frx":241A
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Installation Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image imgCancel 
      Height          =   300
      Index           =   1
      Left            =   240
      Picture         =   "frmBrowse.frx":2807
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image imgOk 
      Height          =   300
      Index           =   1
      Left            =   1320
      Picture         =   "frmBrowse.frx":2BFD
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image imgOk 
      Height          =   300
      Index           =   0
      Left            =   1320
      Picture         =   "frmBrowse.frx":2FA2
      Top             =   2040
      Width           =   750
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
ChDir Dir1.path
txtDir.Text = Dir1.path
End Sub

Private Sub Drive1_Change()
On Error GoTo vinerror
ChDrive Dir1.path
    Dir1.path = Drive1.Drive
    Dir1.Refresh
  Exit Sub
vinerror:
  MsgBox "There is no disk in drive"
End Sub

Private Sub Form_Load()
 Dim pfiles As String
 pfiles = Environ("ProgramFiles")
    ' pfiles = Right(pfiles, Len(pfiles) - 13)
Dir1.path = pfiles
End Sub

Private Sub imgCancel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgCancel(1).Visible = True
 imgCancel(0).Visible = False
End If
End Sub

Private Sub imgCancel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then
 imgCancel(0).Visible = True
 imgCancel(1).Visible = False
End If
Unload Me
End Sub

Private Sub imgOk_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then
 imgOk(0).Visible = True
 imgOk(1).Visible = False
End If
End Sub

Private Sub imgOk_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgOk(1).Visible = True
 imgOk(0).Visible = False
End If
Index = InStrRev(Trim(txtDir.Text), "\")
Index = Len(Trim(txtDir.Text)) - Index
'MsgBox Index
If Index = 0 Then
 txtDir.Text = Left(Trim(txtDir.Text), Len(Trim(txtDir.Text)) - 1)
End If
frmStart.txtDir.Caption = Trim(txtDir.Text)
Unload Me
End Sub
