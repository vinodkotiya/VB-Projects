VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Beta Testing Report"
   ClientHeight    =   4185
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   5550
   Icon            =   "betareport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "unknown"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Uninstall"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Remarks 
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Also write your comments,views,any modification about this software."
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "If softwares give any error number please report it."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Name of Beta Tester"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Please Enter or attach any bug report or log file generated during program execution."
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.Menu mail 
      Caption         =   "Send By &Mail"
   End
   Begin VB.Menu floppy 
      Caption         =   "Send By &Floppy"
   End
   Begin VB.Menu print 
      Caption         =   "&Print"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim r As String
r = InputBox("Please enter the 50 digit uninstallation code" _
& Chr(13) & " you have get after submitting the Beta testing report", "UnInstallation Code")

If Trim(r) <> "" Then MsgBox "The code you have entered is invalid , UnInstallation Failed" & _
Chr(13) & " You must need the Uninstallation Code to remove this software and registry keys"
End Sub

Private Sub floppy_Click()
On Error GoTo vinerror
MsgBox "Insert a floppy disk .Beta report is saved as " & Chr(34) & " A:\" & Text1.Text & ".vin" & Chr(34) & _
" In your floppy disk."
Dim fnum As Integer
fnum = FreeFile
Open "A:\" & Text1.Text & ".vin" For Output As #fnum
Print #fnum, Remarks.Text
Close #fnum
Exit Sub
vinerror:
 MsgBox "Device not ready.Check your floppy disk drive"
End Sub

Private Sub Form_Load()
'On Error GoTo chup
 Dim fnum As Integer
 Dim txt As String
fnum = FreeFile
Open "C:\WINDOWS\vinbakup\user.vin" For Input As #fnum
Line Input #fnum, txt
Close #fnum
Text1.Text = txt
Exit Sub
chup:
 
End Sub

Private Sub mail_Click()
MsgBox "Beta report is saved as " & Chr(34) & " C:\" & Text1.Text & ".vin" & Chr(34) & _
" In your hard drive.Please attach this file with your email to " & _
" vinodkotiya24@yahoomail.com"
Dim fnum As Integer
fnum = FreeFile
Open "c:\" & Text1.Text & ".vin" For Output As #fnum
Print #fnum, Remarks.Text
Close #fnum
End Sub

