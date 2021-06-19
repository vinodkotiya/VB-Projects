VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm2 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Script: (ALL2HTML CONVERTER)"
   ClientHeight    =   4245
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   4740
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtScript 
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   1120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5106
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frm2.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   450
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Choose portion where script to be added"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   3255
      Begin VB.OptionButton optBody 
         BackColor       =   &H00FF0000&
         Caption         =   "inside <BODY>"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optHead 
         BackColor       =   &H00FF0000&
         Caption         =   "inside <HEAD>"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Click here for some spectacular INBUILT SCRIPTS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4040
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm2.frx":0082
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu script 
      Caption         =   "&Inbuilt Scripts"
      Visible         =   0   'False
      Begin VB.Menu mnuscripts 
         Caption         =   "Dancing Star"
         Index           =   0
      End
      Begin VB.Menu mnuscripts 
         Caption         =   "Elastic Image"
         Index           =   1
      End
      Begin VB.Menu mnuscripts 
         Caption         =   "Ellipsing Text"
         Index           =   2
      End
      Begin VB.Menu mnuscripts 
         Caption         =   "Magic Wend"
         Index           =   3
      End
      Begin VB.Menu mnuscripts 
         Caption         =   "Magic Wend II"
         Index           =   4
      End
      Begin VB.Menu mnuscripts 
         Caption         =   "Cursor Trailor"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frm2.Visible = False
End Sub

Private Sub Command2_Click()
txtScript.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift And vbAltMask Then
  If KeyCode = 67 Or KeyCode = 99 Then
   PopupMenu script
   End If
 End If
End Sub

Private Sub Label2_Click()
PopupMenu script
End Sub

Private Sub mnuscripts_Click(Index As Integer)
Index = Index + 1
txtScript.LoadFile App.Path & "\data\script" & Index & ".vin", rtfText
optBody.Value = True
End Sub

Private Sub txtScript_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift And vbAltMask Then
  If KeyCode = 67 Or KeyCode = 99 Then
   PopupMenu script
   End If
 End If
End Sub
