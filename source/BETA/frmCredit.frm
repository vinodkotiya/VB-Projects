VERSION 5.00
Begin VB.Form frmCredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDIT"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   Icon            =   "frmCredit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Text            =   "click me twice"
      Top             =   4880
      Width           =   1215
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   885
      Left            =   3120
      OleObjectBlob   =   "frmCredit.frx":0ECA
      SourceDoc       =   "F:\credit\MORE...htm"
      TabIndex        =   0
      Top             =   4320
      Width           =   1440
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   6240
      Picture         =   "frmCredit.frx":56E2
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6060
      Left            =   0
      Picture         =   "frmCredit.frx":65AC
      Top             =   0
      Width           =   7560
   End
End
Attribute VB_Name = "frmCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()

End Sub

