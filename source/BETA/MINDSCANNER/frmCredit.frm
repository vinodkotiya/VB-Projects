VERSION 5.00
Begin VB.Form frmCredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDIT"
   ClientHeight    =   4380
   ClientLeft      =   1575
   ClientTop       =   2430
   ClientWidth     =   9030
   Icon            =   "frmCredit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9030
   Visible         =   0   'False
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   765
      Left            =   4080
      OleObjectBlob   =   "frmCredit.frx":0ECA
      SourceDoc       =   "F:\credit\mind.html"
      TabIndex        =   0
      Top             =   3600
      Width           =   1305
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4380
      Left            =   0
      Picture         =   "frmCredit.frx":5CE2
      Top             =   0
      Width           =   9060
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


