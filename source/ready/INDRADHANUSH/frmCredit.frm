VERSION 5.00
Begin VB.Form frmCredit 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDIT"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   Icon            =   "frmCredit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "click me twice"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   4320
      Left            =   1080
      Picture         =   "frmCredit.frx":0ECA
      Top             =   4200
      Width           =   5280
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   885
      Left            =   240
      OleObjectBlob   =   "frmCredit.frx":3EE6
      SourceDoc       =   "F:\credit\MORE...htm"
      TabIndex        =   0
      Top             =   720
      Width           =   1440
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   6240
      Picture         =   "frmCredit.frx":86FE
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6060
      Left            =   -240
      Picture         =   "frmCredit.frx":95C8
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

