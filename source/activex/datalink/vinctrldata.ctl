VERSION 5.00
Begin VB.UserControl vinctrldata 
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14610
   ScaleHeight     =   11520
   ScaleWidth      =   14610
   Begin VB.VScrollBar VScroll1 
      Height          =   11535
      Left            =   14040
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ARE YOU READY FOR THE EXAM"
      Height          =   1575
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT QUESTION"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox que 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Que"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "vinctrldata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim i As Integer

Private Sub Command1_Click()
 Data1.Recordset.MoveNext
 
End Sub

Private Sub Command2_Click()
  que(0).Visible = True
  For i = 1 To 15
  Load que(i)
  que(i).Top = que(i - 1).Top + 4000
  que(i).Visible = True
  que(i).Left = que(i - 1).Left
 Next
End Sub

Private Sub UserControl_Initialize()

End Sub

Private Sub VScroll1_Change()
 UserControl.Height = VScroll1.Value + 11520
 
End Sub
