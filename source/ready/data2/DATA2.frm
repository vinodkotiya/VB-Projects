VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Last"
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Previous"
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "Comments"
      DataSource      =   "Data1"
      Height          =   1095
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "DATA2.frx":0000
      Top             =   3720
      Width           =   4935
   End
   Begin VB.TextBox Text4 
      DataField       =   "Subject"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      DataField       =   "Description"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "DATA2.frx":0006
      Top             =   360
      Width           =   4335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   5880
      Width           =   6255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "COMMENTS"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "SUBJECT"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DISCRIPTION"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ISBN"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TITLE"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then
 MsgBox "You are on the First record"
 Data1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
 MsgBox "You are on the last record"
 Data1.Recordset.MoveLast
End If
End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveLast
End Sub
