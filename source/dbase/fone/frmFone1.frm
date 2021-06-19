VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Modify Records"
   ClientHeight    =   6960
   ClientLeft      =   390
   ClientTop       =   345
   ClientWidth     =   6915
   Icon            =   "frmFone1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6915
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      DataField       =   "Emails"
      DataSource      =   "Datafone"
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Data Datafone 
      Caption         =   " Directory"
      Connect         =   "Access"
      DatabaseName    =   "D:\PROGRAME FILE\fone\data\DIRECTORY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Fone"
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox Text10 
      DataField       =   "Mobile"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      DataField       =   "FoneOffice"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      DataField       =   "FoneResident"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "STD"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text6 
      DataField       =   "City"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      DataField       =   "Area"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "Address"
      DataSource      =   "Datafone"
      Height          =   615
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      DataField       =   "Post"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "Surname"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "Name"
      DataSource      =   "Datafone"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6360
      X2              =   6000
      Y1              =   360
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   3480
      X2              =   5520
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   960
      X2              =   3360
      Y1              =   960
      Y2              =   720
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "If You want to go to first or last Record  Use Rightmost or Leftmost  Arrow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "If You want to go through the whole directory Use these Arrow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1440
      TabIndex        =   26
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "To delete any record first display it Then Press Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   5055
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To add a new record first click on Add and Then Enter data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Emails"
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
      Left            =   120
      TabIndex        =   23
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Height          =   375
      Left            =   4800
      TabIndex        =   22
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fone(O)"
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
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fone(R)"
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
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Std"
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
      Left            =   960
      TabIndex        =   19
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
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
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
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
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbFields_Change()

End Sub


Private Sub Command3_Click()
'Datafone.EditMode = DynaSet
Datafone.Recordset.MoveLast
Datafone.Recordset.AddNew
End Sub

Private Sub Command4_Click()
On Error Resume Next
 Datafone.RecordsetType = 1
    Datafone.Recordset.Delete
    If Not Datafone.Recordset.EOF Then
        Datafone.Recordset.MoveNext
    ElseIf Not Datafone.Recordset.BOF Then
        Datafone.Recordset.MovePrevious
    Else
        MsgBox "This was the last record in the table"
    End If
End Sub

Public Function GenerateSQL() As String
Dim seteq As String    'will add = sign before txtSearch
Dim kya As Byte
 kya = InStr(frmStart.txtSearch.Text, "*")
  
 If kya > 0 Then     'means * is found
  

  GenerateSQL = frmStart.cmbFields.Text & " LIKE '" & frmStart.txtSearch.Text & "'"
  'return from here
'if * is not found execute this
Else
         'if to be searched is a integer value direct add it
   If frmStart.cmbFields.ListIndex = 4 Or frmStart.cmbFields.ListIndex = 5 Then
    seteq = "=" & frmStart.txtSearch.Text
         'if to be searched is a string value add ' ' to it
   Else
    seteq = "= '" & frmStart.txtSearch.Text & "'"
    'seteq = "=" & seteq
   End If
     GenerateSQL = frmStart.cmbFields.Text & " " & seteq
End If
    'GenerateSQL = frmStart.cmbFields.Text & " " & seteq
End Function

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub txtSearch_Change()

End Sub



