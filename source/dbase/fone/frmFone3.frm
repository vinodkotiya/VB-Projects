VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Modify Records"
   ClientHeight    =   7620
   ClientLeft      =   330
   ClientTop       =   510
   ClientWidth     =   6390
   Icon            =   "frmFone3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H0080FF80&
      Caption         =   "Modify"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox Text 
      DataField       =   "Remarks"
      DataSource      =   "Datafone"
      Height          =   615
      Index           =   14
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5400
      Width           =   4815
   End
   Begin VB.TextBox Text 
      DataField       =   "Website"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   13
      Left            =   1320
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4920
      Width           =   4815
   End
   Begin VB.TextBox Text 
      DataField       =   "Dateofbirth"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   12
      Left            =   3120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text 
      DataField       =   "Email"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   11
      Left            =   1320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4410
      Width           =   4815
   End
   Begin VB.TextBox Text 
      DataField       =   "Mobile"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   10
      Left            =   4440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text 
      DataField       =   "FoneOffice"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   9
      Left            =   3240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text 
      DataField       =   "FoneResident"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   8
      Left            =   2040
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text 
      DataField       =   "Std"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text 
      DataField       =   "City"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text 
      DataField       =   "Area"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text 
      DataField       =   "Address"
      DataSource      =   "Datafone"
      Height          =   495
      Index           =   4
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox Text 
      DataField       =   "Designation"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text 
      DataField       =   "Surname"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      DataField       =   "Name"
      DataSource      =   "Datafone"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3600
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
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
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Data Datafone 
      Caption         =   " Directory"
      Connect         =   "Access"
      DatabaseName    =   ""
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
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      DataField       =   "Name"
      DataSource      =   "Datafone"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To Modify any record shown above click modify then Update it"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Index           =   3
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Website"
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
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DoBirth"
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
      Index           =   1
      Left            =   5040
      TabIndex        =   31
      Top             =   1920
      Width           =   975
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
      Left            =   120
      TabIndex        =   30
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
      TabIndex        =   29
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "To delete any record first display it Then Press Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   7200
      Width           =   5055
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To add a new record first click on Add and Then Enter data"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   6720
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
      TabIndex        =   26
      Top             =   4440
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
      TabIndex        =   25
      Top             =   4080
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
      TabIndex        =   24
      Top             =   4080
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
      TabIndex        =   23
      Top             =   4080
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
      TabIndex        =   22
      Top             =   4080
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
      Left            =   3240
      TabIndex        =   21
      Top             =   3000
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
      TabIndex        =   20
      Top             =   3000
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      Index           =   0
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
Dim ismodify  As Boolean     'modify records when true
Dim storTxtWhenDown As String 'temp store textdata when modify is not clicked and any one want to modify data by typing in text box

Private Sub cmbFields_Change()

End Sub



Private Sub cmdModify_Click()
Dim i As Integer
If ismodify = False Then
 ismodify = True
 cmdModify.Caption = "Update"
 For i = 0 To 14
   Text(i).TabStop = True
  Next
ElseIf ismodify = True Then
 ismodify = False
 cmdModify.Caption = "Modify"
  For i = 0 To 14
   Text(i).TabStop = False
  Next
End If
End Sub

Private Sub Command3_Click()
'Datafone.EditMode = DynaSet
Datafone.Recordset.MoveLast
Datafone.Recordset.AddNew
cmdModify_Click
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
'//now all are strings //////////////////
         'if to be searched is a integer value direct add it
  ' If frmStart.cmbFields.ListIndex = 4 Or frmStart.cmbFields.ListIndex = 5 Then
   ' seteq = "=" & frmStart.txtSearch.Text
         'if to be searched is a string value add ' ' to it
   'Else
    seteq = "= '" & frmStart.txtSearch.Text & "'"
    'seteq = "=" & seteq
   'End If
     GenerateSQL = frmStart.cmbFields.Text & " " & seteq
End If
    'GenerateSQL = frmStart.cmbFields.Text & " " & seteq
End Function

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub txtSearch_Change()

End Sub



Private Sub Form_Load()
ismodify = False
Datafone.DatabaseName = App.Path & "\data\DIRECTORY.mdb"
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If ismodify = False Then storTxtWhenDown = Text(Index).Text
End Sub


Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If ismodify = False Then Text(Index).Text = storTxtWhenDown
End Sub

Private Sub Text_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ismodify = False Then cmdModify.SetFocus
End Sub
