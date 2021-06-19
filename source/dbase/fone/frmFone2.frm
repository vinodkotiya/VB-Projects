VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Modify Records"
   ClientHeight    =   6735
   ClientLeft      =   330
   ClientTop       =   510
   ClientWidth     =   6540
   Icon            =   "frmFone2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6540
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Modify It"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text 
      DataField       =   "Remarks"
      DataSource      =   "Datafone"
      Height          =   615
      Index           =   14
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   4560
      Width           =   4815
   End
   Begin VB.TextBox Text 
      DataField       =   "Website"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   13
      Left            =   1320
      TabIndex        =   32
      Top             =   4080
      Width           =   4815
   End
   Begin VB.TextBox Text 
      DataField       =   "Dateofbirth"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   12
      Left            =   3120
      TabIndex        =   31
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text 
      DataField       =   "Email"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   11
      Left            =   1320
      TabIndex        =   30
      Top             =   3690
      Width           =   4815
   End
   Begin VB.TextBox Text 
      DataField       =   "Mobile"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   10
      Left            =   4440
      TabIndex        =   29
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text 
      DataField       =   "FoneOffice"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   9
      Left            =   3240
      TabIndex        =   28
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text 
      DataField       =   "FoneResident"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   27
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text 
      DataField       =   "Std"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   26
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text 
      DataField       =   "City"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   6
      Left            =   4080
      TabIndex        =   25
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text 
      DataField       =   "Area"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   24
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text 
      DataField       =   "Address"
      DataSource      =   "Datafone"
      Height          =   495
      Index           =   4
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      DataField       =   "Designation"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   22
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text 
      DataField       =   "Surname"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   21
      Top             =   1080
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
      TabIndex        =   20
      Top             =   1080
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Add New"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
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
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "Name"
      DataSource      =   "Datafone"
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   2
      X1              =   5880
      X2              =   6120
      Y1              =   360
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      Index           =   1
      X1              =   5520
      X2              =   2760
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To remove any Record from database click on Delete"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   36
      Top             =   6360
      Width           =   4215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To change record shown above click here"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   4935
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
      TabIndex        =   19
      Top             =   4560
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
      TabIndex        =   18
      Top             =   4080
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
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      Index           =   0
      X1              =   960
      X2              =   2760
      Y1              =   600
      Y2              =   360
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
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   0
      Width           =   3375
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
      Left            =   240
      TabIndex        =   15
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "To add a new record first click on Add and Then Enter data"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   5880
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
      Left            =   240
      TabIndex        =   13
      Top             =   3720
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
      TabIndex        =   12
      Top             =   3360
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
      TabIndex        =   11
      Top             =   3360
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
      TabIndex        =   10
      Top             =   3360
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
      TabIndex        =   9
      Top             =   3360
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
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
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
      TabIndex        =   7
      Top             =   2640
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
      TabIndex        =   6
      Top             =   1920
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
      TabIndex        =   5
      Top             =   1440
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
      TabIndex        =   4
      Top             =   1080
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
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ignore As Boolean  'no update if true

Private Sub cmbFields_Change()

End Sub


Private Sub Command1_Click()

If Command1.Caption = "Modify It" Then
    Ignore = False
    Command1.Caption = "Now Update It"
    Timer1.Interval = 1000
    Datafone.Enabled = False
    Label12(1).Caption = "When changes completes you must click here to update database."
Else
   Ignore = True
    Command1.Caption = "Modify It"
    Timer1.Interval = 0
    Command1.BackColor = &H80FF80
    Datafone.Enabled = True
    Label12(1).Caption = "To change record shown above click here"
End If
End Sub

Private Sub Command3_Click()
If Ignore = False Then
 MsgBox "Please First Click on 'Now Update It'"
 Exit Sub
End If

'Datafone.EditMode = DynaSet
Datafone.Recordset.MoveLast
Datafone.Recordset.AddNew
End Sub

Private Sub Command4_Click()
If Ignore = False Then
 MsgBox "Please First Click on 'Now Update It'"
 Exit Sub
End If

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
Ignore = True
End Sub

Private Sub Text_Change(Index As Integer)
'MsgBox Text(Index).DataField
Dim i As Integer
For i = 0 To 14
 If Text(i).Text <> "" Then
  Text(i).BackColor = &HFFFFC0
 Else
  Text(i).BackColor = vbWhite
 End If
Next
End Sub

Private Sub Text_Click(Index As Integer)
'If Ignore Then MsgBox "Please Click on 'Modify' to change records"
End Sub

Private Sub Text_GotFocus(Index As Integer)
Text_Click (Index)
End Sub

Private Sub Timer1_Timer()
If Command1.BackColor = &H80FF80 Then
   Command1.BackColor = vbWhite
Else
    Command1.BackColor = &H80FF80
End If
End Sub
