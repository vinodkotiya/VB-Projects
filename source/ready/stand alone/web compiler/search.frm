VERSION 5.00
Begin VB.Form SearchForm 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VIN Web-Compiler Search & Replace"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "search.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   660
      Width           =   2100
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   195
      Width           =   2115
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Whole word only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   2040
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Case sensitive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2040
   End
   Begin VB.CommandButton ReplaceAllButton 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Replace All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   630
      Width           =   1275
   End
   Begin VB.CommandButton ReplaceButton 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Replace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton FindNextButton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Find Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton FindButton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Find"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Replace with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   660
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find what"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   150
      TabIndex        =   8
      Top             =   195
      Width           =   1410
   End
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Position As Currency

Private Sub FindButton_Click()
Dim FindFlags As Integer

    Position = 0
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = HTMLEdit.RichTextBox1.Find(Text1.Text, Position + 1, , FindFlags)
    If Position >= 0 Then
         HTMLEdit.SetFocus
    Else
        MsgBox "String " & Text1.Text & " not found"
        
    End If
    
End Sub

Private Sub FindNextButton_Click()
Dim FindFlags

FindFlags = Check1.Value * 4 + Check2.Value * 2
Position = HTMLEdit.RichTextBox1.Find(Text1.Text, Position + 1, , FindFlags)
If Position > 0 Then
    HTMLEdit.SetFocus
Else
    MsgBox "String " & Text1.Text & " not found"
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If

End Sub

Private Sub Command5_Click()

    SearchForm.Hide
    
End Sub

Private Sub Form_Load()
Text1.Text = HTMLEdit.RichTextBox1.SelText
End Sub

Private Sub Form_Unload(Cancel As Integer)
SearchForm.ReplaceButton.Enabled = False
SearchForm.ReplaceAllButton.Enabled = False
End Sub

Private Sub ReplaceButton_Click()
Dim FindFlags As Integer

    HTMLEdit.RichTextBox1.SelText = Text2.Text
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = HTMLEdit.RichTextBox1.Find(Text1.Text, Position + 1, , FindFlags)
    If Position > 0 Then
        HTMLEdit.SetFocus
    Else
        MsgBox "String " & Text1.Text & " not found"
       
    End If
    
End Sub

Private Sub ReplaceAllButton_Click()
Dim FindFlags As Integer
On Error GoTo chupchap
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    HTMLEdit.RichTextBox1.SelText = Text2.Text
    Position = HTMLEdit.RichTextBox1.Find(Text1.Text, Position + 1, , FindFlags)
    While Position > 0
        HTMLEdit.RichTextBox1.SelText = Text2.Text
        Position = HTMLEdit.RichTextBox1.Find(Text1.Text, Position + 1, , FindFlags)
        'position should be currency to prevent overflow in searching big files
    Wend
        
        MsgBox "Replacing of " & Text1.Text & " With " & Text2.Text & " is done "
        Exit Sub
chupchap:
    
End Sub

