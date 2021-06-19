VERSION 5.00
Begin VB.Form frmAgree 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Step2 :"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   5760
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFF80&
      Caption         =   "Preview"
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next >>"
      Height          =   375
      Index           =   1
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< Back"
      Height          =   375
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtAgree 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   2
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4440
      Width           =   5655
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   615
   End
   Begin VB.CheckBox chkAgree 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display ReadMe After Installation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtAgree 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   1
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2640
      Width           =   5655
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox chkAgree 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display Lisence Agreement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtAgree 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   0
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   5655
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CheckBox chkAgree 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display Software Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Step2>>    TERMS and CONDITIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmAgree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim notWhite As Boolean 'true when not white

Private Sub chkAgree_Click(Index As Integer)
'If chkAgree(2).Value Then frmEnd.chkSys(0).Value = Checked
'If chkAgree(2).Value = 0 Then frmEnd.chkSys(0).Value = Unchecked
If chkAgree(Index).Value Then
 cmdBrowse(Index).Enabled = True
 txtAgree(Index).Enabled = True
Else
 cmdBrowse(Index).Enabled = False
 txtAgree(Index).Enabled = False
End If
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error GoTo vinerror
mdifrmMain.CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   mdifrmMain.CommonDialog1.fileName = ""
   mdifrmMain.CommonDialog1.InitDir = App.path & "\data"
   mdifrmMain.CommonDialog1.Filter = "Text Files|*.txt|All Files|*.*"
   mdifrmMain.CommonDialog1.ShowOpen
   If mdifrmMain.CommonDialog1.fileName = "" Then
      MsgBox "No file is opened"
      Exit Sub
   End If
  Dim fsys As New FileSystemObject
  Dim thisFile As File
  Set thisFile = fsys.GetFile(mdifrmMain.CommonDialog1.fileName)
  If thisFile.Size > 50000 Then
    MsgBox "File is too long to open..."
    Exit Sub
  End If
  Dim fnum As Integer
  Dim currentline As String
  Dim length As Long
  Dim noOfLoop As Long
  txtAgree(Index).Text = ""
  fnum = FreeFile
  
   Open mdifrmMain.CommonDialog1.fileName For Input As fnum      'dont use #1 for multiple file openings
   length = LOF(fnum)
   noOfLoop = length / 60
   length = 0
   While Not EOF(fnum)
     'Line Input #FNum, currentline
       length = length + 1
       If noOfLoop > length Then
         currentline = Input$(60, #fnum)
       Else
        Line Input #fnum, currentline
       End If
     txtAgree(Index).Text = txtAgree(Index).Text & currentline & vbCrLf
     DoEvents
   Wend
   Close #fnum
   Exit Sub
vinerror:
 If Err.Number = cdlCancel Then
  MsgBox "File not Opened"
  Exit Sub
 End If
  'MsgBox "File too long or any error occured"
End Sub

Private Sub cmdBrowse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 cmdBrowse(Index).BackColor = &HE0E0E0
 notWhite = True
End If
End Sub

Private Sub cmdDir_Click(Index As Integer)
If Index = 1 Then
frmButton.imgStepOver_Click (2)
Else
 frmButton.imgStepOver_Click (0)
End If
 
End Sub

Private Sub cmdDir_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 cmdDir(Index).BackColor = &HE0E0E0
 notWhite = True
End If
End Sub

Private Sub cmdPreview_Click()
frmPrev.Visible = True
 frmPrev.step2
End Sub

Private Sub cmdPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 notWhite = True
 cmdPreview.BackColor = &HFF00FF
End If
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.path & "\data\back.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = True Then
 cmdBrowse(0).BackColor = vbWhite
 cmdBrowse(1).BackColor = vbWhite
 cmdBrowse(2).BackColor = vbWhite
 cmdPreview.BackColor = 16777088
 cmdDir(0).BackColor = vbWhite
 cmdDir(1).BackColor = vbWhite
 notWhite = False
End If

End Sub
Public Function MakeFile2() As String
Dim txtSave As String
txtSave = "<<<Agreement>>>" & vbCrLf

If chkAgree(0).Value Then
 txtSave = txtSave & "0" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
If chkAgree(1).Value Then
 txtSave = txtSave & "1" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
If chkAgree(2).Value Then
 txtSave = txtSave & "2" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
txtSave = txtSave & " <Software Info>" & vbCrLf
txtSave = txtSave & txtAgree(0).Text & vbCrLf
txtSave = txtSave & " </Software Info>" & vbCrLf
txtSave = txtSave & " <Lisence>" & vbCrLf
txtSave = txtSave & txtAgree(1).Text & vbCrLf
txtSave = txtSave & " </Lisence>" & vbCrLf
txtSave = txtSave & " <Read Me>" & vbCrLf
txtSave = txtSave & txtAgree(2).Text & vbCrLf
txtSave = txtSave & " </Read Me>" & vbCrLf
MakeFile2 = txtSave
End Function

Private Sub txtAgree_Change(Index As Integer)
isCompiled = False
End Sub
