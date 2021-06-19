VERSION 5.00
Begin VB.Form frmSys 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Step 4:"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   3045
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFF80&
      Caption         =   "Preview"
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< Back"
      Height          =   375
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next >>"
      Height          =   375
      Index           =   1
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame frLnk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create Registry Entries"
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
      Height          =   2775
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   5775
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLink 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   10
         Text            =   "\company"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtLink 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   11
         Text            =   "vinsoft"
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create"
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         Height          =   255
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         Height          =   255
         Index           =   2
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ListBox listLink 
         Height          =   960
         ItemData        =   "frmSys.frx":0000
         Left            =   120
         List            =   "frmSys.frx":0002
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   1680
         Width           =   5535
      End
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         Height          =   255
         Index           =   3
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name of SUBKEY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "KeyValue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdGen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate"
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   840
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registration Code"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show System Administrators Company Name"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show System Administrators Name"
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
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Step4>>   SYSTEM INFORMATION"
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
      Left            =   480
      TabIndex        =   21
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim notWhite As Boolean 'true when not white
Dim editListno As Integer 'no of linklist which is in editing when edit set to -5

Private Sub chkSys_Click(Index As Integer)
If chkSys(0).Value Then
SysAd = True
ElseIf chkSys(0).Value = Unchecked Then
SysAd = False
End If
If chkSys(1).Value Then
SysCompany = True
ElseIf chkSys(1).Value = Unchecked Then
SysCompany = False
End If
If chkSys(2).Value Then
 RegCode = True
ElseIf chkSys(2).Value = Unchecked Then
 RegCode = False
End If
End Sub

Private Sub cmdDir_Click(Index As Integer)
If Index = 1 Then
 frmButton.imgStepOver_Click (4)
Else
frmButton.imgStepOver_Click (2)
End If

End Sub

Private Sub cmdDir_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 notWhite = True
 cmdDir(Index).BackColor = &HE0E0E0
 
End If
End Sub

Private Sub cmdGen_Click()
Dim digits As Long
Dim i As Integer
For i = 0 To 3
digits = Int(Rnd(Minute(Time) * Second(Time)) * 9999)
If digits < 1000 Then digits = Int(Rnd(300) * 9999)
If digits < 1000 Then digits = Int(Rnd(300) * 9999)

If 65 < (digits Mod 100) And 90 > (digits Mod 100) Then
 'MsgBox (digits Mod 100) & (digits / 10)
 txtReg(i).Text = Int(digits / 10) & Chr(digits Mod 100)
Else
 txtReg(i).Text = digits
End If
Next
End Sub

Private Sub cmdLink_Click(Index As Integer)
Dim pos As Integer
Dim posD As Integer
If Index = 0 Or Index = 3 Then
    If Index = 3 And editListno <> -5 Then
     listLink.RemoveItem (editListno)
     editListno = -5     'updated
    End If
    
  If "\" <> Left(Trim(txtLink(0).Text), 1) Then txtLink(0).Text = "\" & Trim(txtLink(0).Text)
 listLink.AddItem Combo1.Text & txtLink(0).Text & "  #  " & txtLink(1).Text
ElseIf Index = 1 Then
 pos = InStr(1, listLink.Text, "\")
 posD = InStr(1, listLink.Text, "#")
 txtLink(0).Text = Trim(Mid(listLink.Text, pos, posD - pos))
 txtLink(1).Text = Trim(Right(listLink.Text, Len(listLink.Text) - posD))
 If "D" = Left(listLink.Text, 1) Then
  Combo1.ListIndex = 0
 ElseIf "P" = Left(listLink.Text, 1) Then
  Combo1.ListIndex = 1
 ElseIf "P" = Left(listLink.Text, 1) Then
  Combo1.ListIndex = 2
 End If
 'listLink.RemoveItem (listLink.ListIndex)
 editListno = listLink.ListIndex
ElseIf Index = 2 Then
 listLink.RemoveItem (listLink.ListIndex)
 txtLink(0).Text = ""
 txtLink(1).Text = ""
End If
End Sub

Private Sub cmdPreview_Click()
frmPrev.Visible = True
 frmPrev.step4
End Sub
Private Sub cmdPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If notWhite = False Then
 notWhite = True
 cmdPreview.BackColor = &HFF00FF
 End If
End Sub

Private Sub Form_Load()
Combo1.AddItem "HKEY_CLASS_ROOT"
Combo1.AddItem "HKEY_CURRENT_USER"
Combo1.AddItem "HKEY_LOCAL_MACHINE"
Combo1.ListIndex = 2
Me.Picture = LoadPicture(App.path & "\data\back.jpg")
SysAd = True
SysCompany = True
RegCode = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = True Then
 cmdDir(0).BackColor = vbWhite
 cmdDir(1).BackColor = vbWhite
 cmdPreview.BackColor = 16777088
 notWhite = False
End If

End Sub

Private Sub listLink_Click()
Dim i As Integer
'MsgBox listLink.Text
For i = 0 To listLink.ListCount - 1
 If i <> listLink.ListIndex Then
  listLink.Selected(i) = False
 End If
 
Next
End Sub

Private Sub txtReg_Change(Index As Integer)
isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
If Len(txtReg(Index).Text) = 4 And frmSys.Visible Then
 If Index < 3 Then
  txtReg(Index + 1).SetFocus
 ElseIf Index = 3 Then
  cmdGen.SetFocus
 End If
End If
End Sub
Public Function MakeFile4() As String
Dim txtSave As String
Dim i As Integer
txtSave = "<<<System Information>>>" & vbCrLf
If chkSys(0).Value Then
 txtSave = txtSave & "0" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
If chkSys(1).Value Then
 txtSave = txtSave & "1" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
If chkSys(2).Value Then
 txtSave = txtSave & "2" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
txtSave = txtSave & txtReg(0).Text & vbCrLf
txtSave = txtSave & txtReg(1).Text & vbCrLf
txtSave = txtSave & txtReg(2).Text & vbCrLf
txtSave = txtSave & txtReg(3).Text & vbCrLf
txtSave = txtSave & " <ListLink>" & vbCrLf
For i = 0 To listLink.ListCount - 1
 txtSave = txtSave & listLink.List(i) & vbCrLf
Next
txtSave = txtSave & " </ListLink>" & vbCrLf
MakeFile4 = txtSave
End Function

