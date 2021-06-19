VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Step1:"
   ClientHeight    =   6585
   ClientLeft      =   3045
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
      TabIndex        =   34
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame frIntro 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display at the Left Side Of The Installer (Recomended:- Image)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Top             =   3600
      Width           =   7215
      Begin VB.CheckBox chkBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tranceparent"
         Height          =   255
         Index           =   1
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   21
         Text            =   "Installing My App"
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optMsg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Image"
         Height          =   255
         Index           =   3
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   19
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton optCol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Background Color"
         Height          =   255
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optCol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Foreground Color"
         Height          =   255
         Index           =   3
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optMsg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Message"
         Height          =   255
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "approx 75 X 225"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   23
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next >>"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame frIntro 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome Splash Screen to be displayed on Installer at startup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   7215
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "5"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tranceparent"
         Height          =   255
         Index           =   0
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton optCol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Foreground Color"
         Height          =   255
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optCol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Background Color"
         Height          =   255
         Index           =   0
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton optMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Image"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Text            =   "Welcome To My App Installer"
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Message"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time(in sec):"
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
         Index           =   6
         Left            =   3480
         TabIndex        =   32
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "approx 360 X 330"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Text            =   "VINSOFT"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   3
      Text            =   "1.0"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Text            =   "My App"
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Step1>>    WELCOME SCREEN"
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
      Left            =   840
      TabIndex        =   27
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company Name"
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
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
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
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application Name"
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim notWhite As Boolean 'true when not white



Private Sub chkBack_Click(Index As Integer)
If chkBack(0).Value Then
 optCol(0).Value = False
  optCol(1).Value = False
 lblCol(0).BackColor = vbGreen
 optMsg(0).Value = True
ElseIf chkBack(1).Value Then
 optCol(2).Value = False
  optCol(3).Value = False
 lblCol(2).BackColor = &HDDBC20
 optMsg(2).Value = True
End If
End Sub

Private Sub cmdBrowse_Click(Index As Integer)

mdifrmMain.CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   mdifrmMain.CommonDialog1.fileName = ""
   mdifrmMain.CommonDialog1.InitDir = App.path & "\data\images"
   mdifrmMain.CommonDialog1.Filter = "*.jpg|*.jpg|*.bmp|*.bmp|*.wmf|*.wmf|*.gif|*.gif|All Files|*.*"
   mdifrmMain.CommonDialog1.ShowOpen
   If mdifrmMain.CommonDialog1.fileName = "" Or Err.Number = cdlCancel Then
      MsgBox "No file is opened"
      Exit Sub
   End If
   If Index = 0 Then
    txtMsg(1).Text = mdifrmMain.CommonDialog1.fileName
   Else
   txtMsg(3).Text = mdifrmMain.CommonDialog1.fileName
   End If

End Sub

Private Sub cmdBrowse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 cmdBrowse(Index).BackColor = &HE0E0E0
 notWhite = True
End If
End Sub

Private Sub cmdDir_Click()
frmButton.imgStepOver_Click (1)

End Sub

Private Sub cmdDir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 cmdDir.BackColor = &HE0E0E0
 notWhite = True
End If

End Sub

Private Sub cmdPreview_Click()

frmPrev.Visible = True
 frmPrev.step1
End Sub

Private Sub cmdPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 notWhite = True
  cmdPreview.BackColor = &HFF00FF
End If
End Sub





Private Sub Form_Load()
optMsg_Click (3)
optMsg(3).Value = True
Me.Picture = LoadPicture(App.path & "\data\back.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = True Then
 cmdDir.BackColor = vbWhite
 cmdPreview.BackColor = 16777088
 notWhite = False
 
End If
End Sub

Private Sub frIntro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = True Then
 cmdBrowse(0).BackColor = vbWhite
 cmdBrowse(1).BackColor = vbWhite
 notWhite = False
 End If
End Sub

Private Sub optCol_Click(Index As Integer)

If Index = 0 Then
  chkBack(0).Value = Unchecked
  optMsg(0).Value = True
End If
If Index = 2 Then
  chkBack(1).Value = Unchecked
  optMsg(2).Value = True
End If
If Index = 1 Then optMsg(0).Value = True
If Index = 3 Then optMsg(1).Value = True
Dim CDFlags As Long
Dim Rang As Long

On Error GoTo ColorError

    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)

    mdifrmMain.CommonDialog1.Flags = CDFlags
    mdifrmMain.CommonDialog1.CancelError = True
    mdifrmMain.CommonDialog1.ShowColor
    Rang& = mdifrmMain.CommonDialog1.Color      'obtained BGR color
    lblCol(Index).BackColor = Rang&
    isCompiled = False
    Exit Sub
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
    
    Else
        MsgBox "An error occured"
    End If


End Sub

Public Function MakeFile1() As String
Dim txtSave As String
txtSave = "<<<Starting Form>>>" & vbCrLf
txtSave = txtSave & txtInfo(0).Text & vbCrLf
txtSave = txtSave & txtInfo(1).Text & vbCrLf
txtSave = txtSave & txtInfo(2).Text & vbCrLf
txtSave = txtSave & " <Welcome>" & vbCrLf
If optMsg(0).Value Then
 txtSave = txtSave & "0" & vbCrLf
Else
 txtSave = txtSave & "1" & vbCrLf
End If
txtSave = txtSave & txtMsg(0).Text & vbCrLf
txtSave = txtSave & lblCol(0).BackColor & vbCrLf
If chkBack(0).Value Then
 txtSave = txtSave & "0" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
txtSave = txtSave & lblCol(1).BackColor & vbCrLf
txtSave = txtSave & txtMsg(1).Text & vbCrLf
If optMsg(2).Value Then
 txtSave = txtSave & "2" & vbCrLf
Else
 txtSave = txtSave & "3" & vbCrLf
End If
txtSave = txtSave & txtMsg(2).Text & vbCrLf
txtSave = txtSave & lblCol(2).BackColor & vbCrLf
If chkBack(1).Value Then
 txtSave = txtSave & "1" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If

txtSave = txtSave & lblCol(3).BackColor & vbCrLf
txtSave = txtSave & txtMsg(3).Text & vbCrLf
MakeFile1 = txtSave
End Function

Private Sub optMsg_Click(Index As Integer)

If Index = 1 Then
 optCol(0).Enabled = False
 optCol(1).Enabled = False
 chkBack(0).Enabled = False
 cmdBrowse(0).Enabled = True
 WelMessage = True
ElseIf Index = 3 Then
optCol(2).Enabled = False
 optCol(3).Enabled = False
 chkBack(1).Enabled = False
 cmdBrowse(1).Enabled = True
 DispMessage = True
ElseIf Index = 0 Then
 optCol(0).Enabled = True
 optCol(1).Enabled = True
 chkBack(0).Enabled = True
 cmdBrowse(0).Enabled = False
 WelMessage = False
ElseIf Index = 2 Then
chkBack(0).Enabled = True
 optCol(2).Enabled = True
 optCol(3).Enabled = True
 cmdBrowse(1).Enabled = False
 DispMessage = False
End If


End Sub

Private Sub txtInfo_Change(Index As Integer)
isCompiled = False
End Sub

Private Sub txtMsg_Change(Index As Integer)
isCompiled = False
End Sub

Private Sub txtTime_Change()
If IsNumeric(txtTime.Text) = False And Len(txtTime.Text) > 0 Then
 MsgBox "Please Enter any numeric value"
 txtTime.Text = ""
End If

End Sub
