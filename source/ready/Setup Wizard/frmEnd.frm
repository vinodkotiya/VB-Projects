VERSION 5.00
Begin VB.Form frmEnd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Step5 :"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   3045
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFF80&
      Caption         =   "Preview"
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtRunBack 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4800
      TabIndex        =   29
      Text            =   "register.exe"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run This Program in background"
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
      Index           =   3
      Left            =   480
      TabIndex        =   28
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2400
      TabIndex        =   24
      Text            =   "C:\VIN Setups"
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< Back"
      Height          =   375
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FINISH..."
      Height          =   375
      Index           =   1
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame frIntro 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Finishing Splash Screen to be displayed on Installer after Installation"
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
      Left            =   480
      TabIndex        =   11
      Top             =   3360
      Width           =   7215
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "3"
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox chkBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tranceparent"
         Height          =   255
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
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
         TabIndex        =   22
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   21
         Top             =   480
         Width           =   2415
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
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   14
         Text            =   "Thanking You ........."
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optCol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Background Color"
         Height          =   255
         Index           =   0
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1455
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
      Begin VB.OptionButton optMsg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Message"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "approx 360 X 330"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   32
         Top             =   840
         Width           =   1335
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
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prompt to Reboot Computer After Installation"
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
      TabIndex        =   10
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Frame frEnd 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   6615
      Begin VB.TextBox txtTarget 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Text            =   "Myapp.exe"
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optChk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UnChecked"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optChk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Checked"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Source Dir\"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name of Application to be Launched"
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
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prompt to Launch Application  After Installation"
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
      TabIndex        =   4
      Top             =   1200
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.Frame frEnd 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      Begin VB.OptionButton optChk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UnChecked"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optChk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Checked"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prompt For Readme.txt After Installation"
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
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Source Dir\"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   34
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Specify Directory where setup file will be generated when BUILD"
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
      Left            =   480
      TabIndex        =   26
      Top             =   5040
      Width           =   5895
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Output Directory"
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
      Index           =   3
      Left            =   480
      TabIndex        =   25
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Step5>>    FINISHING SETUP"
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
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim notWhite As Boolean 'true when not white


Private Sub chkBack_Click()
If chkBack.Value Then
 optCol(0).Value = False
  optCol(1).Value = False
 lblCol(0).BackColor = vbGreen
End If
End Sub

Private Sub chkSys_Click(Index As Integer)
If Index = 2 Then
 chkSys(3).Value = Unchecked
ElseIf Index = 3 Then
 chkSys(2).Value = Unchecked
End If
'If chkSys(3).Value = 0 Then chkSys(2).Value = Checked

'If chkSys(2).Value = 0 Then chkSys(3).Value = Checked
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
     txtMsg(1).Text = mdifrmMain.CommonDialog1.fileName
End Sub

Private Sub cmdBrowse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 cmdBrowse(Index).BackColor = &HE0E0E0
 notWhite = True
End If
End Sub

Private Sub cmdDir_Click(Index As Integer)
If Index = 1 Then

frmSys.Visible = False
frmEnd.Visible = False
frmStart.Visible = False
frmAgree.Visible = False
frmAppl.Visible = False
Else
frmButton.imgStepOver_Click (3)
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
 frmPrev.step5
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
 cmdDir(0).BackColor = vbWhite
 cmdDir(1).BackColor = vbWhite
 cmdPreview.BackColor = &HFFFF80
  notWhite = False
End If

End Sub

Private Sub frIntro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = True Then
 cmdBrowse(0).BackColor = vbWhite
 
 notWhite = False
 End If
End Sub

Private Sub optCol_Click(Index As Integer)
If Index = 0 Then
 chkBack.Value = Unchecked
End If
 optMsg(0).Value = True
Dim CDFlags As Long
Dim Rang As Long

On Error GoTo ColorError

    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)

    mdifrmMain.CommonDialog1.Flags = CDFlags
    mdifrmMain.CommonDialog1.CancelError = True
    mdifrmMain.CommonDialog1.ShowColor
    Rang& = mdifrmMain.CommonDialog1.Color      'obtained BGR color
    lblCol(Index).BackColor = Rang&
    isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
    Exit Sub
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
    
    Else
        MsgBox "An error occured"
    End If
End Sub

Private Sub optMsg_Click(Index As Integer)
If Index = 1 Then
 optCol(0).Enabled = False
 optCol(1).Enabled = False
 chkBack.Enabled = False
 cmdBrowse(0).Enabled = True
 EndMessage = True
ElseIf Index = 0 Then
  optCol(0).Enabled = True
 optCol(1).Enabled = True
 chkBack.Enabled = True
  cmdBrowse(0).Enabled = False
  EndMessage = False
End If
End Sub
Public Function MakeFile5() As String
Dim txtSave As String
txtSave = "<<<Finishing Form>>>" & vbCrLf
If chkSys(0).Value Then
 txtSave = txtSave & "0" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
  If optChk(0).Value Then
   txtSave = txtSave & "0" & vbCrLf
  ElseIf optChk(1).Value Then
   txtSave = txtSave & "1" & vbCrLf
  End If
If chkSys(1).Value Then
 txtSave = txtSave & "1" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
  If optChk(2).Value Then
   txtSave = txtSave & "2" & vbCrLf
  ElseIf optChk(3).Value Then
   txtSave = txtSave & "3" & vbCrLf
  End If
txtSave = txtSave & txtTarget.Text & vbCrLf
If chkSys(2).Value Then 'reboot
 txtSave = txtSave & "2" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
If chkSys(3).Value Then 'background
 txtSave = txtSave & "3" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If
 txtSave = txtSave & txtRunBack.Text & vbCrLf
txtSave = txtSave & " <Alvida>" & vbCrLf
If optMsg(0).Value Then
 txtSave = txtSave & "0" & vbCrLf
Else
 txtSave = txtSave & "1" & vbCrLf
End If
txtSave = txtSave & txtMsg(0).Text & vbCrLf
txtSave = txtSave & lblCol(0).BackColor & vbCrLf
If chkBack.Value Then
 txtSave = txtSave & "0" & vbCrLf
Else
 txtSave = txtSave & "-1" & vbCrLf
End If

txtSave = txtSave & lblCol(1).BackColor & vbCrLf
txtSave = txtSave & txtMsg(1).Text & vbCrLf
txtSave = txtSave & txtOutput.Text & vbCrLf
MakeFile5 = txtSave
End Function

Private Sub txtMsg_Change(Index As Integer)
isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
End Sub

Private Sub txtRunBack_Change()
isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
End Sub

Private Sub txtTarget_Change()
isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
End Sub

Private Sub txtTime_Change()
If IsNumeric(txtTime.Text) = False And Len(txtTime.Text) > 0 Then
 MsgBox "Please Enter any numeric value"
 txtTime.Text = ""
End If
End Sub
