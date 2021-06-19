VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAlarm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   5430
   ClientTop       =   270
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Index           =   3
      Left            =   0
      TabIndex        =   31
      Top             =   6600
      Width           =   4095
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Find Day"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "1982"
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   36
         Text            =   "May"
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   35
         Text            =   "24"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         Height          =   375
         Index           =   2
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   39
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00FF8080&
         Caption         =   "Day Finder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   2280
         TabIndex        =   34
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label 
         BackColor       =   &H00FF8080&
         Caption         =   "Date"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Index           =   2
      Left            =   0
      TabIndex        =   19
      Top             =   5160
      Width           =   4095
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   3720
         Top             =   480
      End
      Begin VB.TextBox txtTim 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   3
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   26
         Text            =   "000"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtTim 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   2
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "00"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtTim 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   1
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "00"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtTim 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   0
         Left            =   720
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "00"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdTimer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stop Timer"
         Height          =   375
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdTimer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Timer"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         Height          =   375
         Index           =   1
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFF80&
         Caption         =   " Hr    Min  Sec  mSec"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   28
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFF80&
         Caption         =   "TIMER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   27
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   3240
      Width           =   4095
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         Height          =   375
         Index           =   0
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtDisp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "frmAlarm.frx":0000
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3720
      Top             =   2280
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   30
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkExe 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Execute This Program"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "10"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "10"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "00"
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Text            =   "AM"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Text            =   "It's the time to shutdown PC"
         Top             =   840
         Width           =   3135
      End
      Begin VB.CheckBox chkPlay 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Play this sound"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CommandButton cmdSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Alarm"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel Alarm"
         Height          =   375
         Index           =   1
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Open File"
         Height          =   255
         Index           =   1
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Default"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Value           =   -1  'True
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3360
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ALARM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   9
         Left            =   2760
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hour Min  Sec"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Message"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alarm Time"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblStart As Double
Dim Tstring As String
Dim startTimer As Boolean ' timer is on when true
Dim onlyonce As Boolean 'to set on top for alarmmessageonce inside timer
Private Declare Function GetTickCount Lib "Kernel32" () As Long


Private Sub chkExe_Click()
'chkExe.Value = Not chkExe.Value
If chkExe.Value Then
  CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "exe files|*.exe"
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then
      txtMsg(2).Text = "No executable file is selected."
      Exit Sub
   End If
   exeFile = CommonDialog1.FileName '*.*.vin
   txtMsg(2).Text = exeFile
  runExe = True
 Else
  runExe = False
End If
  
End Sub

Private Sub chkPlay_Click()
'chkPlay.Value = Not chkPlay.Value
If chkPlay.Value Then
 isplaySound = True
Else
 isplaySound = False
End If
'MsgBox playSound
End Sub



Private Sub cmbDate_Change(Index As Integer)
If Val(cmbDate(0).Text) > 31 Or Val(cmbDate(0).Text) < 0 Then
  MsgBox "Invalid Date"
  cmbDate(0).Text = "1"
End If
'If Val(cmbDate(1).Text) > 12 Or Val(cmbDate(1).Text) < 0 Then MsgBox "Invalid Month"
End Sub

Private Sub cmdExit_Click(Index As Integer)
'Me.Hide
Unload Me
frmTime.Show
End Sub

Private Sub cmdExit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit(Index).BackColor = &HE0E0E0
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
 Dim daynum As Integer
 Dim monthnum As Integer
 Dim yearnum As Integer
 Dim maxday As Integer
 Dim datetofind As String
 Dim wkdy As String
 Dim wkday As Integer
 'cmbDate(0).ListIndex = Val(cmbDate(0).Text) - 1
 daynum = Val(cmbDate(0).Text)
 ' daynum = cmbDate(0).ListIndex + 1
  monthnum = cmbDate(1).ListIndex + 1
  yearnum = Val(Right$(txtYear.Text, 2))
If monthnum = 1 Then maxday = 31
If monthnum = 2 And yearnum / 4 = Int(yearnum / 4) Then maxday = 29 Else maxday = 28
If monthnum = 3 Then maxday = 31
If monthnum = 4 Then maxday = 30
If monthnum = 5 Then maxday = 31
If monthnum = 6 Then maxday = 30
If monthnum = 7 Then maxday = 31
If monthnum = 8 Then maxday = 31
If monthnum = 9 Then maxday = 30
If monthnum = 10 Then maxday = 31
If monthnum = 11 Then maxday = 31
If monthnum = 12 Then maxday = 31

If daynum > maxday Then
   daynum = maxday
  cmbDate(0).ListIndex = maxday - 1
End If
datetofind = DateSerial(Val(txtYear.Text), monthnum, daynum)
 ' & " " &  & " " &
'Let Text4.Text = Datetofind
wkday = Weekday(datetofind, vbSunday)
If wkday = 1 Then wkdy$ = "Sunday"
If wkday = 2 Then wkdy$ = "Monday"
If wkday = 3 Then wkdy$ = "Tuesday"
If wkday = 4 Then wkdy$ = "Wednesday"
If wkday = 5 Then wkdy$ = "Thursday"
If wkday = 6 Then wkdy$ = "Friday"
If wkday = 7 Then wkdy$ = "Saturday"

Label(8).Caption = wkdy$

End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdFind.BackColor = &HE0E0E0
End Sub

Private Sub cmdSet_Click(Index As Integer)
If Index = 0 Then
 If Combo1.ListIndex = 0 Then
  AlarmHr = txtTime(0).Text
  Else
   AlarmHr = txtTime(0).Text + 12
  End If
 AlarmMin = txtTime(1).Text
 AlarmSec = txtTime(2).Text
 alarmMsg = txtMsg(0).Text 'transfer message
 Alarm = True
Else
 Alarm = False
End If
Timer1.Interval = 0
'Me.Hide
Unload Me
frmTime.Show

End Sub



Private Sub cmdSet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSet(Index).BackColor = &HE0E0E0
End Sub

Private Sub cmdTimer_Click(Index As Integer)
If Index = 0 Then
 startTimer = True
 dblStart = GetTickCount
Else
 If cmdTimer(Index).Caption = "Stop Timer" Then
   cmdTimer(Index).Caption = "Resume Timer"
   startTimer = False
 Else
   cmdTimer(Index).Caption = "Stop Timer"
   startTimer = True
 End If
End If
End Sub

Private Sub cmdTimer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdTimer(Index).BackColor = &HE0E0E0
End Sub

Private Sub Form_Load()
Combo1.AddItem "AM"
Combo1.AddItem "PM"

  If Hour(Time) > 12 Then
   Combo1.ListIndex = 1
   txtTime(0).Text = Hour(Time) - 12
  Else
   Combo1.ListIndex = 0
   txtTime(0).Text = Hour(Time)
  End If
  If Minute(Time) + 5 < 60 Then
   txtTime(1).Text = Minute(Time) + 5
  Else
   txtTime(1).Text = Minute(Time)
  End If

isplaySound = True
Timer1.Interval = 500
txtDisp.Text = alarmMsg
onlyonce = True
Dim i As Integer
For i = 1 To 31
 cmbDate(0).AddItem i
Next
 cmbDate(0).ListIndex = 23
 cmbDate(1).AddItem "Jan"
 cmbDate(1).AddItem "Feb"
 cmbDate(1).AddItem "Mar"
 cmbDate(1).AddItem "Apr"
 cmbDate(1).AddItem "May"
 cmbDate(1).AddItem "Jun"
 cmbDate(1).AddItem "Jul"
 cmbDate(1).AddItem "Aug"
 cmbDate(1).AddItem "Sep"
 cmbDate(1).AddItem "Oct"
 cmbDate(1).AddItem "Nov"
 cmbDate(1).AddItem "Dec"
 cmbDate(1).ListIndex = 4
End Sub

Private Sub Frame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSet(0).BackColor = vbWhite
cmdSet(1).BackColor = vbWhite
cmdExit(0).BackColor = vbWhite
cmdExit(1).BackColor = vbWhite
cmdExit(2).BackColor = vbWhite
cmdTimer(0).BackColor = vbWhite
cmdTimer(1).BackColor = vbWhite
cmdFind.BackColor = vbWhite
End Sub

Private Sub opt_Click(Index As Integer)
If Index = 1 Then
  CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "wav files|*.wav|midi files|*.mid"
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then
      txtMsg(2).Text = "No file is selected for merging"
      Exit Sub
   End If
   playFile = CommonDialog1.FileName '*.*.vin
Else
 playFile = App.Path & "\data\alarm.mid"
End If
   txtMsg(1).Text = playFile
   
End Sub

Private Sub Timer1_Timer()
Label(3).Caption = Time
'any one if will execute b/c one of the 3 frame is visible at a time
If Frame(0).Visible And onlyonce = True Then
  opt_Click (0)
  onlyonce = False
End If

If Frame(1).Visible Then
 If txtDisp.BackColor = vbWhite Then
  txtDisp.BackColor = &HFF80FF
 Else
  txtDisp.BackColor = vbWhite
 End If
 If onlyonce = True Then
  If Screen.Width > 15000 Then        'for 1024X 768
       SetWindowPos Me.hwnd, HWND_TOPMOST, 350, 270, 273, 125, SWP_SHOWWINDOW
  Else                                                    '800 X 600
       SetWindowPos Me.hwnd, HWND_TOPMOST, 260, 210, 273, 125, SWP_SHOWWINDOW
  End If
  onlyonce = False
 End If
End If
End Sub

Private Sub Timer2_Timer()
'Label(4).Caption = Time
If startTimer Then
  Dim tmin, tsec, thour, thour2, tmin2, tsec2, length As Integer
  Tstring = Str(GetTickCount - dblStart)
  txtTim(3).Text = Right(Tstring, 3)
  length = Len(Str(Val(Tstring)))
  Tstring = Right(Tstring, length)
  If length > 4 Then
   tsec = Val(Left(Tstring, (length - 3)))
  End If

  tmin = Int(tsec / 60)
  thour = Int(tsec / 3600)
  tsec2 = tsec: If tsec > 59 Then tsec2 = tsec2 - Int(tsec / 60) * 60
  tmin2 = tmin: If tmin > 59 Then tmin2 = tmin2 - Int(tmin / 60) * 60
  thour2 = thour: If thour > 11 Then thour2 = thour2 - Int(thour / 12) * 12

  txtTim(2).Text = tsec2
  txtTim(1).Text = tmin2
  txtTim(0).Text = thour2
End If
End Sub

Private Sub txtTime_Change(Index As Integer)
If Index = 0 Then
 If Not IsNumeric(txtTime(Index).Text) Or Val(txtTime(Index).Text) < 0 Or Val(txtTime(Index).Text) > 24 Then
    MsgBox "Only Enter values between 0 to 24 "
    txtTime(Index).Text = "14"
 End If
Else
 If Not IsNumeric(txtTime(Index).Text) Or Val(txtTime(Index).Text) < 0 Or Val(txtTime(Index).Text) > 59 Then
  MsgBox "Only Enter values between 0 to 59 "
  txtTime(Index).Text = "24"
 End If
End If
End Sub

Private Sub txtYear_Click()
txtYear.SelStart = 0
txtYear.SelLength = Len(txtYear)

End Sub
