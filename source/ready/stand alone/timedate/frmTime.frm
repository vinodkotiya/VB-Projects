VERSION 5.00
Begin VB.Form frmTime 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "vinod"
   ClientHeight    =   2220
   ClientLeft      =   495
   ClientTop       =   825
   ClientWidth     =   2325
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picUp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2000
      Picture         =   "frmTime.frx":1CCA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   550
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "SET"
      Height          =   855
      Left            =   1800
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox saal 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   "yy"
      ToolTipText     =   "Year eg. 1982"
      Top             =   1800
      Width           =   630
   End
   Begin VB.TextBox maah 
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Text            =   "mm"
      ToolTipText     =   "Month eg. 05"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox din 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "dd"
      ToolTipText     =   "Day eg 24"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtss 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "ss"
      ToolTipText     =   "Seconds eg. 46"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtmm 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Text            =   "mm"
      ToolTipText     =   "Minute eg. 13"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txthh 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "hh"
      ToolTipText     =   "Hour eg. 22"
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox picDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2000
      Picture         =   "frmTime.frx":20AB
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      ToolTipText     =   "Set the system time and date"
      Top             =   550
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1920
      Top             =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      Height          =   765
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   270
   End
   Begin VB.PictureBox red 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Picture         =   "frmTime.frx":248C
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   2
      ToolTipText     =   "Click on red  button to  Display on top of all windows"
      Top             =   540
      Width           =   495
   End
   Begin VB.PictureBox green 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Picture         =   "frmTime.frx":282C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      ToolTipText     =   "Click on green button to  toggle Display on top of all windows"
      Top             =   540
      Width           =   255
   End
   Begin VB.Label txtDate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label txtTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loading......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Change System Date"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "  Change System Time"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Menu adrishya 
      Caption         =   "adrishya"
      Visible         =   0   'False
      Begin VB.Menu speak 
         Caption         =   "&Speak Time"
      End
      Begin VB.Menu settime 
         Caption         =   "Set System &Time"
      End
      Begin VB.Menu settop 
         Caption         =   "Set on top of all &windows"
         Checked         =   -1  'True
      End
      Begin VB.Menu trans 
         Caption         =   "Set Transparency"
         Begin VB.Menu smnuTrans 
            Caption         =   "25%"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu smnuTrans 
            Caption         =   "50%"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu smnuTrans 
            Caption         =   "75%"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu smnuTrans 
            Caption         =   "Solid"
            Checked         =   -1  'True
            Index           =   3
         End
      End
      Begin VB.Menu tick 
         Caption         =   "T&ick-Tick"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAlarm 
         Caption         =   "Set &Alarm"
      End
      Begin VB.Menu mnuTimer 
         Caption         =   "Ti&mer"
      End
      Begin VB.Menu mnuFinder 
         Caption         =   "&Day Finder"
      End
      Begin VB.Menu ew 
         Caption         =   "-"
      End
      Begin VB.Menu vinkit 
         Caption         =   "VIN Utility Kit"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu mnucredit 
         Caption         =   "Credit"
      End
      Begin VB.Menu sd 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
'///////////////// VIN Talking Clock v1.0 ///////////////////////////////////
'//////////        Created By : - VINOD KOTIYA             ///////////////////////////
'/////////          free on http://vinodkotiya.tripod.com ////////////////////
'/////////          help to promote the site if you want     ///////////////////////////////
'/////////          tell to your friend.         ///////////////////////////////////////
'/////////          provide advertizer or online job ///////
'/////////          Proudly releasing version 1.0 ////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////
Option Explicit
Dim scroll As String    'used scrolling text
Dim madhya As Integer     'contain splitted scroll no
Dim tickS As Boolean  'true when menu tick checked
Dim oldSecond As Integer

''SPEAKUP
Dim errorCode As Integer
Dim returnStr As String * 255
Dim cmd As String * 255

Private Declare Function mciSendString Lib "winmm.dll" _
    Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long
Dim speakonce As Boolean     'speak when true
Dim checkminonce As Boolean   'check is min = 00 once time when true
'prevent speaking many times when min = 00
'tranceparency
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
'sound
Private Declare Function playSound Lib "winmm.dll" Alias _
    "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, _
        ByVal dwFlags As Long) As Long



Private Sub Check1_Click()
'Dim X As Long
'Dim Y As Long
'X = frmTime.ScaleX(frmTime.Top, vbTwips, vbPixels)
'Y = frmTime.ScaleY(frmTime.Left, vbTwips, vbPixels)
On Error Resume Next

End Sub

Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

End Sub

Private Sub mnuabout_Click()
On Error Resume Next
MsgBox "ABOUT" & Chr(13) & _
"*************************************" & Chr(13) & Chr(13) & _
"VIN Talking Clock v1.0 " & Chr(13) & "             is a part of VIN UTILITY KIT " & Chr(13) _
& "Created By VINOD KOTIYA " & Chr(13) & "Created On 10-March-2003." & Chr(13) & _
"This version only speakup the hour because it is not possible" & Chr(13) & " to export full functionality in a floppy disk. " _
 & Chr(13) & _
"VIN Talking Clock is also available stand alone with complete time speaking capability in every minute."

End Sub

Private Sub close_Click()
End
End Sub

Private Sub mnucredit_Click()
On Error Resume Next
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\credit.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "///////////////// VIN Talking Clock v1.0 ///////////////////////////////////" & vbCrLf & _
 "//////////        Created By : - VINOD KOTIYA             ///////////////////////////" & vbCrLf & _
"/////////          free on http://vinodkotiya.tripod.com ////////////////////"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then PopupMenu adrishya
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
cmd = "close vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
speakonclick
 cmd = "close vinmid"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
Unload frmAlarm
End Sub

Private Sub green_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
PopupMenu adrishya

End Sub

Private Sub mnuAlarm_Click()
On Error Resume Next
Load frmAlarm ' make hidden
frmTime.Hide
frmAlarm.Frame(0).Visible = True
frmAlarm.Frame(1).Visible = False
frmAlarm.Frame(2).Visible = False
frmAlarm.Frame(3).Visible = False
frmAlarm.Frame(0).Top = frmAlarm.ScaleTop
frmAlarm.Height = frmAlarm.Frame(0).Top + frmAlarm.Frame(0).Height
frmAlarm.Show
End Sub

Private Sub mnuFinder_Click()
On Error Resume Next
frmTime.Hide
Load frmAlarm ' make hidden
frmAlarm.Frame(0).Visible = False
frmAlarm.Frame(1).Visible = False
frmAlarm.Frame(3).Visible = True
frmAlarm.Frame(2).Visible = False
frmAlarm.Frame(3).Top = frmAlarm.ScaleTop
frmAlarm.Height = frmAlarm.Frame(3).Top + frmAlarm.Frame(3).Height
frmAlarm.Show
End Sub

Private Sub mnuTimer_Click()
On Error Resume Next
frmTime.Hide
Load frmAlarm ' make hidden
frmAlarm.Frame(0).Visible = False
frmAlarm.Frame(1).Visible = False
frmAlarm.Frame(2).Visible = True
frmAlarm.Frame(3).Visible = False
frmAlarm.Frame(2).Top = frmAlarm.ScaleTop
frmAlarm.Height = frmAlarm.Frame(2).Top + frmAlarm.Frame(2).Height
frmAlarm.Timer1.Interval = 250
frmAlarm.Show
End Sub

Private Sub picDown_Click()
On Error Resume Next
picUp.Visible = True
picDown.Visible = False
frmTime.Height = frmTime.Height + 1500
txthh.Text = Str(Hour(Time))
txtmm.Text = Str(Minute(Time))
txtss.Text = Str(Second(Time))
din.Text = Str(Day(Date))
maah.Text = Str(Month(Date))
saal.Text = Str(Year(Date))

'speak time
speakonclick

End Sub

Private Sub Command1_Click()
On Error GoTo vinter
Time = txthh.Text & ":" & txtmm.Text & ":" & txtss.Text
Date = din.Text & "-" & maah.Text & "-" & saal.Text
frmTime.Height = frmTime.Height - 1500
picUp.Visible = False
picDown.Visible = True

'speak time
speakonclick
Exit Sub
vinter:
MsgBox "Invalid data entered "
txthh.Text = Str(Hour(Time))
txtmm.Text = Str(Minute(Time))
txtss.Text = Str(Second(Time))
din.Text = Str(Day(Date))
maah.Text = Str(Month(Date))
saal.Text = Str(Year(Date))

settime.Visible = True

End Sub

Private Sub Form_Load()
' init global variable of form

On Error Resume Next
madhya = 1
scroll = "     VIN Talking Clock : by Vinod Kotiya    ************"
speakonce = True
checkminonce = True
tickS = True
tick.Checked = True
'smnuTrans(0).Checked = False
'smnuTrans(2).Checked = False
'smnuTrans(3).Checked = False

''set topmost
red.Visible = False
green.Visible = True



Dim retValue As Long
    'Load Form1
    retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, 30, 30, _
               158, 80, SWP_SHOWWINDOW)
If Environ("OS") = "Windows_NT" Then smnuTrans_Click (2)
End Sub





Private Sub green_Click()
On Error Resume Next
green.Visible = False
red.Visible = True
   Dim reetValue As Long
   
    reetValue = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 30, 30, _
              158, 80, SWP_SHOWWINDOW)
              
'speak time
speakonclick

End Sub

Private Sub picUp_Click()
On Error Resume Next
picUp.Visible = False
picDown.Visible = True
frmTime.Height = frmTime.Height - 1500
'speak time
speakonclick

End Sub

Private Sub red_Click()
On Error Resume Next
red.Visible = False
green.Visible = True

Dim retValue As Long
    'Load Form1
    retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, 30, 30, _
               158, 80, SWP_SHOWWINDOW)
'speak time
    speakonclick
End Sub

Private Sub red_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 PopupMenu adrishya

End Sub

Private Sub settime_Click()
On Error Resume Next
picDown_Click
settime.Visible = False
End Sub

Private Sub settop_Click()
On Error Resume Next
settop.Checked = Not settop.Checked
If settop.Checked = True Then
   red_Click
ElseIf settop.Checked = False Then
   green_Click
End If
End Sub

Private Sub smnuTrans_Click(Index As Integer)
On Error Resume Next
If Environ("OS") = "Windows_NT" Then

Dim Level As Byte
If Index = 0 Then
  Level = 64
ElseIf Index = 1 Then
  Level = 128
ElseIf Index = 2 Then
  Level = 192
ElseIf Index = 3 Then
  Level = 255
End If
smnuTrans(Index).Checked = vbChecked
Dim i As Integer
For i = 0 To 3
 If i <> Index Then smnuTrans(i).Checked = vbUnchecked
Next
Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED)
Call GetWindowLong(Me.hwnd, GWL_EXSTYLE)
Call SetLayeredWindowAttributes(Me.hwnd, 0, Level, LWA_ALPHA)
Else
 MsgBox "Tranceparency Feature is only available for windows XP/NT"
End If
End Sub

Private Sub speak_Click()
On Error Resume Next
speakonclick

End Sub



Private Sub tick_Click()
On Error Resume Next
tick.Checked = Not tick.Checked
If tick.Checked = True Then
   tickS = True
ElseIf tick.Checked = False Then
   tickS = False
End If
End Sub


Private Sub Timer1_Timer()
'On Error GoTo skiperror

On Error Resume Next
Text1.Text = "v"
txtTime.Caption = Time
txtDate.Caption = Date
If picDown.Visible = True Then
  Text1.SetFocus
End If

If checkminonce = False And Minute(Time) = 1 Then
 checkminonce = True
End If

If speakonce = False Then 'enters once when min = 00 and make speakonce true
Dim min As Integer
  min = Minute(Time)
  If min = 0 And checkminonce = True Then
    speakonce = True
    checkminonce = False    'will became true whem min != 0
  End If
End If

'//// show scrolling
frmTime.Caption = Mid$(scroll, madhya, Len(scroll) - madhya)
 frmTime.Caption = frmTime.Caption & Mid$(scroll, 1, madhya) 'temp
 madhya = madhya + 1
 If madhya > Len(scroll) Then
  madhya = 1
 End If
 '/////////////////////
 DoEvents
 
If speakonce = True Then
Dim hr As Integer
hr = Hour(Time)
 If hr > 12 And hr < 24 Then
   hr = hr - 12
 End If
SpeakIt (hr)
speakonce = False
End If
 
''tick tick


If oldSecond <> Second(Time) And tickS = True Then
    oldSecond = Second(Time)
    hr = PlayWaveFile(App.Path & "\data\" & "tick.wav", True)  'hr used temporary
    
 End If
 
 '' alarm
 If Alarm = True Then
  If AlarmHr = Hour(Time) And AlarmMin = Minute(Time) And AlarmSec = Second(Time) Then
  
  If "WAV" = UCase(Right(playFile, 3)) And isplaySound = True Then
    tickS = False 'stop tick
    tick.Checked = False
    hr = frmTime.PlayWaveFile(playFile, True)
  ElseIf "MID" = UCase(Right(playFile, 3)) And isplaySound = True Then
    cmd = "close vinmid"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    ' now open the DAYS.WAV file as DAYS
    cmd = "open " & Chr(34) & playFile & Chr(34) & " type sequencer alias vinmid"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    'play the song
    errorCode = mciSendString("play vinmid", returnStr, 255, 0)
  End If
  'MsgBox frmAlarm.txtMsg(0).Text
 Unload frmAlarm
  Load frmAlarm
  Me.Hide
  frmAlarm.Frame(0).Visible = False
  frmAlarm.Frame(1).Visible = True
  frmAlarm.Frame(2).Visible = False
  frmAlarm.Frame(3).Visible = False
  frmAlarm.Frame(1).Top = frmAlarm.ScaleTop
  frmAlarm.Height = frmAlarm.Frame(1).Top + frmAlarm.Frame(1).Height
  frmAlarm.Show
  
  If runExe Then  'execute program
   Dim pos As Long
   Dim strDir As String
     pos = InStrRev(exeFile, "\", -1, vbBinaryCompare)
     strDir = Left(exeFile, pos)
     pos = ShellExecute(Me.hwnd, vbNullString, Right(exeFile, Len(exeFile) - pos), vbNullString, strDir, SW_SHOWNORMAL)
  End If
  Alarm = False
  End If
End If
 
 Exit Sub
skiperror:  'if any error occured skip it chupchap


End Sub
Public Function PlayWaveFile(strFileName As String, _
    Optional blnAsync As Boolean) As Boolean
    On Error Resume Next
    Dim lngFlags As Long
    Const snd_sync = &H0
    Const snd_Async = &H1
    Const snd_Nodefault = &H2
    Const snd_Filename = &H20000
    lngFlags = snd_Nodefault Or snd_Filename Or snd_sync
    If blnAsync Then lngFlags = lngFlags Or snd_Async
    PlayWaveFile = playSound(strFileName, 0&, lngFlags)
End Function

Private Sub Timer2_Timer()
'If ballred.Visible = True Then
' ballmag.Visible = True
'ElseIf ballmag.Visible = True Then
' ballred.Visible = True
'End If
On Error Resume Next
End Sub

    




Private Sub SpeakIt(Index As Integer)
'Dim errorCode As Integer    getting  error so i put them in option explicit
'Dim returnStr As Integer
'Dim cmd As String * 255
    On Error Resume Next
    ' make sure that device with the vin alias is open
    cmd = "close vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    ' now open the vin.WAV file as vin
    cmd = "open " & Chr(34) & App.Path & "\data\num.wav " & Chr(34) & " type waveaudio alias vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    If errorCode <> 0 Then
        MsgBox "There was an error on opening the num.WAV file." & vbCrLf _
               & "Please make sure the num.WAV file in the same folder as the application"
        Exit Sub
    End If
    Select Case Index
        Case 1: errorCode = mciSendString("play vin from 0 to 370 wait", returnStr, 255, 0)
        Case 2: errorCode = mciSendString("play vin from 370 to 690 wait", returnStr, 255, 0)
        Case 3: errorCode = mciSendString("play vin from 680 to 1000 wait", returnStr, 255, 0)
        Case 4: errorCode = mciSendString("play vin from 970 to 1200 wait", returnStr, 255, 0)
        Case 5: errorCode = mciSendString("play vin from 1200 to 1500 wait", returnStr, 255, 0)
        Case 6: errorCode = mciSendString("play vin from 1500 to 1900 wait", returnStr, 255, 0)
        Case 7: errorCode = mciSendString("play vin from 1900 to 2380 wait", returnStr, 255, 0)
        Case 8: errorCode = mciSendString("play vin from 2350 to 2700 wait", returnStr, 255, 0)
        Case 9: errorCode = mciSendString("play vin from 2670 to 3070 wait", returnStr, 255, 0)
        Case 10: errorCode = mciSendString("play vin from 3050 to 3390 wait", returnStr, 255, 0)
        Case 11: errorCode = mciSendString("play vin from 3370 to 3860 wait", returnStr, 255, 0)
        Case 12: errorCode = mciSendString("play vin from 3840 to 4370 wait", returnStr, 255, 0)
    End Select
     errorCode = mciSendString("play vin from 4340", returnStr, 255, 0)
End Sub


Private Sub speakonclick()
On Error Resume Next
Dim hr As Integer
hr = Hour(Time)
 If hr > 12 And hr < 24 Then
   hr = hr - 12
 End If
SpeakIt (hr)

End Sub




Private Sub txtDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then PopupMenu adrishya

End Sub



Private Sub txtTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then PopupMenu adrishya

End Sub

Private Sub vinkit_Click()

On Error Resume Next
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\vin_utility.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'vinutility.EXE' is not found in its " _
  & "Default directory  "

End Sub
