VERSION 5.00
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN UTILITY KIT"
   ClientHeight    =   4845
   ClientLeft      =   465
   ClientTop       =   1365
   ClientWidth     =   7785
   FillColor       =   &H8000000D&
   FillStyle       =   4  'Upward Diagonal
   ForeColor       =   &H8000000E&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   4845
   ScaleWidth      =   7785
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0080FF80&
      Caption         =   "I"
      Height          =   315
      Index           =   2
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Troubleshooter"
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0080FF80&
      Caption         =   "N"
      Height          =   315
      Index           =   1
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "About"
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "VIN CONVERT CENTRE"
      Height          =   495
      Index           =   8
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Before Clicking Make Sure That VinConvertCentre Is Not Currently Running."
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "Talking Clock"
      Height          =   495
      Index           =   7
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Before Clicking Make Sure That VinTimer Is Not Currently Running."
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "File Split and Merge"
      Height          =   495
      Index           =   6
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Before Clicking Make Sure That CALVIN Is Not Currently Running."
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "HAVE A BREAK !!"
      Height          =   495
      Index           =   5
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Before Clicking Make Sure That Have A Break Is Not Currently Running."
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "REMIND ME LATER"
      Height          =   495
      Index           =   4
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Before Clicking Make Sure That Remind Me Later Is Not Currently Running."
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "INDRADHANUSH"
      Height          =   495
      Index           =   3
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Before Clicking Make Sure That Indradhanush Is Not Currently Running."
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "FONE DIRECTORY"
      Height          =   495
      Index           =   2
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Before Clicking Make Sure That VinFoneDirectory Is Not Currently Running."
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "ALL2HTML CONVERTER"
      Height          =   495
      Index           =   1
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Before Clicking Make Sure That ALL2HTML Converter Is Not Currently Running."
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdAppl 
      BackColor       =   &H00FFFF00&
      Caption         =   "VIN WEB COMPILER"
      Height          =   495
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Before Clicking Make Sure That VinWebCompiler Is Not Currently Running."
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFF80&
      Caption         =   "DHUN"
      Height          =   255
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0080FF80&
      Caption         =   "V"
      Height          =   315
      Index           =   0
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Credit"
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer3 
      Left            =   6480
      Top             =   4440
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   5880
      Top             =   4440
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "VINCANOID"
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "AIRFORCE"
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "VINTRAP"
      Height          =   255
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5400
      Top             =   120
   End
   Begin VB.Shape shpSeek 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF80&
      FillColor       =   &H0000FFFF&
      Height          =   105
      Left            =   720
      Shape           =   3  'Circle
      Top             =   675
      Width           =   105
   End
   Begin VB.Line greenline 
      BorderColor     =   &H000000FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   5
      Index           =   1
      X1              =   720
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   2175
      Left            =   720
      Top             =   960
      Width           =   6255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF00FF&
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Height          =   2175
      Left            =   720
      Top             =   960
      Width           =   6255
   End
   Begin VB.Image imgOn 
      Height          =   285
      Left            =   7080
      Picture         =   "frmMain.frx":67EB
      ToolTipText     =   "Click to turn background music off"
      Top             =   4440
      Width           =   435
   End
   Begin VB.Image imgOff 
      Height          =   285
      Left            =   7080
      Picture         =   "frmMain.frx":6BF3
      ToolTipText     =   "Click to turn background music on"
      Top             =   4440
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label txtScroll 
      BackStyle       =   0  'Transparent
      Caption         =   "vintips"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1335
      Left            =   720
      TabIndex        =   17
      Top             =   3840
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   720
      Top             =   3240
      Width           =   6255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   5
      Index           =   0
      X1              =   720
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VIN UTILITY KIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   720
      TabIndex        =   16
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   7
      Index           =   0
      X1              =   720
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   2
      Height          =   495
      Left            =   720
      Top             =   3240
      Width           =   6255
   End
   Begin VB.Menu appl 
      Caption         =   "&Application"
      Begin VB.Menu mnuAppl 
         Caption         =   "VIN Web Compiler"
         Index           =   0
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "ALL2HTML Converter"
         Index           =   1
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "VIN Fone Directory"
         Index           =   2
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "IndraDhanush"
         Index           =   3
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "Remind Me Later"
         Index           =   4
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "Have A Break"
         Index           =   5
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "VIN File Split and Merge"
         Index           =   6
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "VIN Timer"
         Index           =   7
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAppl 
         Caption         =   "VIN Convert Centre"
         Index           =   8
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDhun 
         Caption         =   "Dhun"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnucal 
         Caption         =   "Calvin"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu games 
      Caption         =   "&Games"
      Begin VB.Menu game 
         Caption         =   "Vincanoid"
         Index           =   1
      End
      Begin VB.Menu game 
         Caption         =   "AIR FORCE"
         Index           =   2
      End
      Begin VB.Menu game 
         Caption         =   "VINTRAP"
         Index           =   3
      End
   End
   Begin VB.Menu music 
      Caption         =   "&BGMusic"
      Begin VB.Menu sur 
         Caption         =   "Just Beat It"
         Index           =   1
      End
      Begin VB.Menu sur 
         Caption         =   "Hawaii"
         Index           =   2
      End
      Begin VB.Menu sur 
         Caption         =   "Macarena"
         Index           =   3
      End
      Begin VB.Menu sur 
         Caption         =   "Partillie"
         Index           =   4
      End
      Begin VB.Menu sur 
         Caption         =   "Fall"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu sur 
         Caption         =   "Blinded"
         Index           =   6
      End
      Begin VB.Menu sur 
         Caption         =   "Bitedust"
         Index           =   7
      End
      Begin VB.Menu user 
         Caption         =   "User Defined"
      End
      Begin VB.Menu rt 
         Caption         =   "-"
      End
      Begin VB.Menu mute 
         Caption         =   "Mute "
      End
      Begin VB.Menu am 
         Caption         =   "About "
      End
   End
   Begin VB.Menu abou 
      Caption         =   "&Help"
      Begin VB.Menu shoot 
         Caption         =   "TroubleShooter"
         Shortcut        =   ^T
      End
      Begin VB.Menu me 
         Caption         =   "About Me"
         Shortcut        =   ^M
      End
      Begin VB.Menu about 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu tf 
         Caption         =   "-"
      End
      Begin VB.Menu soon 
         Caption         =   "Coming Soon....."
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rotate As Integer 'for animation
Dim multiple As Single 'for animation
Dim madhya As Integer
Dim scroll As String
Dim toleft As Boolean  'it is false when tip come from right
'but become true when going toleft after rest

'for tips

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "data\tipofmin.vin"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long

'for playbacking song
'variables are delared in module



Private Sub about_Click()
cmdHelp_Click (1)
End Sub




Private Sub am_Click()
MsgBox "All the background music included in this software" & Chr(13) & _
" are Private / Public Property of their composer and companies " & Chr(13) & _
"I am not using them for any commercial purpose " & Chr(13) & _
"Your system must have MIDI-Sequencer to PlayBack the BGM"
End Sub





Private Sub cmdAppl_Click(Index As Integer)
Dim temp As Long
On Error GoTo exeerror
 If Index = 0 Then
  temp = Shell(App.Path & "\webcompiler.exe", vbNormalFocus)
 ElseIf Index = 1 Then
  temp = Shell(App.Path & "\all2html.exe", vbNormalFocus)
 ElseIf Index = 2 Then
  temp = Shell(App.Path & "\fonedirectory.exe", vbNormalFocus)
 ElseIf Index = 3 Then
  temp = Shell(App.Path & "\indradhanush.exe", vbNormalFocus)
 ElseIf Index = 4 Then
  fillloadrem  'make loadrem.vin full so it can envoke the remindme later
  temp = Shell(App.Path & "\vinreminder.exe", vbNormalFocus)
 ElseIf Index = 5 Then
  temp = Shell(App.Path & "\break.exe", vbNormalFocus)
 ElseIf Index = 6 Then
  temp = Shell(App.Path & "\vinsplit.exe", vbNormalFocus)
 ElseIf Index = 7 Then
  temp = Shell(App.Path & "\vintimer.exe", vbNormalFocus)
 ElseIf Index = 8 Then
  temp = Shell(App.Path & "\convert.exe", vbNormalFocus)
 ElseIf Index = 9 Then
 
 End If
  
  cmdAppl(Index).Enabled = False
Exit Sub
exeerror:
 MsgBox "Application file  is not found in its " _
  & "Default directory  "
End Sub

Private Sub cmdAppl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAppl(Index).BackColor = &HFF80FF
End Sub



Private Sub cmdHelp_Click(Index As Integer)
Dim temp As Long
On Error GoTo exeerror
If Index = 0 Then
temp = Shell(App.Path & "\credit.exe", vbNormalFocus)
ElseIf Index = 1 Then
temp = Shell(App.Path & "\about.exe", vbNormalFocus)
ElseIf Index = 2 Then
 'MsgBox "This will start the the troubleshooter which will restore" & Chr(13) _
 '& "The VIN UTILITY KIT to its ideal condition"
 shoot_Click
End If
Exit Sub
exeerror:
MsgBox "Unable to locate specified exe"
End Sub

Private Sub Command1_Click()
'cmd = "seek vin to " & shpSeek.Top
' errorCode = mciSendString(cmd, returnStr, 255, 0)
' errorCode = mciSendString("play vin", returnStr, 255, 0)
End Sub

Private Sub Command10_Click()
Dim temp As Long
On Error GoTo exeerror
   temp = Shell(App.Path & "\vintrap.exe", vbNormalFocus)
   Exit Sub
exeerror:
MsgBox "VINTRAP is a simple but chellanging game created in vb" & Chr(13) & _
 "It has 3 stages .You have to defeat your opponent computer or human player By trapping the ball" & Chr(13) & _
 "Its size is more than 1MB so it is not included in this kit " & Chr(13) & _
 "But it will be added in futures versions " & Chr(13) & _
 "However you can get it saperatly and placed it in applications default folder."

End Sub


Private Sub Command14_Click()
 
MsgBox "Click on the shortcut of dhun.exe created on desktop or from program " & Chr(13) & _
"start menu ->Vin Utility Kit->Vinsoft->dhun.lnk" & Chr(13) & _
" It is dos based program (created on c++) so not envoked from here"
End Sub




Private Sub Command8_Click()
MsgBox "VINCANOID is based on the historical game arcanoid" & Chr(13) & _
 "It has more than 30 stages with 9 background songs  " & Chr(13) & _
 "Its size is more than 1MB so it is not included in this kit " & Chr(13) & _
 "But it will be added in futures versions or you can get it stand alone. "
End Sub

Private Sub Command9_Click()
MsgBox "AIRFORCE :- Your mission is to destroy the American fighter jets and choppers" & Chr(13) & _
 "It has 2 stages with 5 background songs " & Chr(13) & _
 "Its size is more than 1MB so it is not included in this kit " & Chr(13) & _
 "But it will be added in futures versions or you can get it stand alone."

End Sub








Private Sub Form_Click()
Dim i As Integer  'when any utility start then its button was disabled now enabling again
For i = 0 To 8
cmdAppl(i).Enabled = True
Next
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
'If TypeOf Source Is PictureBox And X < 6120 And X > 690 Then
'shpSeek.Left = X
'End If
End Sub

Private Sub Form_Load()
'init globals
frmmain.Show
madhya = 1
 scroll = "   VIN UTILITY KIT (Beta Version) :- By Vinod Kotiya (Application - Launcher) ****"
 toleft = False
 rotate = 2
 multiple = 3
 'tips
 ' Seed Rnd
    Randomize
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        MsgBox " Tip of the minute file " & TIP_FILE & " was not found? "
        txtScroll.Caption = " Tip of the minute file " & TIP_FILE & " was not found? Please run troubleshooter"
    End If


 
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu music
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 8
cmdAppl(i).BackColor = &HFFFF00
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Trim(mute.Caption) = "Mute" Then    'closesong only when song
closesong                 'is playing
End If
Dim temp As Long
On Error GoTo exeerror
temp = Shell(App.Path & "\about.exe", vbNormalFocus)
Exit Sub
exeerror:
 MsgBox "Application 'about.EXE' is not found in its " _
  & "Default directory about.exe "
  
End Sub

Private Sub game_Click(Index As Integer)
If Index = 1 Then
 Command8_Click
ElseIf Index = 2 Then
  Command9_Click
ElseIf Index = 3 Then
 Command10_Click
End If
End Sub

Private Sub imgOff_Click()
imgOff.Visible = False
imgOn.Visible = True
mute_Click

End Sub

Private Sub imgOn_Click()
imgOff.Visible = True
imgOn.Visible = False
mute_Click
End Sub






Private Sub me_Click()
 cmdHelp_Click (0)
End Sub

Private Sub mnuAppl_Click(Index As Integer)
  cmdAppl_Click (Index)
End Sub

Private Sub mnuDhun_Click()
 Command14_Click
End Sub

Private Sub mute_Click()

'mute.Checked = Not mute.Checked
If Trim(mute.Caption) = "Mute" Then
 closesong
 mute.Caption = "Play"
 imgOff.Visible = True
Else
 playsong
 mute.Caption = "Mute"
 imgOn.Visible = True
End If
End Sub


Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
'Timer1.Interval = 0
'If TypeOf Source Is PictureBox And X < 6120 And X > 690 Then
'shpSeek.Left = 690 + X
' cmd = "seek vin to " & Str(Round(shpSeek.Left / (6120 / songlength)))
' errorCode = mciSendString(cmd, returnStr, 255, 0)
' errorCode = mciSendString("play vin", returnStr, 255, 0)
' MsgBox "vo" & cmd
'End If
    
'Timer1.Interval = 200
End Sub


Private Sub shoot_Click()
Load frmShoot
frmShoot.Visible = True
End Sub

Private Sub shpSeek_Click()
Timer1.Interval = 0
End Sub

Private Sub shpSeek_DragDrop(Source As Control, X As Single, Y As Single)
'If Source = Line1 Or Source = Line2 Then
'    shpSeek.Left = X
'End If
End Sub

Private Sub shpSeek_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'If Source <> Line1 Then
'    shpSeek.Top = 90
'End If
End Sub





Private Sub soon_Click()
MsgBox "Coming Soon ..............." & Chr(13) & _
"*******************" & Chr(13) & _
"VIN File Lock v1.0" & Chr(13) & _
 "VIN Icon Maker v1.0" & Chr(13) & _
 "VIN Thumbnail Creator v1.0" & Chr(13) & _
 "VIN Sound Editor v1.0" & Chr(13) & _
 "VIN JavaScript/VBScript Compiler" & Chr(13) & _
 "ROAD FIGHTER"
End Sub

Private Sub sur_Click(Index As Integer)
sur(Index).Checked = Not sur(Index).Checked
Dim i As Integer
For i = 1 To 7
  If i <> Index Then sur(i).Checked = False
 Next
If sur(Index).Checked = True Then
      If Index = 1 Then
      songfilename = App.Path & "\data\beatit.mid"
      ElseIf Index = 2 Then
      songfilename = App.Path & "\data\hawaii.mid"
     ElseIf Index = 3 Then
      songfilename = App.Path & "\data\macarena.mid"
      ElseIf Index = 4 Then
      songfilename = App.Path & "\data\partille.mid"
      ElseIf Index = 5 Then
      songfilename = App.Path & "\data\fall.mid"
      ElseIf Index = 6 Then
      songfilename = App.Path & "\data\blinded.mid"
      ElseIf Index = 7 Then
      songfilename = App.Path & "\data\bitedust.mid"
      End If
      
   txtScroll.Caption = "Initializing " & sur(Index).Caption & " ! Please wait ........"
 'playsong
End If
If Trim(mute.Caption) = "Play" Then
 imgOff_Click    'will envoke mute_click and play song
Else
playsong
End If
End Sub


Private Sub Timer1_Timer()

On Error GoTo vinerror
'scrolling caption
frmmain.Caption = Mid$(scroll, madhya, Len(scroll) - madhya)
 'temp = Mid$(scroll, 1, madhya)
 frmmain.Caption = frmmain.Caption & Mid$(scroll, 1, madhya) 'temp
 madhya = madhya + 1
 If madhya > Len(scroll) Then
  madhya = 1
 End If
 DoEvents
   
 'DANCING ANIMATION
 If rotate > 2 Then
  rotate = 0
  shpSeek.BackColor = (650000 * Rnd(45655657))
 End If
 cmdHelp(rotate).BackColor = (650000 * Rnd(440065657))   '&HFF&
 
 If rotate = 0 Then
  cmdHelp(1).BackColor = &H80FF80
  cmdHelp(2).BackColor = &H80FF80
 ElseIf rotate = 1 Then
  cmdHelp(0).BackColor = &H80FF80
  cmdHelp(2).BackColor = &H80FF80
 Else
  cmdHelp(0).BackColor = &H80FF80
  cmdHelp(1).BackColor = &H80FF80
 End If
  rotate = rotate + 1
   cmdHelp(0).Top = cmdHelp(0).Top + (rotate * multiple)
   cmdHelp(1).Top = cmdHelp(1).Top - (rotate * multiple)
  If cmdHelp(0).Top > cmdHelp(2).Top + cmdHelp(2).Height Then
  multiple = -3
  ElseIf cmdHelp(0).Top < cmdHelp(2).Top - cmdHelp(2).Height Then
  multiple = 3

  End If
  
  'if song close than repeat it
 'get the position of song
    cmd = "status vin position"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    shpSeek.Left = 690 + Round(Val(returnStr) * (6120 / songlength)) '6120 is lines length
    greenline(1).X2 = shpSeek.Left '- greenline(1).X1
    If songlength = Val(returnStr) Then
      If sur(1).Checked = True Then
       sur_Click (2)
      ElseIf sur(2).Checked = True Then
       sur_Click (3)
      ElseIf sur(3).Checked = True Then
       sur_Click (4)
       ElseIf sur(4).Checked = True Then
       sur_Click (5)
       ElseIf sur(5).Checked = True Then
       sur_Click (6)
       ElseIf sur(6).Checked = True Then
       sur_Click (7)
       ElseIf sur(7).Checked = True Then
       sur_Click (1)
      End If
     'playsong 'errorCode = mciSendString("play vin from 2", returnStr, 255, 0)
    End If
 
 Exit Sub
vinerror:
   'if any mci error quit chuchap
End Sub

Private Sub fillloadrem()
Dim fnum As Integer

On Error GoTo FileError
  fnum = FreeFile
  Open App.Path & "\data\loadrem.vin" For Output As #1
   Print #fnum, "loadremember Vinod Kotiya " _
               & " is calling you from a program "
 Close #fnum
 Exit Sub

FileError:
    
    MsgBox "Unable to write . Unkown error while filling file " & "data\loadrem.vin"

End Sub

Private Sub DoNextTip()
    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd(43433)) + 1)
    frmmain.DisplayCurrentTip
 End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        txtScroll.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub Timer2_Timer()
If txtScroll.Left > cmdAppl(0).Left And toleft = False Then
 txtScroll.Left = txtScroll.Left - 50
ElseIf toleft = False Then
 Timer2.Interval = 0
 Timer3.Interval = 9000
 toleft = True
End If
If txtScroll.Left + txtScroll.Width > cmdAppl(0).Left And toleft = True Then
 txtScroll.Left = txtScroll.Left - 50
ElseIf toleft = True Then
 txtScroll.Left = cmdAppl(0).Left + txtScroll.Width 'Timer2.Interval = 0
 toleft = False
 DoNextTip
End If
DoEvents
End Sub

Private Sub Timer3_Timer()
Timer2.Interval = 20
Timer3.Interval = 0
End Sub



Private Sub txtScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu music
End Sub

Private Sub user_Click()
closesong
'fall.Checked = False
'maca.Checked = False
'partr.Checked = False
songfilename = InputBox("Please type the location of any midi file below for playback" & _
   "Use full path with extension ", "File Name Confirmation", "d:\music\kashmir.mid")
 txtScroll.Caption = "Initializing " & songfilename & " ! Please wait ........"
'  playsong
  

If Trim(mute.Caption) = "Play" Then
 imgOff_Click    'will envoke mute_click and play song
Else
playsong
End If

End Sub




