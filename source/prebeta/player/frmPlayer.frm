VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmPlayer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "VIN MEDIA PLAYER"
   ClientHeight    =   4680
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   0
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   4680
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2002
      Month           =   8
      Day             =   25
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calander"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OpenMedia"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FullScreen"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   3135
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   0
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu FileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ViewMenu 
      Caption         =   "View"
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Size As Integer


Private Sub Command1_Click()
MediaPlayer1.DisplaySize = mpFullScreen
Size = 1
End Sub

Private Sub Command2_Click()
Dim avi As String
Dim FNum As Integer

On Error GoTo FileError:
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.DefaultExt = "avi"
    CommonDialog1.Filter = "AVI Files|*.AVI|VIDEO Files|*.dat|MPEG|*.mpeg|MP3|*.mp3|All Files|*.*"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    MediaPlayer1.FileName = CommonDialog1.FileName
    Exit Sub
    
FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & CommonDialog1.FileName
    'OpenFile = ""

End Sub

Private Sub Command3_Click()
Calendar1.Visible = True
End Sub

Private Sub FileExit_Click()
Unload frmPlayer
End Sub

Private Sub FileOpen_Click()
Dim avi As String
Dim FNum As Integer

On Error GoTo FileError:
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.DefaultExt = "avi"
    CommonDialog1.Filter = "AVI Files|*.AVI|VIDEO Files|*.dat|MPEG|*.mpeg|MP3|*.mp3|All Files|*.*"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen

    'If UCase(Right(CommonDialog1.FileName, 3)) = "AVI" Then
       'tmode = avi
       
    'Else
     '  tmode = mp3
    'End If
    MediaPlayer1.FileName = CommonDialog1.FileName
    
    'MediaPlayer1.LoadFile CommonDialog1.FileName 'tmode
    'OpenFile = CommonDialog1.FileName
    Exit Sub
    
FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & CommonDialog1.FileName
    'OpenFile = ""

End Sub

Private Sub Form_Load()
Size = 0
End Sub

Private Sub MediaPlayer1_DblClick(Button As Integer, ShiftState As Integer, x As Single, y As Single)
Text1.Text = "Hi"
If Size = 0 Then
 MediaPlayer1.DisplaySize = mpFullScreen
 Size = 1
Else
 MediaPlayer1.DisplaySize = mpOneSixteenthScreen
 Size = 0
 End If

End Sub
