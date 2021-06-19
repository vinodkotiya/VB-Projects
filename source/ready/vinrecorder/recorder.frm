VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmRecord 
   Caption         =   "Form1"
   ClientHeight    =   2520
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton delAft 
      Caption         =   "Delete After"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton delBef 
      Caption         =   "Delete Before"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton tothere 
      Caption         =   "^"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton fromhere 
      Caption         =   "v"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin ComctlLib.Slider hsbSeek 
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   327682
      Max             =   200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   120
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Sec"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblSec 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu saveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim seekpressed As Boolean 'true when mouse is draging seek
Dim seekichanged As Boolean  'true when seek changed by mouse
Dim YahanSe As Long
Dim VahanTak As Long


Private Sub Command1_Click()
'cmd = "record vin from 2000 to 8000 "
'errorCode = mciSendString(cmd, returnStr, 255, 0)
cmd = "save vin c:\temp.wav "
errorCode = mciSendString(cmd, returnStr, 255, 0)
If errorCode <> 0 Then
 MsgBox "error"
End If

End Sub

Private Sub cmdPlay_Click()

If cmdPlay.Caption = "Play" And Trim(songfilename) <> " " Then
 cmdPlay.Caption = "Stop"
 playsong
ElseIf cmdPlay.Caption = "Stop" Then
 cmdPlay.Caption = "Play"
 hsbSeek.Value = 0
 closesong
End If
End Sub

Private Sub Command2_Click()

MsgBox Str((songlength / hsbSeek.Max) * hsbSeek.Value)
End Sub

Private Sub delAft_Click()
cmd = "delete vin from " & Str(Int(songlength / hsbSeek.Max) * hsbSeek.Value)
errorCode = mciSendString(cmd, returnStr, 255, 0)
If errorCode <> 0 Then
 MsgBox "error to  jj"
End If

cmd = "save vin c:\temp.wav "
errorCode = mciSendString(cmd, returnStr, 255, 0)
closesong
tempfile = "c:\temp.wav"
playsong
If errorCode <> 0 Then
 MsgBox "error"
End If
End Sub

Private Sub delBef_Click()
cmd = "delete vin to " & Str(Int(songlength / hsbSeek.Max) * hsbSeek.Value)
errorCode = mciSendString(cmd, returnStr, 255, 0)
If errorCode <> 0 Then
 MsgBox "error to  jj"
End If

cmd = "save vin c:\temp.wav "
errorCode = mciSendString(cmd, returnStr, 255, 0)
closesong
tempfile = "c:\temp.wav"
playsong
If errorCode <> 0 Then
 MsgBox "error"
End If

End Sub

Private Sub Form_Load()
seekpressed = False
seekichanged = False
End Sub

Private Sub Form_Terminate()
'closesong
End Sub

Private Sub Form_Unload(Cancel As Integer)
closesong
End Sub

Private Sub hsbSeek_Change()

If seekichanged = True Then
Timer1.Interval = 0
'MsgBox hsbSeek.Value
cmd = "seek vin to " & Str(Round(hsbSeek.Value * (songlength / hsbSeek.Max)))
errorCode = mciSendString(cmd, returnStr, 255, 0)
errorCode = mciSendString("play vin", returnStr, 255, 0)


    
Timer1.Interval = 250
seekichanged = False
End If
End Sub

Private Sub hsbSeek_GotFocus()
'Timer1.Interval = 0
End Sub

Private Sub hsbSeek_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
seekpressed = True
End Sub

Private Sub hsbSeek_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
seekpressed = False
seekichanged = True
End Sub

Private Sub open_Click()
Dim dpos As Integer
Dim ext  As String
    CommonDialog1.Filter = "WAV|*.wav"
    CommonDialog1.ShowOpen
    If Trim(CommonDialog1.FileName) = "" Then Exit Sub
    dpos = InStr(CommonDialog1.FileName, ".")
    If dpos > 0 Then ext = Mid$(CommonDialog1.FileName, dpos + 1)
    If UCase$(ext) = "WAV" Then
      '  RichTextBox1.LoadFile CommonDialog1.FileName, 1
       ' WebBrowser1.Navigate CommonDialog1.FileName
        songfilename = CommonDialog1.FileName
        tempfile = songfilename
    End If
'playsong
End Sub

Private Sub save_Click()
    If songfilename <> "" Then
         cmd = "save vin " & songfilename
         errorCode = mciSendString(cmd, returnStr, 255, 0)
    Else
        FileSaveAs_Click
    End If
    
End Sub

Private Sub SaveAs_Click()

    CommonDialog1.DefaultExt = "wav"
    CommonDialog1.Filter = ".wav|*.wav|VIN Files|*.vin|All Files|*.*"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
  
     cmd = "save vin " & CommonDialog1.FileName
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    tempfile = CommonDialog1.FileName
    


End Sub

Private Sub Timer1_Timer()
 'if song close than repeat it
 'get the position of song
If cmdPlay.Caption = "Stop" And seekpressed = False Then
    cmd = "status vin position"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    'MsgBox (returnStr)
    If Val(returnStr) < songlength Then
     lblSec.Caption = Val(returnStr / 1000)
     hsbSeek.Value = Round(Val(returnStr) * (hsbSeek.Max / songlength))
    End If
    'If songlength = Val(returnStr) Then
    'playsong 'errorCode = mciSendString("play vin from 2", returnStr, 255, 0)
    'End If
End If
End Sub
