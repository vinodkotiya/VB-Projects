VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReminder 
   BackColor       =   &H00FF8080&
   Caption         =   "Remind Me Later"
   ClientHeight    =   5145
   ClientLeft      =   1530
   ClientTop       =   1380
   ClientWidth     =   5865
   Icon            =   "frmReminder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmReminder.frx":0ECA
   ScaleHeight     =   5145
   ScaleWidth      =   5865
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   120
   End
   Begin VB.Frame frmRecord 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Record/Play the Message"
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
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   5655
      Begin VB.CommandButton cmdPR 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pause"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStoprec 
         BackColor       =   &H000000FF&
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdPlay 
         BackColor       =   &H00FFFF80&
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Record"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 Sec"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Message Type"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   5655
      Begin VB.OptionButton optVoice 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Voice Message"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   210
         Width           =   2055
      End
      Begin VB.OptionButton optText 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Text Message"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFF00&
      Caption         =   "<<--Delete This Message "
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "See Message For-->>"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   16777152
      ScrollBars      =   2
      TextRTF         =   $"frmReminder.frx":4967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbMessage 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFF00&
      Caption         =   "OK Remind Me Later by above TEXT Message"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Remind Me"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optBefore 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Before"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optOn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "On"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optAfter 
         BackColor       =   &H00FFC0C0&
         Caption         =   "After"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   210
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   5760
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   5760
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Menu about 
      Caption         =   "&Help"
      Begin VB.Menu help 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu smnabout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
      Begin VB.Menu credit 
         Caption         =   "Credit"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu vin 
      Caption         =   "&VIN UTILITY KIT"
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim delinstruction As Boolean  ' delete the instructions when mouse click
Dim errorCode As Integer
Dim returnStr As String * 255
Dim cmd As String * 255
Private Declare Function mciSendString Lib "winmm.dll" _
    Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long
Dim isWavClose As Boolean    'true when wave closed
Dim wavlength As Long  'store lenght of wave file
    


Private Sub cmbDate_GotFocus()
filldate
End Sub

Private Sub cmbMessage_Change()
'RichTextBox1.LoadFile "c:\22mar2003", rtfText 'App.Path & "\RTFText.rtf", rtfRTF
'Dim datefile As String
'Dim pos As Integer
'datefile = "messages\" & cmbMessage.Text & ".vin"
'pos = InStr(datefile, "v")
'RichTextBox1.LoadFile datefile, rtfText
End Sub



Private Sub cmbMonth_Change()
filldate
End Sub


Private Sub cmbMonth_GotFocus()
filldate
End Sub



Private Sub cmdDelete_Click()
Dim Fsys As New Scripting.FileSystemObject
Dim datefile As String
On Error GoTo vinerror
If InStr(6, cmbMessage.Text, "V0ICE", vbBinaryCompare) > 0 Then
  closeWav    'if open then not delete
  datefile = App.Path & "\messages\" & cmbMessage.Text & ".wav"
Else
 datefile = App.Path & "\messages\" & cmbMessage.Text & ".vin"
End If
 cmbMessage.RemoveItem cmbMessage.ListIndex
 Fsys.DeleteFile datefile
cmbMessage.ListIndex = 0
Exit Sub

vinerror:
 MsgBox "Unknown error when deleting the " & datefile
End Sub

Private Sub cmdOk_Click()
Dim reply As String



'Dim txt As String

Dim j As Integer
Dim newitem As String           'store date
'Dim reply As String
newitem = cmbDate.Text & "-" & cmbMonth.Text & "-" & cmbYear.Text 'Year(Date)
'newMssgFile = newitem
reply = MsgBox("This will save the message as " & Chr(34) & newitem & Chr(34) & " displayed on text box/recorded." & Chr(13) _
 & "Are you sure that you have write/record the message and want to save it?", vbYesNo, "Message Confirmation") '& Chr(13) & Chr(13) & "You can give any other Name for saving ", "Message Confirmation", newitem)
If reply = vbNo Then
  MsgBox "NO"
  Exit Sub            'if canceled then exit
End If
'add acording optBefore.value & optOn.value & optAfter.value
' when to envoke the message like 24may2003B
If optBefore.Value = True Then
 newitem = newitem & "B"
ElseIf optOn.Value = True Then
 newitem = newitem & "O"
Else
 newitem = newitem & "A"
End If
'also add mssg is text or voice
If optText.Value = True Then
 newitem = newitem & "(TEXT)"
ElseIf optVoice.Value = True Then
 newitem = newitem & "(V0ICE)"
End If
'scan whole txtsearchlist to prevent duplicasy of new entery
'add the item only when some messages are given
If Trim(RichTextBox1.Text) <> "" Then
  For j = 0 To cmbMessage.ListCount
   If cmbMessage.List(j) = newitem Then
        j = Round(Rnd(100) * Second(Time))  'used temporarily
        newitem = newitem & j 'item already exist so give another name with random no like 24may2003B34
        j = cmbMessage.ListCount
   End If
  Next
 cmbMessage.AddItem newitem
Else
 MsgBox "Please Enter some message for remind you later "
 Exit Sub
End If
'save dates list in file dates.vin
savedates
 ''now save the message
 If optText.Value = True Then
 RichTextBox1.Text = "Message " & newitem & " :" & Chr(13) & RichTextBox1.Text & Chr(13) & " Message given on " & Now
 newitem = App.Path & "\messages\" & newitem & ".vin"
 'MsgBox newitem
 RichTextBox1.SaveFile (newitem), rtfText
 ElseIf optVoice.Value = True Then
   newitem = App.Path & "\messages\" & newitem & ".wav"
   cmd = "save recwave " & Chr(34) & newitem & Chr(34)
    errorCode = mciSendString(cmd, returnStr, 255, 0)
   If errorCode <> 0 Then MsgBox "not saved use troubleshooter" & cmd
  closeWav
   isWavClose = True
   optVoice_Click    'reset variables
End If
 
 Exit Sub
FileError:
    
    MsgBox "Unkown error while saving file " & "dates.vin" _
    & Chr(13) & "Please click on Trobleshooting in the utility kit "
End Sub
Private Sub closeWav()
 cmd = "close recwave"
     errorCode = mciSendString(cmd, returnStr, 255, 0)
'   If errorCode <> 0 Then MsgBox "Error when closing device " & returnStr
End Sub
Private Sub savedates()
Dim FNum As Integer
Dim i As Integer
Dim txt As String
On Error GoTo FileError
    FNum = FreeFile
          'txt = "c:\Dates.vin"
    Open App.Path & "\data\dates.vin" For Output As #1    'create a file similar to envoked date
                        'erasing filename
      If cmbMessage.ListCount > 0 Then
        For i = cmbMessage.ListCount - 1 To 0 Step -1
         txt = txt & cmbMessage.List(i) & Chr(13)
        Next
       Print #FNum, txt
       End If
   Close #FNum
    'OpenFile = "c:\vin.vin" 'CommonDialog1.FileName
 
 Exit Sub
FileError:
    
    MsgBox "Unkown error while saving file " & "dates.vin" _
    & "Please click on Trobleshooter in the utility kit "
End Sub
Private Sub Command1_Click()
'Dim tempa As Long
'tempa = Shell(App.Path & "\days.vbs")


End Sub

Private Sub cmdPlay_Click()
 cmd = "status recwave length"
 errorCode = mciSendString(cmd, returnStr, 255, 0)
 If errorCode <> 0 Then MsgBox "Device failed to find length" & returnStr
 wavlength = Val(returnStr)
 cmd = "play recwave from 0"
 errorCode = mciSendString(cmd, returnStr, 255, 0)
 If errorCode <> 0 Then MsgBox "Device failed to playback the file" & returnStr
 cmdStoprec.Enabled = True
 cmdPlay.Enabled = False
 cmdPR.Enabled = True
 Timer1.Interval = 500
End Sub

Private Sub cmdPR_Click()
If cmdPR.Caption = "Pause" Then
 cmdPR.Caption = "Resume"
 cmd = "pause recwave "
 errorCode = mciSendString(cmd, returnStr, 255, 0)
 If errorCode <> 0 Then MsgBox "Device failed to pause" & returnStr
  Timer1.Interval = 0
ElseIf cmdPR.Caption = "Resume" Then
 cmdPR.Caption = "Pause"
  cmd = "resume recwave "
 errorCode = mciSendString(cmd, returnStr, 255, 0)
 If errorCode <> 0 Then MsgBox "Device failed to resume" & returnStr
  Timer1.Interval = 500
End If
End Sub

Private Sub cmdRecord_Click()
'cue waveaudio input
Dim tmpWave As String
tmpWave = App.Path & "\data\vin.wav"
closeWav    'if open then not record
cmd = "open " & Chr(34) & tmpWave & Chr(34) & " type waveaudio alias recwave" ' & Chr(34) & tempfile & Chr(34) & " type waveaudio alias vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    'If errorCode <> 0 Then MsgBox "Device failed "
cmd = "record recwave from 0"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
If errorCode <> 0 Then MsgBox "Device not supporting recording " & returnStr
cmdStoprec.Enabled = True
cmdPR.Enabled = True
cmdRecord.Enabled = False
Timer1.Interval = 500
End Sub

Private Sub cmdStoprec_Click()
cmd = "stop recwave "
    errorCode = mciSendString(cmd, returnStr, 255, 0)
 If errorCode <> 0 Then MsgBox "Device failed to pause" & returnStr
   ' cmdRecord.Enabled = True
    cmdStoprec.Enabled = False
    cmdPR.Enabled = False
    cmdPR.Caption = "Pause"
    cmdPlay.Enabled = True
    Timer1.Interval = 0
End Sub

Private Sub Command2_Click()
Dim datefile As String
Dim pos As Integer
On Error GoTo vinerror
If InStr(6, cmbMessage.Text, "V0ICE", vbBinaryCompare) > 0 Then
   MsgBox "This is VOICE Message.Press Play to listen the message."
   Dim tmpWave As String
   closeWav    'if open then
   tmpWave = App.Path & "\messages\" & cmbMessage.Text & ".wav"
    cmd = "open " & Chr(34) & tmpWave & Chr(34) & " type waveaudio alias recwave" ' & Chr(34) & tempfile & Chr(34) & " type waveaudio alias vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    If errorCode <> 0 Then MsgBox "Device failed "
    frmRecord.Enabled = True
    cmdRecord.Enabled = False
    cmdPlay.Enabled = True
    wavlength = 167665565
    isWavClose = False
    lblTime.Caption = "0 Sec"
    RichTextBox1.Enabled = False
Else
 datefile = App.Path & "\messages\" & cmbMessage.Text & ".vin"
'pos = InStr(datefile, "v")
 RichTextBox1.Enabled = True
 RichTextBox1.LoadFile datefile, rtfText
 frmRecord.Enabled = False
 cmdRecord.Enabled = False
 cmdPR.Enabled = False
 cmdStoprec.Enabled = False
 cmdPlay.Enabled = False
End If
Exit Sub
vinerror:
MsgBox "The messages for the date " & datefile & " may be deleted or not exist " & Chr(13) & _
"Or you have not selected any message to display/Play. Please Select first."
End Sub

Private Sub credit_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\credit.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'CREDIT.EXE' is not found in its " _
  & "Default directory CREDIT.exe "
End Sub

Private Sub Form_Load()
'OLE1.SourceDoc = App.Path & "\lnkRem.vbs"

emptyloadrem          'makeempty the loadrem.vin

delinstruction = True   ' delete the instructions when mouse click on rtf
RichTextBox1.Text = RichTextBox1.Text & Chr(13) _
     & "To See any previous message select" & Chr(13) & " the message name in form of date (eg. 23-Nov-2003B )  " _
      & " from dropdown and " & Chr(13) & " press the button " & Chr(34) & " See Message For " & Chr(34) & Chr(13) _
      & " Similarly delete any message by pressing " & Chr(34) & " Delete This Message " & Chr(34)

Dim i As Integer
Dim InFile As Integer   ' Descriptor for file.
Dim nextdate As String
    ' Obtain the next free file descriptor.
'  On Error GoTo FileError
InFile = FreeFile
Open App.Path & "\data\dates.vin" For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, nextdate
        If Trim(nextdate) <> "" Then
        cmbMessage.AddItem nextdate
        End If
    Wend
    'nextdate = Input(LOF(InFile), #InFile)
 Close InFile
 
 'Dim parts() As String
  'parts = Split(nextdate, "^")
 'split txt and save it to arry then add to messagebox
  'For i = 1 To UBound(parts)
   '   cmbMessage.AddItem parts(i - 1)
  'Next
'cmbMessage.AddItem i
'Next
'cmbMessage.ListIndex = 0
''fill month box
cmbMonth.AddItem "Jan"
cmbMonth.AddItem "Feb"
cmbMonth.AddItem "Mar"
cmbMonth.AddItem "Apr"
cmbMonth.AddItem "May"
cmbMonth.AddItem "Jun"
cmbMonth.AddItem "Jul"
cmbMonth.AddItem "Aug"
cmbMonth.AddItem "Sep"
cmbMonth.AddItem "Oct"
cmbMonth.AddItem "Nov"
cmbMonth.AddItem "Dec"
i = Month(Date)
cmbMonth.ListIndex = i - 1

''fill year box
For i = 0 To 5
cmbYear.AddItem Year(Date) + i
Next
cmbYear.ListIndex = 0
filldate
Exit Sub
FileError:
    MsgBox "Unkown error while opening file " & "date.vin" _
    & "Please open notepad don't write anything and save it as " _
    & " ''data\date.vin'' in the application's default folder" _
    & "or click on Trobleshooting in the utility kit "
End Sub

Private Sub filldate()
''fill the date box
Dim i As Integer

Dim dateforday As String
Dim prevday As Integer        'store preveos day
Dim j As Integer
For j = cmbDate.ListCount To 1 Step -1
  cmbDate.RemoveItem (j - 1)         'make empty
 ' MsgBox j
Next
cmbDate.AddItem "1"
dateforday = "1" & "-" & cmbMonth.Text & "-" & Mid(cmbYear.Text, 1, 2)
For j = 31 To 0 Step -1
 i = Day(DateAdd("d", 1, dateforday))
 If i < prevday Then
    Exit For
 End If
 
 dateforday = DateAdd("d", 1, dateforday)
 prevday = i
cmbDate.AddItem i
Next
i = Day(Date)
cmbDate.ListIndex = i - 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
savedates
If isWavClose = False Then
 cmd = "close recwave"
     errorCode = mciSendString(cmd, returnStr, 255, 0)
   'If errorCode <> 0 Then MsgBox "error closing "
End If
End Sub

Private Sub help_Click()
MsgBox " HELP " & Chr(13) & "**************" & Chr(13) & _
"FEATURES :- If you want any message to be displayed/played on/after/before any date" & Chr(13) & _
"            When you start your computer." & Chr(13) & _
"How to Give a Message" & Chr(13) & "*********************" & Chr(13) & _
" 1. Select the Date with On/Before/After option. It specify that when to recall the message." & Chr(13) & _
" 2. Now choose the Message type TEXT or VOICE" & Chr(13) & _
" 3. If your message type is TEXT then write it in the textbox " & Chr(13) & _
"     And Press " & Chr(34) & "OK Remind Me the Above Text Message " & Chr(34) & Chr(13) & _
" 4. If your message type is VOICE then press " & Chr(34) & "Record" & Chr(34) & "To start recording " & Chr(13) & _
"     Message. Press " & Chr(34) & "Stop" & Chr(34) & " to finish recording" & Chr(13) & _
"How to See/Play a Previous Message" & Chr(13) & "************************" & Chr(13) & _
" At bottom of application First select your message from dropDown" & Chr(13) & _
"  and then Press " & Chr(34) & "See Message For-->>" & Chr(34) & Chr(13) & _
"  If message is text it will displayed on textbox.If message is VOICE then press PLAY " & Chr(13) & _
" Similarly you can delete any message by Selecting it and Pressing" & Chr(34) & "<<--Delete This Message" & Chr(34) & Chr(13)
End Sub

Private Sub optText_Click()
If optText.Value = True Then
 frmRecord.Enabled = False
  cmdRecord.Enabled = False
 cmdPR.Enabled = False
 cmdStoprec.Enabled = False
 cmdPlay.Enabled = False
 RichTextBox1.Enabled = True
End If
End Sub

Private Sub optVoice_Click()
If optVoice.Value = True Then
 wavlength = 167665565
 isWavClose = False
 cmdOk.Caption = "OK Remind Me Later by above VOICE Message"
 RichTextBox1.Enabled = False
 frmRecord.Enabled = True
 cmdRecord.Enabled = True
 cmdPR.Enabled = False
 cmdStoprec.Enabled = False
 cmdPlay.Enabled = False
End If
End Sub

Private Sub RichTextBox1_Click()
If delinstruction = True Then  ' delete the instructions when mouse click
 RichTextBox1.Text = " "
 cmbMessage.Text = " "
 delinstruction = False
End If
End Sub
Private Sub emptyloadrem()
'make empty the loadrem.vin whenever call
'because this form will display directly when loadrem.vin
'is full which is made full by calling program
'at calling from startup loadrem.vin should be empty
Dim FNum As Integer

On Error GoTo FileError
  FNum = FreeFile
  Open App.Path & "\data\loadrem.vin" For Output As #1
   '''''''''''Print #FNum, "" making file empty
 Close #FNum
 Exit Sub

FileError:
    
    MsgBox "Unkown error while flushing file " & "data\loadrem.vin"
    

End Sub

Private Sub smnabout_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\about.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'about.EXE' is not found in its " _
  & "Default directory about.exe "
End Sub

Private Sub Timer1_Timer()
 cmd = "status recwave position"
 errorCode = mciSendString(cmd, returnStr, 255, 0)
 If errorCode <> 0 Then MsgBox "Device failed to pause" & returnStr
lblTime.Caption = Str(Val(returnStr) / 1000) & " Sec"
If wavlength < Val(returnStr) Or wavlength = Val(returnStr) Then
  cmdStoprec.Enabled = False
  cmdPlay.Enabled = True
  cmdPR.Enabled = False
  Timer1.Interval = 0
  'MsgBox "go" & Val(returnStr)
End If

End Sub

Private Sub vin_Click()
Dim temp As String
temp = MsgBox("Please check that any copy of VIN UTILITY KIT is running or not" & Chr(13) & _
"If it is already running then press NO and if it is not running press YES ." & Chr(13), vbYesNo)
If temp = vbYes Then
 Dim tempa As Long
On Error GoTo Exeerror
Unload Me
tempa = Shell(App.Path & "\vin_utility.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'vin_utility.EXE' is not found in its " _
  & "Default directory vin_utility.exe "
 
 
End If

End Sub
