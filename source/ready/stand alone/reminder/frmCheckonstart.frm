VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCheckonstart 
   BackColor       =   &H00FF0000&
   Caption         =   "Remind Me Later"
   ClientHeight    =   1935
   ClientLeft      =   1320
   ClientTop       =   1080
   ClientWidth     =   6360
   ForeColor       =   &H00000000&
   Icon            =   "frmCheckonstart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6360
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   6375
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go to Next VOICE Message --->> "
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter New Messages"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1215
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2143
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmCheckonstart.frx":0ECA
   End
End
Attribute VB_Name = "frmCheckonstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txtloadrem As String
Dim voiceLength As Long
Dim isExit As Boolean 'true when X clicked
Dim messageUsed As Boolean 'true when any date of dates.vin is used as txt or voice display
Dim voiceFile As New Collection   'store  pure voice message name like 1-Apr-2003A(V0ICE)
' for mci command
Dim cmd As String * 255
Dim errorCode As Integer
Dim returnStr As String * 255
Private Declare Function mciSendString Lib "winmm.dll" _
    Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long


Private Sub cmdStop_Click()
   cmd = "seek recwave to " & voiceLength
   errorCode = mciSendString(cmd, returnStr, 255, 0)
   'If errorCode <> 0 Then MsgBox "Device failed to seek to end" & returnStr

End Sub


Private Sub Command2_Click()
On Error Resume Next
Load frmSplash
'frmSplash.Show
Unload Me
End Sub

Private Sub Form_Load()
'any of these two forms will display at a time
isExit = False
messageUsed = False
checkloadreminder     'will return txtloadrem
If Trim(txtloadrem) = "" Then    ''if loadreminder.VIN is null
frmCheckonstart.Hide      ''load reminder from startup and check the date
mayiremindyou
'MsgBox txtloadrem & "I am in startup"
Else                ''else IF REMINDER.VIN IS FULL THEN START from program
Command2_Click
'MsgBox "i am invoked from a program"
Unload Me
End If


End Sub
Private Sub checkloadreminder()
' check from where the reminder is called
' from kit or from startup
'STORE AS DATA\LOADREM.VIN
'while 'messages files should be stored in message folder
Dim FNum As Integer



On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\loadrem.vin" For Input As #1
    txtloadrem = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
  '  If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & "loadrem.vin" _
     & "To eliminate problem reinstall the software "
          
End Sub

Private Sub mayiremindyou()
' check from what message to be display or not today

On Error GoTo FileError
Dim txtdate As String

'if dates.vin file is empty means no messages then unload me from startup
'so load rtf box and check dates.vin is empty or not
RichTextBox1(0).LoadFile App.Path & "\data\dates.vin", rtfText
If Trim(RichTextBox1(0).Text) = "" Then
         Load frmReminder
         frmReminder.Show
         frmReminder.Visible = True
         Unload Me
         
          Exit Sub
End If


Dim InFile As Integer   ' Descriptor for file.
Dim nextdate As String       'in loop store nextdate of file dates.vin
Dim messagefile As String     'store message file name to be opened
Dim pos As Integer  'store position of B,O,A
Dim j As Integer
j = 1
Text1.Text = " Now Displaying Text Messages"
'OPEN THE FILE DATE.VIN AND STORE EACH LINE IN NEXTDATE TILL END
'AND CHECK THE MESSAGE IS ON,AFTER OR BEFORE TODAYS DATE

InFile = FreeFile
Open App.Path & "\data\dates.vin" For Input As InFile
    While Not EOF(InFile)
    
      Line Input #InFile, nextdate
       If InStr(6, nextdate, "V0ICE", vbBinaryCompare) > 0 Then       'seperate text and voice message
        voiceFile.Add nextdate       'store all voice message here then playback after
        'j = j + 1
       Else
        messagefile = App.Path & "\messages\" & nextdate & ".vin"   'messages file should be stored in message folder
         If Trim(nextdate) <> "" Then
             If InStr(6, nextdate, "B", vbBinaryCompare) > 0 Then  'true when find B MEANS BEFORE
              pos = InStr(6, nextdate, "B", vbBinaryCompare) ' GET POSN OF B
              nextdate = Mid(nextdate, 1, pos - 1)    'TRUNCATE B AND SUFFIX TO FILTER DATE
                If DateValue(nextdate) > DateValue(Date) Then  'IS THE DATE GREATER THEN TODAY
                 'MsgBox nextdate         'SO DISPLAY THE MESSAGE FOR BEFORE DATE
                 frmCheckonstart.Show
                 Load RichTextBox1(j)      'LOAD TEXT BOX WITH MESSAGE FILE  ON RUNTIME
                 RichTextBox1(j - 1).Visible = True
                 messageUsed = True
                 RichTextBox1(j).Top = RichTextBox1(j - 1).Top + RichTextBox1(0).Height
                 RichTextBox1(j - 1).LoadFile messagefile, rtfText
                 j = j + 1
                End If               'end of  >
             ElseIf InStr(6, nextdate, "A", vbBinaryCompare) > 0 Then  'true when find
              pos = InStr(6, nextdate, "A", vbBinaryCompare)  ' GET POSN OF A
              nextdate = Mid(nextdate, 1, pos - 1)          'TRUNCATE A AND SUFFIX TO FILTER DATE
               If DateValue(nextdate) < DateValue(Date) Then     'IS THE DATE LESS THEN TODAY
                 'MsgBox nextdate      'SO DISPLAY THE MESSAGE FOR AFTER DATE
                 frmCheckonstart.Show
                 Load RichTextBox1(j)
                 RichTextBox1(j - 1).Visible = True  'LOAD TEXT BOX WITH MESSAGE FILE  ON RUNTIME
                 messageUsed = True
                 RichTextBox1(j).Top = RichTextBox1(j - 1).Top + RichTextBox1(0).Height
                 RichTextBox1(j - 1).LoadFile messagefile, rtfText
                 j = j + 1
                End If               'end of  <
              ElseIf InStr(6, nextdate, "O", vbBinaryCompare) > 0 Then  'true when find O MEANS ON
              pos = InStr(6, nextdate, "O", vbBinaryCompare)    'GET POSITION OF O
              nextdate = Mid(nextdate, 1, pos - 1)     'TRUNCATE O AND SUFFIX TO FILTER DATE
                If DateValue(nextdate) = DateValue(Date) Then
                  'MsgBox nextdate
                 frmCheckonstart.Show
                 Load RichTextBox1(j)      'LOAD TEXT BOX WITH MESSAGE FILE  ON RUNTIME
                 RichTextBox1(j - 1).Visible = True
                 RichTextBox1(j).Top = RichTextBox1(j - 1).Top + RichTextBox1(0).Height
                 messageUsed = True
                 RichTextBox1(j - 1).LoadFile messagefile, rtfText
                 j = j + 1
                 
                End If               'end of  =
             End If     'end of B/A/O checkings
             'setheight of form
              frmCheckonstart.Height = RichTextBox1(j - 1).Top + RichTextBox1(j - 1).Height - 550
                   If j > 7 Then
                  MsgBox "Some other messages are waiting for you "
                  frmCheckonstart.Height = 2175     'set initial height
                     For j = 1 To 7
                     Unload RichTextBox1(j)
                     Next
                     j = 1      'now again display onprevious forms
                     
                   End If
                 
        'ElseIf Trim(nextdate) = " " Then
             
         '       frmCheckonstart.Visible = True
          '       MsgBox "HI"
        End If      'end of trim
       End If     'voice and text msg seperated
    Wend
    
    'nextdate = Input(LOF(InFile), #InFile)
 Close InFile

Text1.Text = "Now Playing Voice Messages"
EditVoiceMsg       'edit voice mssg if any then play
 If messageUsed = False Then End   'all dates in dates.vin are not usable
Command2.Enabled = True

cmdStop.Enabled = False
Text1.Text = "Press Enter New Messages"
cmdStop.Caption = "No VOICE Message Playing Currently."
'frmCheckonstart.BorderStyle = 4
'frmCheckonstart.MaxButton = False
Exit Sub
FileError:
  '  If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file data\dates.vin or message text file or fails to playback" & "" _
     & "The file may not exist on your harddrive."
      Close InFile
   ' OpenFile = ""
     
End Sub

Private Sub EditVoiceMsg()
'MsgBox voiceFile.Count
Dim j As Integer
If voiceFile.Count < 1 Then Exit Sub
For j = 1 To voiceFile.Count
   Dim pos As Integer
   Dim nextdate As String
       ' PlayVoiceMsg (voiceFile.Item(j))
  If InStr(6, voiceFile.Item(j), "B", vbBinaryCompare) > 0 Then   'true when find B MEANS BEFORE
       pos = InStr(6, voiceFile.Item(j), "B", vbBinaryCompare) ' GET POSN OF B
       nextdate = Mid(voiceFile.Item(j), 1, pos - 1)    'TRUNCATE B AND SUFFIX TO FILTER DATE
       If DateValue(nextdate) > DateValue(Date) Then  'IS THE DATE GREATER THEN TODAY
         'MsgBox nextdate         'SO DISPLAY THE MESSAGE FOR BEFORE DATE
          PlayVoiceMsg (voiceFile.Item(j))
       End If               'end of  >
  ElseIf InStr(6, voiceFile.Item(j), "A", vbBinaryCompare) > 0 Then  'true when find
       pos = InStr(6, voiceFile.Item(j), "A", vbBinaryCompare)  ' GET POSN OF A
       nextdate = Mid(voiceFile.Item(j), 1, pos - 1)          'TRUNCATE A AND SUFFIX TO FILTER DATE
       If DateValue(nextdate) < DateValue(Date) Then     'IS THE DATE LESS THEN TODAY
          'MsgBox nextdate      'SO DISPLAY THE MESSAGE FOR AFTER DATE
          PlayVoiceMsg (voiceFile.Item(j))
       End If               'end of  <
  ElseIf InStr(6, voiceFile.Item(j), "O", vbBinaryCompare) > 0 Then  'true when find O MEANS ON
       pos = InStr(6, voiceFile.Item(j), "O", vbBinaryCompare)    'GET POSITION OF O
       nextdate = Mid(voiceFile.Item(j), 1, pos - 1)     'TRUNCATE O AND SUFFIX TO FILTER DATE
       If DateValue(nextdate) = DateValue(Date) Then
         'MsgBox nextdate
          PlayVoiceMsg (voiceFile.Item(j))
       End If               'end of  =
  End If     'end of B/A/O checkings
 Next
End Sub

Private Sub PlayVoiceMsg(Voice As String)
Dim msgnm As String
msgnm = Voice
Voice = App.Path & "\messages\" & Voice & ".wav"
messageUsed = True
   cmd = "open " & Chr(34) & Voice & Chr(34) & " type waveaudio alias recwave"  ' & Chr(34) & tempfile & Chr(34) & " type waveaudio alias vin"
   errorCode = mciSendString(cmd, returnStr, 255, 0)
   If errorCode <> 0 Then MsgBox "Device failed "
   cmd = "play recwave from 0"
   errorCode = mciSendString(cmd, returnStr, 255, 0)
   If errorCode <> 0 Then MsgBox "Device failed to playback the file" & returnStr
   
   cmd = "status recwave length"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
   If errorCode <> 0 Then MsgBox "error " & returnStr
   voiceLength = Val(returnStr)
   Text1.Text = "Now Playing your voice message " & msgnm & " of Length " & voiceLength / 1000 & " Seconds"
   Dim sav As Double     'used to give status command rarely
   sav = 0
   Do
    If sav Mod voiceLength = 0 Then
      cmd = "status recwave position"
      errorCode = mciSendString(cmd, returnStr, 255, 0)
      If errorCode <> 0 Then Exit Sub 'MsgBox "error " & returnStr
    End If
    sav = sav + 1
     If isExit = True Then
       cmd = "close recwave"
       errorCode = mciSendString(cmd, returnStr, 255, 0)
       End   'means X is clicked when playing
     End If
    DoEvents
   Loop Until voiceLength = Val(returnStr)
   cmd = "close recwave "
   errorCode = mciSendString(cmd, returnStr, 255, 0)
   If errorCode <> 0 Then MsgBox "error " & returnStr
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim j As Integer
'For j = 1 To voiceFile.Count
 'cmdStop_Click
'Next
isExit = True
' cmd = "close recwave"
' errorCode = mciSendString(cmd, returnStr, 255, 0)
 'Unload Me
End Sub

Private Sub Text1_Click()

If Command2.Enabled = True Then
 Command2.SetFocus
Else
 cmdStop.SetFocus
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Command2.Enabled = True Then
 Command2.SetFocus
Else
 cmdStop.SetFocus
End If

End Sub
