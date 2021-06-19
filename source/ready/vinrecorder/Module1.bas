Attribute VB_Name = "Module1"
Option Explicit

Public errorCode As Integer
Public returnStr As String * 255
Public cmd As String * 255
Public songfilename As String   'which is opened
Public tempfile As String   'point to playing file
Public Declare Function mciSendString Lib "winmm.dll" _
    Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long

Public songlength As Long    'store the length of song

Public Sub playsong()
    ' make sure that device with the DAYS alias is open
    cmd = "close vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    ' now open the DAYS.WAV file as DAYS
    cmd = "open " & Chr(34) & tempfile & Chr(34) & " type waveaudio alias vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    'set vin  time format samples
    'play the song
    errorCode = mciSendString("play vin", returnStr, 255, 0)
    'get the length of song
    cmd = "status vin length"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    songlength = Val(returnStr)
    'MsgBox songlength
    If errorCode <> 0 Then
        MsgBox "There was an error on opening the file." & songfilename & vbCrLf _
               & "Please make sure the  file " & songfilename & " in the data folder of the application" & _
               "Or your system may not support MIDI Sequencer. Use troubleshoot"
        Exit Sub
    End If
   

End Sub

Public Sub closesong()
  cmd = "close vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    
     If errorCode <> 0 Then
        MsgBox "There was an error while closing file." & songfilename & vbCrLf _
               & "Please make sure the " & songfilename & "file in the data folder of the application" & _
               "Or your system may not support MIDI Sequencer. Use troubleshoot"
        Exit Sub
    End If
End Sub

