VERSION 5.00
Begin VB.Form speak 
   Caption         =   "Play Days"
   ClientHeight    =   3735
   ClientLeft      =   1140
   ClientTop       =   1470
   ClientWidth     =   4140
   LinkTopic       =   "PlayWave"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   4140
   Begin VB.CommandButton Play 
      Caption         =   "12"
      Height          =   495
      Index           =   11
      Left            =   2280
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "11"
      Height          =   495
      Index           =   10
      Left            =   2280
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "10"
      Height          =   495
      Index           =   9
      Left            =   2280
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "9"
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "8"
      Height          =   495
      Index           =   7
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "7"
      Height          =   495
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Play 
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "speak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim errorCode As Integer
Dim returnStr As String * 255
Dim cmd As String * 255

Private Declare Function mciSendString Lib "winmm.dll" _
    Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long
Dim songlength As Long
    



Private Sub Form_Unload(Cancel As Integer)
    
    cmd = "close vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    'MsgBox "sONG CLOSED " & errorCode
    ' If errorCode <> 0 Then
     '   MsgBox "There was an error on opening the vin.WAV file." & vbCrLf _
               & "Please make sure the vin.WAV file in the same folder as the application"
    '    Exit Sub
   ' End If
End Sub

Private Sub Play_Click(Index As Integer)
'Dim errorCode As Integer    getting  error so i put them in option explicit
'Dim returnStr As Integer
'Dim cmd As String * 255
    
    ' make sure that device with the vin alias is open
    cmd = "close vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    ' now open the vin.WAV file as vin
    cmd = "open " & Chr(34) & App.Path & "\num.wav " & Chr(34) & " type waveaudio alias vin"
    errorCode = mciSendString(cmd, returnStr, 255, 0)
    If errorCode <> 0 Then
        MsgBox "There was an error on opening the num.WAV file." & vbCrLf _
               & "Please make sure the num.WAV file in the same folder as the application"
        Exit Sub
    End If
    Index = Index + 1
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

