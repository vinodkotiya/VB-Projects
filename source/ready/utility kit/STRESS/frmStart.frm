VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2445
   ClientLeft      =   1530
   ClientTop       =   2505
   ClientWidth     =   7125
   DrawStyle       =   3  'Dash-Dot
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStart.frx":030A
   ScaleHeight     =   2445
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Credit"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "VIN Utility Kit"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdpreview 
      BackColor       =   &H00FFFF00&
      Caption         =   "PREVIEW"
      Height          =   615
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Stop"
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "SET IT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "5"
      Top             =   840
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   1
      X1              =   240
      X2              =   6960
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   0
      X1              =   240
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      Index           =   1
      X1              =   240
      X2              =   6960
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      Index           =   0
      X1              =   240
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Have A Break !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Give me a  break after every             Minutes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6255
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////////////////
'/////////////////  Programmer: VINOD KOTIYA //////////////////////////
'///////////////// created on jan-2003 //////////////////////////////////
'///////////////// providing free on http://vinodkotiya.tripod.com//////
'//////////////// help to promote the site if u want ///////////////////////////////
'///////////////////////////////////////////////////////////////////////////////
'///////////////////   Made In India  (BHOPAL) /////////////////////////////////////////
Option Explicit
Dim Timeprev As String




Private Sub cmdOk_Click()
If Trim(Text1.Text) = "" Then
 MsgBox "Please enter your breaktime"
 Exit Sub
End If

Timer1.Interval = 30000 'Val(Text1.Text) * 60000


If Val(Text1.Text) > 59 Then
 MsgBox "It is harmful for your eyes and health if you want to work " _
 & Text1.Text & " Minutes continuously on system"
 Text1.Text = "59"
End If
MsgBox "Now you can take a break after every " & Text1.Text & " Minutes"
Timeprev = Time
frmStart.Hide
End Sub

Private Sub Command1_Click()
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

Private Sub Command2_Click()
Dim temp As String
temp = MsgBox("If you are working on computer for a long while , then this utility is better for you." & Chr(13) & _
"It redirects your working screen to another display after some time set by you." & Chr(13) & _
"Thus preventing your eyes from a continuous radiation " & Chr(13) & _
"It is better than a screen saver because a screen saver only saves " & Chr(13) & _
"your screen from phosphores burning and activate only when there " & Chr(13) & _
"is nothing happen on your system. But this utility PopUps when" & Chr(13) & _
"you are working. Tip:- Always Try to work on higher Resolutions" & Chr(13) & Chr(13) & _
"Do you want to terminate the application ", vbYesNo)
If temp = vbYes Then
 Dim tempa As Long
On Error GoTo Exeerror
'Unload Me
tempa = Shell(App.Path & "\about.exe", vbNormalFocus)
End
Exit Sub
Exeerror:
 MsgBox "Application 'about.EXE' is not found in its " _
  & "Default directory about.exe "
 End
 
End If
End Sub


Private Sub cmdpreview_Click()
 ' Make sure there isn't another one running.
            CheckShouldRun

            ' Display the cover form.
            Load Form1
            Form1.Show
           ' ShowCursor False
End Sub

Private Sub Command3_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell("credit.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'CREDIT.EXE' is not found in its " _
  & "Default directory CREDIT.exe "
End Sub

Private Sub Form_Load()
frmStart.Show
frmStart.Refresh
cmdOk.SetFocus
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If IsNumeric(Chr(KeyCode)) = False Then Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
Dim diff As String
diff = TimeValue(Time) - TimeValue(Timeprev)
If Val(Text1.Text) = Minute(Time) - Minute(Timeprev) Or Val(Text1.Text) = 60 + Minute(Time) - Minute(Timeprev) Then
'MsgBox Minute(Time) & Minute(Timeprev)

'If Timer1.Interval > 0 Then
'MsgBox "vin"
'End If
 ' Make sure there isn't another one running.
    Timer1.Interval = 0
            CheckShouldRun

            ' Display the cover form.
            Load Form1
            Form1.Show
           ' ShowCursor False
 End If
 DoEvents
 
 
End Sub
