VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Image Wipes"
   ClientHeight    =   2190
   ClientLeft      =   4305
   ClientTop       =   4560
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4080
      Top             =   1440
   End
   Begin VB.CommandButton Exit 
      Caption         =   "E X I T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7380
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton WipeCenter 
      Caption         =   "Wipe From Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2580
      TabIndex        =   2
      Top             =   5070
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2160
      Left            =   0
      Picture         =   "splash.frx":0000
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm" () As Long
Private StartTime As Long
Private TotalDuration As Integer
Dim wipeOnce As Boolean
Dim counter As Integer
Dim StripeHeight As Integer

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40

Private Sub Exit_Click()
    End
End Sub

Private Sub Timer1_Timer()
If wipeOnce = True Then
Dim Stripes As Integer
Dim i As Integer, j As Integer

Dim mseconds As Integer

   Picture2.Picture = LoadPicture()

    Stripes = Fix(Picture1.ScaleHeight / StripeHeight)
    On Error Resume Next
    mseconds = TotalDuration / StripeHeight
    For j = 1 To StripeHeight
        StartDelay
        For i = 0 To Stripes
            Picture2.PaintPicture Picture1.Picture, 0, i * StripeHeight, _
            Picture1.ScaleWidth, j, _
            0, i * StripeHeight, _
            Picture1.ScaleWidth, j, &HCC0020
        Next
        EndDelay (mseconds)
    Next
    wipeOnce = False
    counter = 0
End If
If wipeOnce = False And counter < 8 Then counter = counter + 1
If counter = 8 Then
 Load frmMain
 frmMain.Visible = True
 Unload Me
End If
End Sub


Private Sub Form_Load()
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
    Picture2.Top = Picture1.Top
    TotalDuration = 4000
    wipeOnce = True
    If Screen.Width > 15000 Then        'for 1024X 768
      SetWindowPos Me.hwnd, HWND_TOPMOST, 350, 310, 305, 145, SWP_SHOWWINDOW
          StripeHeight = 10
    Else                                                    '800 X 600
       SetWindowPos Me.hwnd, HWND_TOPMOST, 260, 210, 305, 145, SWP_SHOWWINDOW
           StripeHeight = 6
    End If
  '  Horizontal_Click
End Sub



Sub StartDelay()
    StartTime = timeGetTime()
End Sub

Sub EndDelay(N As Integer)
    While timeGetTime() - StartTime < N
    Wend
End Sub

