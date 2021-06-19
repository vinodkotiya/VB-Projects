VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5370
   ClientLeft      =   3600
   ClientTop       =   2820
   ClientWidth     =   6390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FriendsChatSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   4200
         Top             =   120
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "FriendsChatSplash.frx":1272
         Top             =   120
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF8080&
         X1              =   0
         X2              =   6360
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mails at sonal3k@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1875
         TabIndex        =   4
         Top             =   4560
         Width           =   2505
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         Caption         =   "  Friends Chat !"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   600
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   3405
      End
      Begin VB.Image Image1 
         Height          =   2250
         Left            =   1560
         Picture         =   "FriendsChatSplash.frx":16B4
         Top             =   240
         Width           =   3000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         X1              =   0
         X2              =   6360
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Sonal Dubey, Vinita Shrivastava && Rupali Nagle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Friends Chat ! © 2003   All rights reserved."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   3960
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim i As Integer, k As Integer, j As Integer, l As Integer

Private Sub Form_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
SetWindowRgn hWnd, CreateEllipticRgn(30, 15, 390, 350), True
Load Form1
DoEvents
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
k = 10
l = 10
For i = 410 To 200 Step -5
For j = 410 To 200 Step -5
k = k + 5
l = l + 5
SetWindowRgn hWnd, CreateEllipticRgn(k, l, i, j), True
DoEvents
frmSplash.Refresh
Next
frmSplash.Refresh
DoEvents
Next
Form1.Show
Form1.Refresh
Unload Me
End Sub

Private Sub Timer1_Timer()
Image2.Left = Image2.Left + 200
Image2.Top = Image2.Top + 200
If Image2.Left >= 3120 Then Image2.Left = 0
If Image2.Top >= 4120 Then Image2.Top = 0
End Sub

Private Sub Frame1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Image1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Image2_Click()
k = 10
l = 10
For i = 410 To 200 Step -5
For j = 410 To 200 Step -5
k = k + 5
l = l + 5
SetWindowRgn hWnd, CreateEllipticRgn(k, l, i, j), True
DoEvents
frmSplash.Refresh
Next
frmSplash.Refresh
DoEvents
Next
Form1.Show
Form1.Refresh
Unload Me
End Sub

Private Sub Label1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Label3_Click()
Form1.Show
Unload Me
End Sub
