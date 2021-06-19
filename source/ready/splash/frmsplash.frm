VERSION 5.00
Begin VB.Form WipesForm 
   Caption         =   "Image Wipes"
   ClientHeight    =   4425
   ClientLeft      =   1770
   ClientTop       =   2295
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9600
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4560
      Top             =   2880
   End
   Begin VB.CheckBox ClearDestination 
      Caption         =   "Clear destination before wipe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4950
      TabIndex        =   12
      Top             =   60
      Value           =   1  'Checked
      Width           =   3030
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
      TabIndex        =   9
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Vertical 
      Caption         =   "Vertical Blinds"
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
      TabIndex        =   11
      Top             =   4635
      Width           =   2055
   End
   Begin VB.CommandButton WipeRight 
      Caption         =   "Wipe From Right"
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
      Left            =   240
      TabIndex        =   10
      Top             =   4635
      Width           =   1935
   End
   Begin VB.CommandButton WipeRightLeft 
      Caption         =   "Wipe Right && Left"
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
      TabIndex        =   8
      Top             =   4185
      Width           =   1935
   End
   Begin VB.CommandButton WipeUpDown 
      Caption         =   "Wipe Up && Down"
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
      TabIndex        =   7
      Top             =   4635
      Width           =   1935
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
      TabIndex        =   6
      Top             =   5070
      Width           =   1935
   End
   Begin VB.CommandButton StretchBottom 
      Caption         =   "Stretch From Bottom"
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
      Left            =   4935
      TabIndex        =   5
      Top             =   4635
      Width           =   2085
   End
   Begin VB.CommandButton Horizontal 
      Caption         =   "Horizontal Blinds"
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
      TabIndex        =   4
      Top             =   4185
      Width           =   2055
   End
   Begin VB.CommandButton WipeLeft 
      Caption         =   "Wipe From Left"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4185
      Width           =   1935
   End
   Begin VB.CommandButton StretchRight 
      Caption         =   "Stretch From Right"
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
      Left            =   4935
      TabIndex        =   2
      Top             =   4185
      Width           =   2085
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   4950
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   1
      Top             =   345
      Width           =   4515
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3705
      Left            =   240
      Picture         =   "frmsplash.frx":0000
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   278
      TabIndex        =   0
      Top             =   360
      Width           =   4230
   End
End
Attribute VB_Name = "WipesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm" () As Long
Private StartTime As Long
Private TotalDuration As Integer

Private Sub Exit_Click()
    End
End Sub

Private Sub Timer1_Timer()
Horizontal_Click
End Sub

Private Sub Vertical_Click()
Dim Stripes As Integer
Dim i As Integer, j As Integer
Dim StripeHeight As Integer
Dim mseconds As Integer

    Picture2.Cls
    stripewidth = 20
    Stripes = Picture1.ScaleWidth / stripewidth
    On Error Resume Next
    mseconds = TotalDuration / stripewidth
    For j = 1 To stripewidth
        StartDelay
        For i = 0 To Stripes
            Picture2.PaintPicture Picture1.Picture, i * stripewidth, 0, _
                j, Picture1.ScaleHeight, _
                i * stripewidth, 0, _
                j, Picture1.ScaleHeight, &HCC0020
        Next
        EndDelay (mseconds)
    Next

End Sub

Private Sub Form_Load()
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
    Picture2.Top = Picture1.Top
    TotalDuration = 2000
  '  Horizontal_Click
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    CommonDialog1.Filter = "Images|*.BMP;*.GIF;*.JPG"
    CommonDialog1.Action = 1
    If CommonDialog1.FileName = "" Then Exit Sub
    Picture1.Picture = LoadPicture(CommonDialog1.FileName)
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    CommonDialog1.Filter = "Images|*.BMP;*.GIF;*.JPG"
    CommonDialog1.Action = 1
    If CommonDialog1.FileName = "" Then Exit Sub
    Picture2.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub StretchRight_Click()
Dim X As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For X = 1 To Picture1.ScaleWidth Step 3
        Picture2.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, X, _
        Picture1.ScaleHeight, &HCC0020
    Next

End Sub

Private Sub WipeLeft_Click()
Dim X As Integer
Dim mseconds As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    mseconds = TotalDuration / Picture1.ScaleWidth
    For X = 1 To Picture1.ScaleWidth
        StartDelay
        Picture2.PaintPicture Picture1.Picture, 0, 0, _
        X, Picture1.ScaleHeight, 0, 0, X, _
        Picture1.ScaleHeight, &HCC0020
        EndDelay (mseconds)
    Next

End Sub

Private Sub StretchBottom_Click()
Dim X As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For X = 1 To Picture1.ScaleHeight Step 3
        Picture2.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, _
        Picture1.ScaleWidth, X, &HCC0020
    Next
End Sub

Private Sub WipeCenter_Click()
Dim PWidth As Integer, PHeight As Integer
Dim i As Integer
Dim mseconds As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    If Picture1.ScaleWidth > Picture1.ScaleHeight Then
        PWidth = Picture1.ScaleWidth - Picture1.ScaleHeight
        PHeight = 1
    ElseIf Picture1.ScaleWidth < Picture1.ScaleHeight Then
        PWidth = 1
        PHeight = Picture1.ScaleHeight - Picture1.ScaleWidth
    Else
        PWidth = 1
        PHeight = 1
    End If

    mseconds = TotalDuration / (Picture1.ScaleWidth - PWidth)
    For i = 1 To Picture1.ScaleWidth - PWidth
        StartDelay
        Picture2.PaintPicture Picture1.Picture, _
        Int((Picture1.ScaleWidth - PWidth) / 2), Int((Picture1.ScaleHeight - PHeight) / 2), _
        PWidth, PHeight, _
        Int((Picture1.ScaleWidth - PWidth) / 2), Int((Picture1.ScaleHeight - PHeight) / 2), _
        PWidth, PHeight, &HCC0020
        PWidth = PWidth + 1
        PHeight = Height + 1
        EndDelay (mseconds)
    Next

End Sub

Private Sub Horizontal_Click()
Dim Stripes As Integer
Dim i As Integer, j As Integer
Dim StripeHeight As Integer
Dim mseconds As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    StripeHeight = 20
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

End Sub


Private Sub WipeRight_Click()
Dim X As Integer
Dim mseconds As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    mseconds = TotalDuration / Picture1.ScaleWidth
    For X = 1 To Picture1.ScaleWidth
        StartDelay
        Picture2.PaintPicture Picture1.Picture, _
        Picture1.ScaleWidth - X, 0, _
        X, Picture1.ScaleHeight, _
        Picture1.ScaleWidth - X, 0, _
        X, Picture1.ScaleHeight, &HCC0020
        EndDelay mseconds
    Next

End Sub

Private Sub WipeUpDown_Click()
Dim PWidth As Integer, PHeight As Integer
Dim i As Integer
Dim mseconds As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    PWidth = Picture1.ScaleWidth
    PHeight = 1
    mseconds = TotalDuration / (Picture1.ScaleHeight / 2)
    For i = 1 To Picture1.ScaleHeight / 2
        StartDelay
        Picture2.PaintPicture Picture1.Picture, _
        0, (Picture1.ScaleHeight - PHeight) / 2, _
        PWidth, PHeight, _
        0, (Picture1.ScaleHeight - PHeight) / 2, _
        PWidth, PHeight, &HCC0020
        PHeight = PHeight + 2
        EndDelay (mseconds)
    Next
End Sub

Private Sub WipeRightLeft_Click()
Dim PWidth As Integer, PHeight As Integer
Dim i As Integer
Dim mseconds As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    PWidth = 1
    PHeight = Picture1.ScaleHeight
    mseconds = TotalDuration / (Picture1.ScaleWidth / 2)
    For i = 1 To Picture1.ScaleWidth / 2
        StartDelay
        Picture2.PaintPicture Picture1.Picture, _
        (Picture1.ScaleWidth - PWidth) / 2, 0, _
        PWidth, PHeight, _
        (Picture1.ScaleWidth - PWidth) / 2, 0, _
        PWidth, PHeight, &HCC0020
        PWidth = PWidth + 2
        EndDelay (mseconds)
    Next

End Sub

Sub StartDelay()
    StartTime = timeGetTime()
End Sub

Sub EndDelay(N As Integer)
    While timeGetTime() - StartTime < N
    Wend
End Sub

