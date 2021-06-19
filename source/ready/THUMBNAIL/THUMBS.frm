VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThumb 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PAINTPICTURE Demo"
   ClientHeight    =   7725
   ClientLeft      =   285
   ClientTop       =   480
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7725
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2400
      TabIndex        =   29
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtsrcWH 
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optRes 
      Caption         =   "48 X 48"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   24
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton optRes 
      Caption         =   "32 X 32"
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   23
      Top             =   0
      Width           =   1095
   End
   Begin VB.HScrollBar SliderDestWidth 
      Height          =   135
      Left            =   9840
      Max             =   400
      Min             =   1
      TabIndex        =   22
      Top             =   480
      Value           =   1
      Width           =   1095
   End
   Begin VB.VScrollBar SliderDestHeight 
      Height          =   615
      Left            =   9720
      Max             =   300
      Min             =   1
      TabIndex        =   21
      Top             =   0
      Value           =   1
      Width           =   135
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   211
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   211
   End
   Begin VB.CommandButton tyty 
      Caption         =   "Command3"
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   720
      Width           =   1215
   End
   Begin VB.VScrollBar SliderSourceHeight 
      Height          =   1215
      Left            =   0
      Max             =   420
      Min             =   1
      TabIndex        =   17
      Top             =   4200
      Value           =   1
      Width           =   135
   End
   Begin VB.HScrollBar SliderSourceWidth 
      Height          =   135
      Left            =   120
      Max             =   560
      Min             =   1
      TabIndex        =   16
      Top             =   5280
      Value           =   1
      Width           =   1095
   End
   Begin VB.CheckBox Rasternegative 
      Height          =   225
      Left            =   2760
      TabIndex        =   9
      Top             =   0
      Width           =   210
   End
   Begin VB.CommandButton SSCommand2 
      Caption         =   "Copy Image"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5760
      TabIndex        =   8
      Top             =   4080
      Width           =   1650
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse.."
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   5760
      Width           =   855
   End
   Begin VB.PictureBox OUTPUT 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   5880
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   5
      Top             =   1080
      Width           =   5415
   End
   Begin VB.CommandButton RTTY 
      Caption         =   "Command2"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REFRESH DEST"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin MSComctlLib.Slider SliderSourceX 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   5040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   10
      Max             =   1024
      SelStart        =   1
      TickFrequency   =   50
      Value           =   1
   End
   Begin VB.PictureBox Source 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   120
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   0
      Top             =   1080
      Width           =   5745
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   7320
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmap|*.BMP;*.DIB"
      FontSize        =   1.17491e-38
   End
   Begin MSComctlLib.Slider SliderSourceY 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   5400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   10
      Max             =   768
      TickFrequency   =   50
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   341
      ScaleMode       =   0  'User
      ScaleWidth      =   263.203
      TabIndex        =   10
      Top             =   0
      Width           =   4335
      Begin VB.Label lblSourceX 
         Caption         =   "on X-axis"
         Height          =   270
         Left            =   240
         TabIndex        =   15
         Top             =   4920
         Width           =   915
      End
      Begin VB.Label lblSourceY 
         Caption         =   "on Y-axis"
         Height          =   210
         Left            =   270
         TabIndex        =   14
         Top             =   5295
         Width           =   1215
      End
      Begin VB.Label lblSourceWidth 
         Caption         =   "Width"
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   6105
         Width           =   1170
      End
      Begin VB.Label lblSourceHeight 
         Caption         =   "Height"
         Height          =   270
         Left            =   240
         TabIndex        =   12
         Top             =   6525
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Change Selection Area"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   5760
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Full image view"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Negative"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Convert Selected Portion of image"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RasterOp As Long
Dim SourceX As Single, SourceY As Single
Dim SourceWidth As Single, SourceHeight As Single
Dim DestX As Single, DestY As Single
Dim DestWidth As Single, DestHeight As Single
Dim raster(2) As Long
Dim prevx As Single, prevy As Single
Dim isPictureLoaded As Boolean 'true when pic opened



Private Sub Check1_Click()

If isPictureLoaded = True Then
Source.PaintPicture Picture1.Picture, 0, 0, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
OUTPUT.Cls
OUTPUT.PaintPicture Picture1.Picture, 0, 0, OUTPUT.ScaleWidth, OUTPUT.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, RasterOp
End If
End Sub

Private Sub cmdSave_Click()
SavePicture OUTPUT.Image, "c:\temp.ico"
'SavePicture OUTPUT.Image, "c:\temp.jpg"
End Sub

Private Sub Command1_Click()
'OUTPUT.Refresh
OUTPUT.Cls
End Sub

Private Sub Command2_Click()

     'Action = 1
    CMDialog1.InitDir = "D:\poster\pictures\" 'App.Path
    CMDialog1.Filter = "Imagefiles | *.BMP;*.jpg;*.gif;*.dib"
    CMDialog1.ShowOpen
    If CMDialog1.FileName = "" Then Exit Sub
    'On Error GoTo Error1
    Picture1.Picture = LoadPicture(CMDialog1.FileName)
    frmThumb.Refresh
    
   'Source.PaintPicture Picture1.Picture, 0, 0, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
   isPictureLoaded = True
   fullSizeInSource
   SliderDestWidth.Max = 400 'Picture1.ScaleWidth
   SliderDestHeight.Max = 300 'Picture1.ScaleHeight
   
   Exit Sub
Error1:
    MsgBox "Couldn't open file " + CMDialog1.FileName
    Exit Sub
End Sub



Private Sub Command3_Click()
    
    CMDialog1.Action = 1
    CMDialog1.InitDir = App.Path
    If CMDialog1.FileName = "" Then Exit Sub
    On Error GoTo Error1
    OUTPUT.Picture = LoadPicture(CMDialog1.FileName)
    Exit Sub

Error1:
    MsgBox "Couldn't open file " + CMDialog1.FileName
    Exit Sub
End Sub

Private Sub Command4_Click()

    OUTPUT.Picture = LoadPicture("")

End Sub

Private Sub Form_Load()
'    Me.Width = OUTPUT.Left + OUTPUT.Width + 25 * Screen.TwipsPerPixelX
'MsgBox Screen.Width
If Screen.Width = 9000 Then
 Picture1.Width = frmThumb.ScaleX(800, vbPixels, vbTwips)
    Picture1.Height = frmThumb.ScaleY(600, vbPixels, vbTwips)
ElseIf Screen.Width = 15360 Then
 
 Picture1.Width = frmThumb.ScaleX(1024, vbPixels, vbTwips)
    Picture1.Height = frmThumb.ScaleY(768, vbPixels, vbTwips)
    SliderDestWidth.Max = 560
    SliderDestHeight.Max = 420
End If

    ChDir App.Path
    On Error Resume Next
    
    raster(0) = &HCC0020    ' SRCCOPY   (COPY PEN)
    
    raster(1) = &H330008   ' NOTSRCCOPY
    isPictureLoaded = False


    RasterOp = raster(0)
    SliderSourceX_Change
    SliderSourceY_Change
   SourceWidth = 383
   SourceHeight = 278
    SliderDestWidth.Value = 205
    SliderDestHeight.Value = 140
     SliderSourceWidth.Value = Source.ScaleWidth
    SliderSourceHeight.Value = Source.ScaleHeight
    SliderDestX.Value = 20
    SliderDestY.Value = 20

End Sub

Private Sub lblDestHeight_Click()

End Sub

Private Sub optRes_Click(Index As Integer)
optRes(Index).Value = True
If Index = 0 Then
 optRes(1).Value = False
 DestWidth = 32
 DestHeight = 32
    OUTPUT.Width = frmThumb.ScaleX(32, vbPixels, vbTwips)
    OUTPUT.Height = frmThumb.ScaleY(DestHeight, vbPixels, vbTwips)
Else
 optRes(0).Value = False
 DestWidth = 48
 DestHeight = 48
 OUTPUT.Width = frmThumb.ScaleX(DestWidth, vbPixels, vbTwips)
     OUTPUT.Height = frmThumb.ScaleY(DestHeight, vbPixels, vbTwips)

End If
 fullSizeInDestination
End Sub


Private Sub Rasternegative_Click()

If Rasternegative.Value = vbChecked Then
RasterOp = raster(1)
 'MsgBox "here"
 Else 'If Rasternegative.Value = False Then
 RasterOp = raster(0)
'
End If
End Sub

Private Sub RTTY_Click()
'OUTPUT.Width = frmThumb.ScaleX(500, vbPixels, vbTwips)
'OUTPUT.Height = frmThumb.ScaleY(400, vbPixels, vbTwips)
OUTPUT.Picture = Picture1.Picture
'MsgBox Screen.Width
End Sub

Private Sub SliderDestHeight_Scroll()
    DestHeight = SliderDestHeight.Value
  '  lblDestHeight.Caption = "Height" + Space$(4) + Format$(DestHeight, "000")
End Sub

Private Sub SliderDestWidth_Scroll()
    DestWidth = SliderDestWidth.Value
'    lblDestWidth.Caption = "Width" + Space$(4) + Format$(DestWidth, "000")
End Sub

Private Sub SliderDestX_Scroll()
    DestX = 0 'SliderDestX.Value
   ' lblDestX.Caption = "Dest X" + Space$(2) + Format$(DestX, "000")
End Sub

Private Sub SliderDestY_Scroll()
    DestY = 0 'SliderDestY.Value
    'lblDestY.Caption = "Dest Y" + Space$(2) + Format$(DestY, "000")
End Sub

Private Sub SliderSourceHeight_Scroll()
    'SourceHeight = SliderSourceHeight.Value
    'lblSourceHeight.Caption = "Height" + Space$(2) + Format$(SourceHeight, "000")
End Sub

Private Sub SliderSourceWidth_Scroll()
   ' SourceWidth = SliderSourceWidth.Value
    'lblSourceWidth.Caption = "Width" + Space$(2) + Format$(SourceWidth, "000")
End Sub

Private Sub SliderSourceX_Change()
    SourceX = SliderSourceX.Value
    'lblSourceX.Caption = "Source X" + Space$(2) + Format$(SourceX, "000")
End Sub

Private Sub SliderSourceX_Scroll()
'    SourceX = SliderSourceX.Value
 '   lblSourceX.Caption = "Source X" + Space$(2) + Format$(SourceX, "000")
End Sub

Private Sub SliderSourceY_Change()
    SourceY = SliderSourceY.Value
   ' lblSourceY.Caption = "Source Y" + Space$(2) + Format$(SourceY, "000")
End Sub

Private Sub SliderSourceWidth_Change()
'    SourceWidth = SliderSourceWidth.Value
'     Source.Width = frmThumb.ScaleX(SourceWidth, vbPixels, vbTwips)
  '  lblSourceWidth.Caption = "Width" + Space$(2) + Format$(SourceWidth, "000")
  If isPictureLoaded = True Then
Source.Cls
Source.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.Width, Picture1.Height, _
        SliderSourceWidth.Value, SliderSourceHeight.Value, _
        Picture1.Width, Picture1.Height, RasterOp

'    fullSizeInDestination
'OUTPUT.Cls
'    OUTPUT.PaintPicture Picture1.Picture, DestX, DestY, _
  '  DestWidth, DestHeight, _
 '   SliderSourceWidth.Value, SliderSourceHeight.Value, _
  '  SourceWidth, SourceHeight, RasterOp
    End If
End Sub

Private Sub SliderSourceHeight_Change()
'    SourceHeight = SliderSourceHeight.Value
 '   Source.Height = frmThumb.ScaleY(SourceHeight, vbPixels, vbTwips)
  '  lblSourceHeight.Caption = "Height" + Space$(2) + Format$(SourceHeight, "000")
   ' fullSizeInSource
   If isPictureLoaded = True Then
   Source.Cls
Source.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.Width, Picture1.Height, _
        SliderSourceWidth.Value, SliderSourceHeight.Value, _
        Picture1.Width, Picture1.Height, RasterOp

   ' fullSizeInDestination
 '  OUTPUT.Cls
  '  OUTPUT.PaintPicture Picture1.Picture, DestX, DestY, DestWidth, DestHeight, _
    SliderSourceWidth.Value, SliderSourceHeight.Value, SourceWidth, SourceHeight, RasterOp

   End If
End Sub

Private Sub SliderDestX_Change()
  '  DestX = SliderDestX.Value
   ' lblDestX.Caption = "Dest X" + Space$(2) + Format$(DestX, "000")
End Sub

Private Sub SliderDestY_Change()
    'DestY = SliderDestY.Value
    'lblDestY.Caption = "Dest Y" + Space$(2) + Format$(DestY, "000")
End Sub

Private Sub SliderDestWidth_Change()
    DestWidth = SliderDestWidth.Value
    OUTPUT.Width = frmThumb.ScaleX(DestWidth, vbPixels, vbTwips)
'    If frmThumb.Width > 12000 Then frmThumb.Width = frmThumb.Width + frmThumb.ScaleX(DestWidth, vbPixels, vbTwips)

    txtsrcWH.Text = Format$(DestWidth, "000") & " X " & Format$(DestHeight, "000")
    fullSizeInDestination
End Sub

Private Sub SliderDestHeight_Change()
    DestHeight = SliderDestHeight.Value
    
    OUTPUT.Height = frmThumb.ScaleY(DestHeight, vbPixels, vbTwips)
    txtsrcWH.Text = Format$(DestWidth, "000") & " X " & Format$(DestHeight, "000")
    fullSizeInDestination
End Sub

Private Sub SSCommand1_Click()
    OUTPUT.PaintPicture Source.Picture, DestX, DestY, DestWidth, DestHeight, SourceX, SourceY, SourceWidth, SourceHeight, RasterOp
End Sub


Private Sub SliderSourceY_Scroll()
    'SourceY = SliderSourceY.Value
    'lblSourceY.Caption = "Source Y" + Space$(2) + Format$(SourceY, "000")
End Sub

Private Sub Source_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Text1.Text = Str(X) & "   GDFG " & Str(Y) & Str(Picture2.Top) & "GFDG" & Str(Picture2.Left)
'  fullSizeInDestination
Dim newx As Long, newy As Long
newx = prevx + X
newy = prevy + Y
If isPictureLoaded = True Then
Source.MousePointer = 5 'Val(Text1.Text)
If newx > -5 And newy > -5 Then 'And Picture1.ScaleWidth > newx + Source.ScaleWidth And Picture1.Scaleheight > newy + Source.ScaleHeight Then   ' Source.ScaleTop And Y > Source.ScaleLeft And X < Source.ScaleTop + Source.ScaleHeight And Y < Source.ScaleLeft + Source.ScaleWidth Then
If Button = 1 Then
Source.Cls
Source.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, _
        newx, newy, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, RasterOp
SliderSourceX.Value = newx
SliderSourceY.Value = newy
Text1.Text = newx & " Y " & newy
'SSCommand2_Click
fullSizeInDestination
End If
End If
End If
End Sub

Private Sub Source_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
prevx = X
prevy = Y
End Sub

Private Sub SSCommand2_Click()
'OUTPUT.Refresh
'OUTPUT.Cls
 '   OUTPUT.PaintPicture Picture1.Picture, DestX, DestY, DestWidth, DestHeight, SourceX, SourceY, SourceWidth, SourceHeight, RasterOp
 fullSizeInDestination
End Sub

Private Sub OUTPUT_Click()
    OUTPUT.Picture = LoadPicture("")
End Sub

Private Sub fullSizeInSource()
If isPictureLoaded = True Then Source.PaintPicture Picture1.Picture, 0, 0, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
End Sub
Private Sub fullSizeInDestination()
 If isPictureLoaded = True Then
   OUTPUT.Cls
    OUTPUT.PaintPicture Picture1.Picture, DestX, DestY, DestWidth, DestHeight, SourceX, SourceY, SourceWidth, SourceHeight, RasterOp
End If
End Sub

Private Sub tyty_Click()
Source.Cls
Source.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, _
        SliderSourceWidth.Value, SliderSourceHeight.Value, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, RasterOp

    fullSizeInSource

End Sub
