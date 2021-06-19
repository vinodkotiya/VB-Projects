VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   9855
   Begin VB.CommandButton Command3 
      Caption         =   "Copy Fast"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   -120
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub Command3_Click()
   
    Picture2.Cls
    Screen.MousePointer = vbHourglass
' set up source bitmap
    hBMPSource = CreateCompatibleBitmap(Picture1.hdc, _
           Picture1.ScaleWidth, Picture1.ScaleHeight)
    hSourceDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hSourceDC, hBMPSource
' set up destination bitmap
    hBMPDest = CreateCompatibleBitmap(Picture2.hdc, _
           Picture2.ScaleWidth, Picture2.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture2.hdc)
    SelectObject hDestDC, hBMPDest
' Copy picture bitmap to source bitmap
    BitBlt hSourceDC, 0, 0, Picture1.ScaleWidth - 1, _
           Picture1.ScaleHeight - 1, Picture1.hdc, 0, 0, &HCC0020
' Copy pixels between bitmaps
    For i = 0 To Picture1.ScaleWidth - 1
        For j = 0 To Picture1.ScaleHeight - 1
            clr = GetPixel(hSourceDC, i, j)
            SetPixel hDestDC, i, j, clr
        Next
    Next
' transfer the copied pixels to the second PictureBox
    BitBlt Picture2.hdc, 0, 0, Picture1.ScaleWidth - 1, _
           Picture1.ScaleHeight - 1, hDestDC, 0, 0, &HCC0020
    'Picture2.Refresh
' finally, clean up memory
    Call DeleteDC(hSourceDC)
    Call DeleteObject(hBMPSource)
    Call DeleteDC(hDestDC)
    Call DeleteObject(hBMPDest)
    Screen.MousePointer = vbDefault
End Sub
End Sub

Private Sub Form_Load()
Form1.Top = 0
Form1.Left = 0
Picture2.Height = Screen.Height
Picture2.Width = Screen.Width
Form1.Height = Screen.Height
Form1.Width = Screen.Width
End Sub

