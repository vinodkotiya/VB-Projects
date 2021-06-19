VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIndradhanush 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDRA-DHANUSH"
   ClientHeight    =   7500
   ClientLeft      =   3360
   ClientTop       =   -1200
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmIndradhanush.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   6435
   Begin VB.HScrollBar hsbRed 
      Height          =   375
      Left            =   840
      Max             =   255
      TabIndex        =   16
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Technic"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Text            =   "Blue"
      Top             =   6460
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Technic"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Text            =   "Green"
      Top             =   5500
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Technic"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Text            =   "Red"
      Top             =   4540
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmIndradhanush.frx":1CCA
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Text            =   "COLOR NO."
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "COPY CODE"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00000000&
      Height          =   3135
      Left            =   0
      MouseIcon       =   "frmIndradhanush.frx":1CDD
      MousePointer    =   99  'Custom
      ScaleHeight     =   3075
      ScaleWidth      =   3435
      TabIndex        =   9
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "COPY COLOR"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Hcode 
      BackColor       =   &H00FFFF00&
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Dcode 
      BackColor       =   &H00FFFF00&
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CUSTOM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Picture         =   "frmIndradhanush.frx":39A7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtBlue 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Text            =   "Blue"
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtGreen 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Text            =   "Green"
      Top             =   5520
      Width           =   975
   End
   Begin VB.HScrollBar hsbBlue 
      Height          =   375
      Left            =   840
      Max             =   255
      TabIndex        =   2
      Top             =   6480
      Width           =   3135
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   375
      Left            =   840
      Max             =   255
      TabIndex        =   1
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox txtRed 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   4200
      TabIndex        =   0
      Text            =   "Red"
      Top             =   4560
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Set on top of All Windows"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      MaskColor       =   &H000000FF&
      TabIndex        =   17
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   3480
      MouseIcon       =   "frmIndradhanush.frx":4871
      MousePointer    =   99  'Custom
      Picture         =   "frmIndradhanush.frx":653B
      Top             =   120
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      MouseIcon       =   "frmIndradhanush.frx":E905
      MousePointer    =   99  'Custom
      Picture         =   "frmIndradhanush.frx":105CF
      Top             =   0
      Width           =   6450
   End
End
Attribute VB_Name = "frmIndradhanush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Topx As Integer, Lefty As Integer
Dim Rl As Long, Gl As Long, Bl As Long
Dim Color As String
Dim Rs As String, Gs As String, Bs As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_SHOWWINDOW = &H40



Private Sub Check1_Click()
If Check1.Value = vbChecked Then
Dim retValue As Long
    'Load Form1
    retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, 360, 0, _
               436, 532, SWP_SHOWWINDOW)
ElseIf Check1.Value = vbUnchecked Then
   Dim reetValue As Long
   
    reetValue = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 200, 10, _
               436, 532, SWP_SHOWWINDOW)
End If

End Sub

Private Sub Command1_Click()
Dim CDFlags As Long
On Error GoTo ColorError

    'CDFlags = 0
    'For i = 0 To 3
    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)
    'Next
    CommonDialog1.Flags = CDFlags
    CommonDialog1.Color = picDisplay.BackColor
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    picDisplay.BackColor = CommonDialog1.Color
    Dcode.Text = CommonDialog1.Color
    Hcode.Text = "#" & Hex$(CommonDialog1.Color)
    Exit Sub
    
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
        'picDisplay.BackColor = RGB(0, 0, 0)
    Else
        MsgBox "An error occured"
    End If

End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetData picDisplay.Image, vbCFBitmap
End Sub

Private Sub Command3_Click()
Clipboard.Clear
Clipboard.SetText Hcode.Text, vbCFText

End Sub

Private Sub Dcode_Click()
hsbRed.SetFocus
End Sub

Private Sub Form_Load()
frmIndradhanush.Top = 100

End Sub

Private Sub hsbBlue_Change()

Bs = Hex(hsbBlue.Value)
'Gs = Hex(hsbGreen.Value)
'Rs = Hex(hsbRed.Value)
txtBlue.Text = Str(hsbBlue)
'Color = Rs & Gs & Bs
'Bs = Hex(hsbBlue.Value)
'txtDisplay.Text = (Color)
picDisplay.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
Hcode.Text = "#" & Hex$(RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value))
Dcode.Text = RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value)

End Sub

Private Sub hsbGreen_Change()
Gs = Hex(hsbGreen.Value)
'Bs = Hex(hsbBlue.Value)
'Rs = Hex(hsbRed.Value)
txtGreen.Text = Str(hsbGreen.Value)

'Color = Rs & Gs & Bs
'Bs = Hex(hsbBlue.Value)
'txtDisplay.Text = (Color)
picDisplay.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
Hcode.Text = "#" & Hex$(RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value))
Dcode.Text = RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value)

End Sub

Private Sub hsbRed_Change()
Rs = Hex(hsbRed.Value)

txtRed.Text = Str(hsbRed.Value)

'Color = Rs & Gs & Bs
'Bs = Hex(hsbBlue.Value)
'txtDisplay.Text = (Color)
picDisplay.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
Color = Hex$(picDisplay.BackColor)
'Color = Color Mod 10
Hcode.Text = "#" & Hex$(RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value))
Dcode.Text = RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value)

End Sub



