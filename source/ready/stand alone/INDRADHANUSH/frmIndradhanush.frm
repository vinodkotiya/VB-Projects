VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIndradhanush 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDRA-DHANUSH"
   ClientHeight    =   7095
   ClientLeft      =   3360
   ClientTop       =   -1200
   ClientWidth     =   6450
   Icon            =   "frmIndradhanush.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmIndradhanush.frx":1CCA
   ScaleHeight     =   7095
   ScaleWidth      =   6450
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "color code"
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   1095
      Begin VB.OptionButton format 
         BackColor       =   &H000000FF&
         Caption         =   "VB"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   470
         Width           =   855
      End
      Begin VB.OptionButton format 
         BackColor       =   &H000000FF&
         Caption         =   "HTML"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   230
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Display users hexadecimal code"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6600
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2595
      Left            =   120
      MouseIcon       =   "frmIndradhanush.frx":210C
      MousePointer    =   99  'Custom
      Picture         =   "frmIndradhanush.frx":254E
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   218
      TabIndex        =   26
      Top             =   1560
      Width           =   3330
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Open"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   29
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Clear"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   28
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3675
      Left            =   0
      Picture         =   "frmIndradhanush.frx":33AB
      ScaleHeight     =   3615
      ScaleWidth      =   3510
      TabIndex        =   25
      Top             =   720
      Width           =   3570
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Height          =   2280
      Left            =   0
      ScaleHeight     =   148
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0080FF80&
      Caption         =   "Auto copy Color .BMP"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   $"frmIndradhanush.frx":5075
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Auto copy Hex Code"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      ToolTipText     =   $"frmIndradhanush.frx":510B
      Top             =   6600
      WhatsThisHelpID =   1
      Width           =   1095
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "You can enter the value (0-255) of red color component here."
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtPerr 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "0.00 %"
      ToolTipText     =   "This is the percentage of red component in the color currently displaying."
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtPerb 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "0.00 %"
      ToolTipText     =   "This is the percentage of blue component in the color currently displaying."
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtPerg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "0.00 %"
      ToolTipText     =   "This is the percentage of green component in the color currently displaying."
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Random Color Selector"
      Height          =   1215
      Left            =   5160
      Picture         =   "frmIndradhanush.frx":51BE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "This will display the color randomly."
      Top             =   3120
      Width           =   1215
   End
   Begin VB.HScrollBar hsbRed 
      Height          =   375
      Left            =   840
      Max             =   255
      TabIndex        =   1
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Blue"
      Top             =   5745
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Green"
      Top             =   5145
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Red"
      Top             =   4545
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "frmIndradhanush.frx":6E88
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5040
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "COLOR NO."
      Top             =   1290
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "COPY CODE"
      Height          =   735
      Left            =   5160
      Picture         =   "frmIndradhanush.frx":6E9B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopycolor 
      BackColor       =   &H00FFFFC0&
      Caption         =   "COPY COLOR"
      Height          =   735
      Left            =   3840
      Picture         =   "frmIndradhanush.frx":71A5
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Hcode 
      BackColor       =   &H00FFFF00&
      Height          =   495
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   4
      ToolTipText     =   "Currently displaying the hexadecimal code of the color which can be used in your web page."
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Dcode 
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   3840
      MaxLength       =   9
      TabIndex        =   15
      ToolTipText     =   "This is the color number of the color displaying or decimal equivalent of the hex code."
      Top             =   1290
      Width           =   1215
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
      Height          =   1215
      Left            =   3840
      Picture         =   "frmIndradhanush.frx":74AF
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "This is the custom color selection of Microsoft Windows"
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1215
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
      MaxLength       =   4
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "You can enter the value (0-255) of blue color component here."
      Top             =   5760
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "You can enter the value (0-255) of green color component here."
      Top             =   5160
      Width           =   975
   End
   Begin VB.HScrollBar hsbBlue 
      Height          =   375
      Left            =   840
      Max             =   255
      TabIndex        =   3
      Top             =   5760
      Width           =   3135
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   375
      Left            =   840
      Max             =   255
      TabIndex        =   2
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Set on top of All Windows"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      MaskColor       =   &H000000FF&
      TabIndex        =   12
      ToolTipText     =   "This will set the Indradhanush window on the top of all other windows currently opened."
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   4080
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "select 1  from 16777215 colors"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   33
      Top             =   1000
      Width           =   2415
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   945
      Left            =   1680
      OleObjectBlob   =   "frmIndradhanush.frx":8379
      SizeMode        =   1  'Stretch
      SourceDoc       =   "F:\HTML\credit\color.html"
      TabIndex        =   23
      Top             =   -120
      Width           =   1020
   End
   Begin VB.Image imgCredit2 
      Height          =   330
      Left            =   2760
      Picture         =   "frmIndradhanush.frx":E191
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgCredit1 
      Height          =   330
      Left            =   2760
      Picture         =   "frmIndradhanush.frx":E672
      Top             =   120
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   3600
      MouseIcon       =   "frmIndradhanush.frx":EBFF
      MousePointer    =   99  'Custom
      Picture         =   "frmIndradhanush.frx":108C9
      ToolTipText     =   $"frmIndradhanush.frx":18C93
      Top             =   50
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      MouseIcon       =   "frmIndradhanush.frx":18D1F
      MousePointer    =   99  'Custom
      Picture         =   "frmIndradhanush.frx":1A9E9
      ToolTipText     =   "INDRADHANUSH (The Color Picker) is created by Vinod Kotiya.Copywrite August2002 . All rights unreserved."
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
Dim Autocopycode As Boolean, Autocopycolor As Boolean
Dim Per As String * 6
Dim Rang As Long
Dim Lal As Integer, Hara As Integer, Nila As Integer
'Dim Rs As String, Gs As String, Bs As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_SHOWWINDOW = &H40
Dim prevx As Single, prevy As Single



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

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
  Check3.Value = vbUnchecked
 Autocopycode = True
 Else
 Autocopycode = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = vbChecked Then
  Check2.Value = vbUnchecked
 Autocopycolor = True
 Else
  Autocopycolor = False
End If

End Sub

Private Sub cmdCopycolor_Click()
Clipboard.Clear
Clipboard.SetData picDisplay.Image, vbCFBitmap
End Sub

Private Sub Command1_Click()
Picture2.Cls
Dim CDFlags As Long
On Error GoTo ColorError

    'CDFlags = 0
    'For i = 0 To 3
    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)
    'Next
    CmDialog1.Flags = CDFlags
    CmDialog1.Color = picDisplay.BackColor
    CmDialog1.CancelError = True
    CmDialog1.ShowColor
    picDisplay.BackColor = CmDialog1.Color
    Rang& = CmDialog1.Color
    Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    hsbRed.Value = Lal
    hsbGreen.Value = Hara
    hsbBlue.Value = Nila
    'Dcode.Text = CmDialog1.Color
    'Hcode.Text = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)
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
Dim txtCol As String
On Error GoTo colerror
If format(0).Value = True Then
txtCol = "Please Enter the hexadecimal color code in HTML format RGB " & vbCrLf & _
 "eg. ff2134 " & vbCrLf & "Here Red is " & Chr(34) & "ff" & Chr(34) & "  Green is " & Chr(34) & "21" & Chr(34) & " and Blue is " & Chr(34) & "34" & Chr(34) & _
"without double quotes."
ElseIf format(1).Value = True Then
txtCol = "Please Enter the hexadecimal color code in VB format BGR " & vbCrLf & _
 "eg. ff2134 " & vbCrLf & "Here Blue is " & Chr(34) & "ff" & Chr(34) & "  Green is " & Chr(34) & "21" & Chr(34) & " and Red is " & Chr(34) & "34" & Chr(34) & vbCrLf & _
"without double quotes."
End If
txtCol = InputBox(txtCol, "Enter Hexadecimal Color Code")
If Trim(txtCol) = "" Then Exit Sub
txtCol = Trim(txtCol)
Dim txtLal, txtHara, txtNila As String
If format(0).Value = True Then         '///RGB format for HTML
  txtLal = Left(txtCol, 2)
  txtHara = Mid(txtCol, 3, 2)
  txtNila = Right(txtCol, 2)
  picDisplay.BackColor = "&h" & txtLal & txtHara & txtNila
  Hcode.Text = Chr(34) & "#" & txtCol & Chr(34)
ElseIf format(1).Value = True Then          '///BGR    for vb
 txtLal = Right(txtCol, 2)
 txtHara = Mid(txtCol, 3, 2)
 txtNila = Left(txtCol, 2)
 picDisplay.BackColor = "&h" & txtNila & txtHara & txtLal
 Hcode.Text = Chr(34) & "&H" & txtCol & Chr(34)
End If
 Exit Sub
colerror:
 MsgBox "Invalid Color Code"
End Sub

Private Sub Command3_Click()
Clipboard.Clear
Clipboard.SetText Hcode.Text, vbCFText

End Sub

Private Sub Command4_Click()
Picture2.Cls
Dim Rand As Single
Hara = 255
Randomize
Rand = Rnd(1)
Lal = Rand * 1000
If Lal > 255 Then
 Lal = Lal / 4
End If
hsbRed.Value = Lal
Rand = Rnd(1)
Hara = Rand * 1000
If Hara > 255 Then
 Hara = Hara / 4
End If
hsbGreen.Value = Hara
Rand = Rnd(1)
Nila = Rand * 1000
If Nila > 255 Then
 Nila = Nila / 4
End If
hsbBlue.Value = Nila


End Sub





Private Sub Dcode_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim n As String

If KeyCode = 13 Or KeyCode = 40 Then
 
'not working page353
 'n$ = Hex(Val(Dcode.Text))
 'hsbRed.Value = Val("&H" & Right(n$, 2))
 'hsbGreen.Value = Val("&H" & Mid$(n$, 3, 2))
 'hsbBlue.Value = Val("&H" & Mid$(n$, 5, 2))
 
 Rang& = Val(Dcode.Text)
 hsbRed.Value = Rang& Mod 256
 hsbGreen.Value = ((Rang& And &HFF00FF00) / 256&)
 hsbBlue.Value = ((Rang& And &HFF0000) / 65536)
 hsbRed.SetFocus
End If
End Sub

Private Sub Form_Load()
frmIndradhanush.Top = 100
'Picture2.PaintPicture Picture1.Picture, 0, 0,
      '  Picture2.Width , Picture2.Height, _
      '  0, 0, _
      '  Picture2.Width, Picture2.Height
End Sub

Private Sub gg_Click()
Label1_Click (1)
End Sub

Private Sub Hcode_KeyUp(KeyCode As Integer, Shift As Integer)


'Dim storehex As Variant
'If KeyCode = 13 Then
' storehex = CVar(Hcode.Text)
' Dcode.Text = (storehex)
'Rang = storehex Mod 10

 ''txtRed.Text = Hex(15)
 
' '   Lal = Val(Mid$(Hcode.Text, 2, 2))
' '   Dcode.Text = Str(Lal)
' '   Hara = ((Rang& And &HFF00FF00) / 256&)
'    Nila = (Rang& And &HFF00000) / 65536
'    'hsbRed.Value = Lal
'    'hsbGreen.Value = Hara
    'hsbBlue.Value = Nila
    'Dcode.Text = Str(Rang&)
 
' txtRed.SetFocus
''End If
End Sub

Private Sub hh_Click()
Label1_Click (0)
End Sub

Private Sub hsbBlue_Change()

'Bs = Hex(hsbBlue.Value)
'Gs = Hex(hsbGreen.Value)
'Rs = Hex(hsbRed.Value)
txtBlue.Text = Str(hsbBlue.Value)
scrollsetcolor
End Sub

Private Sub hsbGreen_Change()
'Gs = Hex(hsbGreen.Value)
'Bs = Hex(hsbBlue.Value)
'Rs = Hex(hsbRed.Value)
txtGreen.Text = Str(hsbGreen.Value)
scrollsetcolor

End Sub

Private Sub hsbRed_Change()
'Rs = Hex(hsbRed.Value)

txtRed.Text = Str(hsbRed.Value)
scrollsetcolor
End Sub







Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.Text = " COLOR No."
Text2.Text = " HEXADECIMAL CODE"
imgCredit1.Visible = True
imgCredit2.Visible = False
End Sub

Private Sub imgCredit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgCredit2.Visible = True
imgCredit1.Visible = False
End Sub

Private Sub imgCredit2_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell("credit.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'CREDIT.EXE' is not found in its " _
  & "Default directory APPL\CREDIT.exe "

Exit Sub
End Sub

Private Sub Label1_Click(index As Integer)
On Error GoTo Error1
 If index = 0 Then
   Picture2.Cls
 ElseIf index = 1 Then
         'Action = 1
    CmDialog1.InitDir = "D:\poster\pictures\" 'App.Path
    CmDialog1.Filter = "Imagefiles | *.BMP;*.jpg;*.gif;*.dib;*.ico"
    CmDialog1.ShowOpen
    If CmDialog1.FileName = "" Then Exit Sub
      Picture1.Picture = LoadPicture(CmDialog1.FileName)
'    frmThumb.Refresh
    Picture2.Cls
   Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
'   SliderDestWidth.Max = 400 'Picture1.ScaleWidth
 '  SliderDestHeight.Max = 300 'Picture1.ScaleHeight
   'Picture2.PaintPicture Picture1.Picture, 0, 0, _
        Picture2.Width, Picture2.Height, _
        0, 0, _
        Picture2.Width, Picture2.Height
   MsgBox "If image is larger then you can scroll it by dragging it." & Chr(13) & "To do so right click on the image and drag it."
End If
   Exit Sub
Error1:
    MsgBox "Couldn't open file " + CmDialog1.FileName
End Sub



Private Sub picDisplay_Click()
'cmdCopycolor.SetFocus

End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture2.MousePointer = 99 'Val(Text1.Text)
If x > Picture2.ScaleTop And y > Picture2.ScaleLeft Then 'And X < Picture2.ScaleTop + Picture2.ScaleHeight And Y < Picture2.ScaleLeft + Picture2.ScaleWidth Then
If Button = 1 Then
On Error GoTo CHUPCHAP
Picture2.MousePointer = 5
Picture2.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, _
        prevx + x, prevx + y, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, &HCC0020

End If
End If

Dim n As Long
n = (Picture2.Point(x, y))     'getting color code in integer
If n < 0 Then Exit Sub         'extract integer values of component
  Lal = n& Mod 256
  Hara = ((n& And &HFF00FF00) / 256&) ' Mod 256&
  Nila = (n& And &HFF0000) / 65536
  'Hcode.Text = "#" & Hex(n) '(Hex(Picture2.Point(X, Y))) & " rED " & Lal & "GREEN " & Hara & " B " & Nila
  hsbRed.Value = Lal                              'assign values to scroll bar they convert it into color code
    hsbGreen.Value = Hara
    hsbBlue.Value = Nila
 Exit Sub
CHUPCHAP:
    
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
prevx = x
prevy = y
End Sub

Private Sub Text1_Click()
Dcode.SetFocus
End Sub



Private Sub Text2_Click()
Hcode.SetFocus
End Sub







Private Sub Text3_Click()
hsbRed.SetFocus
End Sub

Private Sub Text4_Click()
hsbGreen.SetFocus
End Sub

Private Sub Text5_Click()
hsbBlue.SetFocus
End Sub



Private Sub txtBlue_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
 If Val(txtBlue.Text) > 255 Then
  MsgBox txtBlue.Text & " is not a valid blue color value.The color value must lie between 0 to 255"
  txtBlue.Text = "0"
 Else
  hsbBlue.Value = Val(txtBlue.Text) 'Nila
End If

 Dcode.SetFocus
ElseIf KeyCode = 38 Then
 If Val(txtBlue.Text) > 255 Then
  MsgBox txtBlue.Text & " is not a valid blue color value.The color value must lie between 0 to 255"
  txtBlue.Text = "0"
Else
  hsbBlue.Value = Val(txtBlue.Text) 'Nila
End If

 txtGreen.SetFocus
End If

End Sub

Private Sub txtGreen_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
 
 If Val(txtGreen.Text) > 255 Then
  MsgBox txtGreen.Text & " is not a valid green color value.The color value must lie between 0 to 255"
  txtGreen.Text = "0"
  txtGreen.SetFocus
 Else
  hsbGreen.Value = Val(txtGreen.Text) 'Hara
 End If
txtBlue.SetFocus

ElseIf KeyCode = 38 Then
 
 If Val(txtGreen.Text) > 255 Then
  MsgBox txtGreen.Text & " is not a valid green color value.The color value must lie between 0 to 255"
  txtGreen.Text = "0"
 Else
   hsbGreen.Value = Val(txtGreen.Text) 'Hara
End If
txtRed.SetFocus
End If

End Sub

Private Sub txtRed_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 40 Then
 If Val(txtRed.Text) > 255 Then
  MsgBox txtRed.Text & " is not a valid red color value.The color value must lie between 0 to 255"
  txtRed.Text = "0"
  txtRed.SetFocus
 Else
  'Lali = txtRed.Text
  'Lal = Val(Lali)
  hsbRed.Value = Val(txtRed.Text) 'Lal
 End If

 txtGreen.SetFocus
ElseIf KeyCode = 38 Then
 If Val(txtRed.Text) > 255 Then
  MsgBox txtRed.Text & " is not a valid red color value.The color value must lie between 0 to 255"
  txtRed.Text = "0"
  txtRed.SetFocus
 Else
  hsbRed.Value = Val(txtRed.Text) 'Lal
  txtGreen.SetFocus
 End If
 
 Dcode.SetFocus
 
End If
End Sub

Private Sub scrollsetcolor()
'Color = Rs & Gs & Bs
'Bs = Hex(hsbBlue.Value)
'txtDisplay.Text = (Color)
picDisplay.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
'Hcode.Text = "#" & Hex(picDisplay.BackColor) 'hsbRed.Value) & Hex(hsbGreen.Value) & Hex(hsbBlue.Value) 'Hex$(RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value))
'''/////////// convert in to desired format
If format(0).Value = True Then
    Hcode.Text = Chr(34) & "#" & Hex(hsbRed.Value) & Hex(hsbGreen.Value) & Hex(hsbBlue.Value) & Chr(34) 'Hex$(RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value))
ElseIf format(1).Value = True Then
   Hcode.Text = Chr(34) & "&H" & Hex(hsbBlue.Value) & Hex(hsbGreen.Value) & Hex(hsbRed.Value) & Chr(34) 'Hex$(RGB(hsbBlue.Value, hsbGreen.Value, hsbRed.Value))
End If
Dcode.Text = picDisplay.BackColor

'///////////// show percentage contribution //////////////
If hsbRed.Value > 0 Then
 Per = Str((hsbRed.Value * 100) / (hsbRed.Value + hsbGreen.Value + hsbBlue.Value))
 txtPerr.Text = Per & " %"
End If
If hsbBlue.Value > 0 Then
 Per = Str((hsbBlue.Value * 100) / (hsbRed.Value + hsbGreen.Value + hsbBlue.Value))
 txtPerb.Text = Per & " %"
End If
If hsbGreen.Value > 0 Then
 Per = Str((hsbGreen.Value * 100) / (hsbRed.Value + hsbGreen.Value + hsbBlue.Value))
 txtPerg.Text = Per & " %"
End If
'autocopy the hex  code
If Autocopycode = True Then
 Clipboard.Clear
 Clipboard.SetText Hcode.Text, vbCFText
End If
'autocopy the color
If Autocopycolor = True Then
Clipboard.Clear
Clipboard.SetData picDisplay.Image, vbCFBitmap
End If

End Sub
