VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H008080FF&
      Height          =   1695
      Left            =   1800
      ScaleHeight     =   1635
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Hcode 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Dcode 
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lal As Integer, Hara As Integer, Nila As Integer
Dim Rang As Long


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
    Rang& = CommonDialog1.Color
    Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    Text1.Text = (Hex(Lal))
    Text2.Text = (Hex(Hara))
    Text3.Text = (Hex(Nila))
    Hcode.Text = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)
    Exit Sub
    
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
        'picDisplay.BackColor = RGB(0, 0, 0)
    Else
        MsgBox "An error occured"
    End If


End Sub
