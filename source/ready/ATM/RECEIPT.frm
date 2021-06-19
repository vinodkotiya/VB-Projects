VERSION 5.00
Begin VB.Form RECEIPT 
   Caption         =   "RECEIPT"
   ClientHeight    =   6870
   ClientLeft      =   2550
   ClientTop       =   1425
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   Picture         =   "RECEIPT.frx":0000
   ScaleHeight     =   6870
   ScaleWidth      =   7635
   Begin VB.CommandButton CMD_RETURN 
      BackColor       =   &H80000018&
      Caption         =   "RETURN TO MAIN MENU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label LBL_RECEIPT 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE COLLECT YOUR RECEIPT.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   7215
   End
End
Attribute VB_Name = "RECEIPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_RETURN_Click()
Unload Me
MENUSCREEN.Show
End Sub
