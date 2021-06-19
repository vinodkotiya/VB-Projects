VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help On Fone Directory"
   ClientHeight    =   8295
   ClientLeft      =   1935
   ClientTop       =   1035
   ClientWidth     =   7440
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   7440
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Start Searching"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Line Line 
      BorderColor     =   &H000000FF&
      Index           =   5
      X1              =   3360
      X2              =   4800
      Y1              =   4320
      Y2              =   4800
   End
   Begin VB.Line Line 
      BorderColor     =   &H000000FF&
      Index           =   4
      X1              =   1800
      X2              =   5160
      Y1              =   1920
      Y2              =   1680
   End
   Begin VB.Line Line 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   2880
      X2              =   5400
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      Index           =   2
      X1              =   120
      X2              =   7200
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   7200
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Labail 
      BackStyle       =   0  'Transparent
      Caption         =   "How To Save The Search Results"
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
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   3855
   End
   Begin VB.Label Labail 
      BackStyle       =   0  'Transparent
      Caption         =   "How To Modify The Directory"
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Labail 
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Search"
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "1 : To Change any record Click the button ""Modify It""."
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   12
      Top             =   6240
      Width           =   5655
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "To Modify the directory Click the button ""Modify"".A new window will appear."
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   5295
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0442
      ForeColor       =   &H0000FF00&
      Height          =   855
      Index           =   4
      Left            =   360
      TabIndex        =   10
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":04E6
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   5295
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "3 : Press the button at right ""Start Searching"""
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "2 : Now Enter the Item to be searched in the textbox Labelled 'To Search' "
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "2 : To save the results as Text file,Vin file or any other format choose from Menu ""Save As Text"""
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   7800
      Width           =   6015
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "1 : To save the results as WebPage choose from Menu ""Save As Web Page"""
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   7560
      Width           =   6015
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      Index           =   0
      X1              =   120
      X2              =   7200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   3720
      Picture         =   "frmHelp.frx":057B
      Top             =   960
      Width           =   3540
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   3480
      Picture         =   "frmHelp.frx":14DA
      Top             =   4080
      Width           =   3510
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "3 : To delete any record first display it Then Press ""Delete"" Button"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   6720
      Width           =   4935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "2 : To add a new record first click on button ""Add New"" and Then Enter data"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   6480
      Width           =   6015
   End
   Begin VB.Label Labail 
      BackStyle       =   0  'Transparent
      Caption         =   "How To Search"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":25BF
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

