VERSION 5.00
Begin VB.Form DEPOSIT 
   Caption         =   "DEPOSIT"
   ClientHeight    =   6960
   ClientLeft      =   2610
   ClientTop       =   1230
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   Picture         =   "DEPOSIT.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   7620
   Begin VB.CommandButton CMD_CONT 
      BackColor       =   &H80000018&
      Caption         =   "CONTINUE"
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
      TabIndex        =   2
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox TXT_DEPOSIT 
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0;(#,##0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   6495
   End
   Begin VB.Label LBL_DEPOSIT 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER THE AMOUNT YOU WANT TO DEPOSIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   6495
   End
End
Attribute VB_Name = "DEPOSIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_CONT_Click()
DEPO = TXT_DEPOSIT
BAL = BAL + DEPO
W = TXT_DEPOSIT
TTYPE = "DEPOSIT"
Unload Me
UPDATE.Show
End Sub
