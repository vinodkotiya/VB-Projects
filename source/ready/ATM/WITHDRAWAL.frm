VERSION 5.00
Begin VB.Form WITHDRAWAL 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WITHDRAWAL"
   ClientHeight    =   7140
   ClientLeft      =   2220
   ClientTop       =   1110
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "WITHDRAWAL.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMD_CONT 
      BackColor       =   &H80000018&
      Caption         =   "CONTINUE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
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
   Begin VB.TextBox TXT_WITHDRAWAL 
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
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   6735
   End
   Begin VB.Label LBL_WITHDRAWAL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER THE AMOUNT YOU WANT TO WITHDRAW"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   6855
   End
End
Attribute VB_Name = "WITHDRAWAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_CONT_Click()
DT = 2004
W = (TXT_WITHDRAWAL)

If W > BAL Then
Unload Me
OVERDRAFT.Show
Else
BAL = BAL - W
TTYPE = "WITHDRAWAL"
Unload Me
UPDATE.Show
End If

End Sub
