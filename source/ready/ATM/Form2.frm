VERSION 5.00
Begin VB.Form WITHDRAWAL 
   BackColor       =   &H00FF8080&
   Caption         =   "WITHDRAWAL"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_CONT 
      BackColor       =   &H00FF8080&
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
      BackColor       =   &H00FF8080&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0;(#,##0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   6855
   End
   Begin VB.Label LBL_WITHDRAWAL 
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H80000008&
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
Private Sub Form_Load()

End Sub
