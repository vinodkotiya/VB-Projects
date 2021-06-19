VERSION 5.00
Begin VB.Form CHANGE_PWD 
   Caption         =   "CHANGE_PWD"
   ClientHeight    =   6960
   ClientLeft      =   2550
   ClientTop       =   1365
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   Picture         =   "CHANGE_PWD.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   7605
   Begin VB.CommandButton CMD_CON 
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
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox TXT_CONFIRMPWD 
      BackColor       =   &H80000018&
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
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox TXT_NEWPWD 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox TXT_OLDPWD 
      BackColor       =   &H80000018&
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
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label LBL_CONFIRMPWD 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM     PASSWORD"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label LBL_NEWPWD 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NEW PASSWORD"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label LBL_OLDPWD 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER OLD PASSWORD"
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "CHANGE_PWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub CMD_CON_Click()
If TXT_NEWPWD.Text <> TXT_CONFIRMPWD Then
MsgBox "ENTER NEW PASSWORD AGAIN"
TXT_NEWPWD.Text = ""
TXT_CONFIRMPWD.Text = ""
Else
LOGINSCREEN.Adodc1.Recordset(2) = TXT_NEWPWD.Text
LOGINSCREEN.Adodc1.Recordset.UPDATE
Unload Me
CONFIRM_PWD.Show
End If
End Sub
