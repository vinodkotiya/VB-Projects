VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LOGINSCREEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGINSCREEN"
   ClientHeight    =   6915
   ClientLeft      =   2340
   ClientTop       =   1275
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LOGINSCREEN.frx":0000
   ScaleHeight     =   6915
   ScaleWidth      =   7875
   Begin VB.TextBox Text2 
      DataField       =   "PASSWORD"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "ACCOUNT_NO"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   1080
      Top             =   5880
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=TIGER;User ID=SCOTT;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=TIGER;User ID=SCOTT;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM ACCOUNT"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
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
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox TXT_PWD 
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
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox TXT_ACCOUNTNO 
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
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label LBL_PWD 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label LBL_ACCOUNTNO 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT NO."
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
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label LBL_ENTRY 
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE ENTER YOUR"
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
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "LOGINSCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMD_CONT_Click()
Dim I As Integer
I = 3

ACC = TXT_ACCOUNTNO
PASS = TXT_PWD
Adodc1.Recordset.Find "[ACCOUNT_NO]= " & ACC

If Adodc1.Recordset.EOF Then
MsgBox "INVALID ID "
Unload Me
LOGINSCREEN.Show
Exit Sub
End If

If Adodc1.Recordset(2) <> PASS Then
MsgBox "INVALID PASSWORD"
End If

If Adodc1.Recordset(2) = PASS Then
NM = Adodc1.Recordset(1)
BAL = Adodc1.Recordset(5)
PASS = Adodc1.Recordset(2)
CT = Adodc1.Recordset(4)
ADD = Adodc1.Recordset(3)
Unload Me
MENUSCREEN.Show

End If

End Sub

