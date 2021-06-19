VERSION 5.00
Begin VB.Form frmFool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FOOL"
   ClientHeight    =   3090
   ClientLeft      =   2835
   ClientTop       =   3630
   ClientWidth     =   4680
   Icon            =   "frmFool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdAccept 
      Caption         =   "I WILL TRY"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtAccept 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmFool.frx":0CCA
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ALL RIGHT!!"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtIam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton cmdNo2 
      Caption         =   "NO"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNo1 
      Caption         =   "NO"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "YES"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblRu 
      BackStyle       =   0  'Transparent
      Caption         =   "Are You FOOL ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmFool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mount As Integer

Private Sub cmdAccept_Click()
End
End Sub

Private Sub cmdNo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNo2.Visible = True
cmdNo1.Visible = False
frmFool.BackColor = vbMagenta
Mount = Mount + 1
End Sub

Private Sub cmdNo2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNo1.Visible = True
cmdNo2.Visible = False
frmFool.BackColor = vbBlue
Mount = Mount + 1
End Sub

Private Sub cmdOk_Click()
cmdOk.Visible = False
txtIam.Visible = False
txtAccept.Visible = True
cmdAccept.Visible = True
End Sub

Private Sub cmdYes_Click()
cmdNo1.Visible = False
cmdYes.Visible = False
cmdNo2.Visible = False
lblRu.Visible = False
txtIam.Visible = True
txtIam.Text = "YOU HAVE TRIED " & Str(Mount) & " TIMES TO RUN AWAY FROM REALITY"
cmdOk.Visible = True
End Sub

Private Sub Form_Load()
frmFool.BackColor = vbGreen
Mount = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmStart.Show
End Sub
