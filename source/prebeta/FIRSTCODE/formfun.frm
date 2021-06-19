VERSION 5.00
Begin VB.Form frmFormFun 
   Caption         =   "FORMFUN"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVin 
      Caption         =   "VINOD"
      Height          =   2535
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMagenta 
      Caption         =   "MAGENTA FORM"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlue 
      Caption         =   "BLUE FORM"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "SHOW BUTTONS"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "HIDE BUTTONS"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrow 
      Caption         =   "GROW   FORM"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdShrink 
      Caption         =   "SHRINK FORM"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmFormFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlue_Click()
frmFormFun.BackColor = vbBlue
End Sub

Private Sub cmdGrow_Click()
frmFormFun.Height = frmFormFun.Height + 100
frmFormFun.Width = frmFormFun.Width + 100
End Sub

Private Sub cmdHide_Click()
cmdShrink.Visible = False
cmdGrow.Visible = False
cmdHide.Visible = False
cmdBlue.Visible = False
cmdMagenta.Visible = False
cmdShow.Visible = True
End Sub

Private Sub cmdHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdVin.Visible = True
End Sub

Private Sub cmdMagenta_Click()
frmFormFun.BackColor = vbMagenta
End Sub

Private Sub cmdShow_Click()
cmdShrink.Visible = True
cmdGrow.Visible = True
cmdHide.Visible = True
cmdBlue.Visible = True
cmdMagenta.Visible = True
cmdShow.Visible = False
End Sub

Private Sub cmdShrink_Click()
frmFormFun.Height = frmFormFun.Height - 100
frmFormFun.Width = frmFormFun.Width - 100
End Sub

Private Sub cmdVin_Click()
cmdVin.Visible = False
End Sub
