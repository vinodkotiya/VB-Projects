VERSION 5.00
Begin VB.Form frmSci 
   BorderStyle     =   0  'None
   Caption         =   "SCIENTIFIC"
   ClientHeight    =   3915
   ClientLeft      =   5865
   ClientTop       =   3495
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSci.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdHex 
      Caption         =   "hex "
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "cos"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdTan 
      Caption         =   "tan"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdLogn 
      Caption         =   "ln"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "exp"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "log"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdFact 
      Caption         =   "n !"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdOct 
      Caption         =   "oct"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdSin 
      Caption         =   "sin"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   495
   End
End
Attribute VB_Name = "frmSci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCos_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Cos(Val(frmCal.txtLcd.Text) * 1.74532925199433E-02))
End If
End Sub

Private Sub cmdLog_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Log(Val(frmCal.txtLcd.Text)) / 2.30258509299405)
End If

End Sub

Private Sub cmdLogn_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Log(Val(frmCal.txtLcd.Text)))
End If

End Sub

Private Sub cmdSin_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Sin(Val(frmCal.txtLcd.Text) * 1.74532925199433E-02))
'Form1.Ref = 1
End If
End Sub

Private Sub cmdTan_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Tan(Val(frmCal.txtLcd.Text) * 1.74532925199433E-02))
End If

End Sub

