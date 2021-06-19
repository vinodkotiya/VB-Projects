VERSION 5.00
Begin VB.Form frmSaving 
   Caption         =   "SAVING ACCOUNT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDeposit 
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtMonths 
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "COMPUTE SAVINGS"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Caption         =   "TOTAL SAVINGS"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblTotaling 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblDeposit 
      Alignment       =   2  'Center
      Caption         =   "MONTHLY DEPOSIT"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblMonths 
      Alignment       =   2  'Center
      Caption         =   "NO.OF MONTHS"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "frmSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Deposit As Integer
Dim Months As Integer
Dim Total As Integer


Private Sub cmdCompute_Click()
Deposit = Val(txtDeposit.Text)
Months = Val(txtMonths.Text)
Total = Deposit * Months
lblTotaling.Caption = Str(Total)
End Sub


Private Sub cmdExit_Click()
End
End Sub

