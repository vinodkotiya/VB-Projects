VERSION 5.00
Begin VB.Form CANCEL_PWD 
   BackColor       =   &H80000012&
   Caption         =   "CANCEL_PWD"
   ClientHeight    =   6915
   ClientLeft      =   2415
   ClientTop       =   1290
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   Picture         =   "CANCEL_PWD.frx":0000
   ScaleHeight     =   6915
   ScaleWidth      =   7890
   Begin VB.CommandButton CMD_RETURN 
      BackColor       =   &H80000018&
      Caption         =   "RETURN TO MAIN MENU"
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
      TabIndex        =   1
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label LBL_CANCELPASSWORD 
      BackStyle       =   0  'Transparent
      Caption         =   "    NEW PASSWORD  COULD NOT BE                CONFIRMED                                                    PLEASE TRY AGAIN....."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7575
   End
End
Attribute VB_Name = "CANCEL_PWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
