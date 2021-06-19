VERSION 5.00
Begin VB.Form frmDataEnv 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   5760
      TabIndex        =   14
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtopt4 
      DataField       =   "opt4"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtopt3 
      DataField       =   "opt3"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtopt2 
      DataField       =   "opt2"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtopt1 
      DataField       =   "opt1"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1860
      Width           =   2055
   End
   Begin VB.TextBox txtquest 
      DataField       =   "quest"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   525
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
   End
   Begin VB.TextBox txtID 
      DataField       =   "ID"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1095
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "opt4:"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "opt3:"
      Height          =   255
      Index           =   4
      Left            =   990
      TabIndex        =   7
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "opt2:"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "opt1:"
      Height          =   255
      Index           =   2
      Left            =   990
      TabIndex        =   3
      Top             =   1905
      Width           =   375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "quest:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    DataEnvironment1.rsCommand1.MoveFirst
End Sub

Private Sub Command2_Click()
    If DataEnvironment1.rsCommand1.BOF Then
        Beep
    Else
        DataEnvironment1.rsCommand1.MovePrevious
        If DataEnvironment1.rsCommand1.BOF Then
            DataEnvironment1.rsCommand1.MoveFirst
        End If
    End If
End Sub

Private Sub Command3_Click()
    If DataEnvironment1.rsCommand1.EOF Then
        Beep
    Else
        DataEnvironment1.rsCommand1.MoveNext
        If DataEnvironment1.rsCommand1.EOF Then
            DataEnvironment1.rsCommand1.MoveLast
        End If
    End If
End Sub

Private Sub Command4_Click()
    DataEnvironment1.rsCommand1.MoveLast
End Sub


