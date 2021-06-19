VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN CONVERT CENTRE"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmConvert.frx":1CCA
   ScaleHeight     =   6330
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add New Unit/Delete Existing Unit"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5760
      Width           =   5775
   End
   Begin VB.CommandButton cmdMoreother 
      BackColor       =   &H00FFFF80&
      Caption         =   "More......"
      Height          =   255
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdMore 
      BackColor       =   &H00FFFF80&
      Caption         =   "More......."
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox result 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Index           =   5
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   45
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3840
      Width           =   3135
   End
   Begin VB.ComboBox cmbFrom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   5
      ItemData        =   "frmConvert.frx":BCA2
      Left            =   8040
      List            =   "frmConvert.frx":BCA4
      TabIndex        =   43
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox cmbTo 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   5
      Left            =   8760
      TabIndex        =   42
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   5
      Left            =   7680
      TabIndex        =   41
      Text            =   "1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox result 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Index           =   4
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3840
      Width           =   3135
   End
   Begin VB.ComboBox cmbFrom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   4
      Left            =   4560
      TabIndex        =   36
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox cmbTo 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   4
      Left            =   5280
      TabIndex        =   35
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   4
      Left            =   4200
      TabIndex        =   34
      Text            =   "1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox result 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Index           =   3
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3840
      Width           =   3135
   End
   Begin VB.ComboBox cmbFrom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   1080
      TabIndex        =   29
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox cmbTo 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   1800
      TabIndex        =   28
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   720
      TabIndex        =   27
      Text            =   "1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox result 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Index           =   1
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ComboBox cmbFrom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   8040
      TabIndex        =   22
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox cmbTo 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   8760
      TabIndex        =   21
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   7680
      TabIndex        =   20
      Text            =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   4200
      TabIndex        =   17
      Text            =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbTo 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   5280
      TabIndex        =   16
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox cmbFrom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   4560
      TabIndex        =   15
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox result 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Index           =   0
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox result 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Index           =   2
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ComboBox cmbFrom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox cmbTo 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Text            =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbOthers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      Top             =   5040
      Width           =   4455
   End
   Begin VB.ComboBox cmbFactors 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   51
      Top             =   2475
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   50
      Top             =   2470
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   49
      Top             =   2475
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   2055
      Index           =   5
      Left            =   7080
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   2055
      Index           =   4
      Left            =   3600
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   2055
      Index           =   3
      Left            =   120
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   48
      Top             =   80
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   2055
      Index           =   2
      Left            =   7080
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   47
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are In"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   46
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   40
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are In"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   39
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   33
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are In"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   32
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   26
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are In"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   25
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   19
      Top             =   75
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   2055
      Index           =   1
      Left            =   3600
      Top             =   240
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   2055
      Index           =   1
      Left            =   3600
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   18
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are In"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   80
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   2055
      Index           =   0
      Left            =   120
      Top             =   240
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   2055
      Index           =   0
      Left            =   120
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "How Many"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are In"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Factors"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   2055
      Index           =   2
      Left            =   7080
      Top             =   240
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   2055
      Index           =   3
      Left            =   120
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   2055
      Index           =   4
      Left            =   3600
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   2055
      Index           =   5
      Left            =   7080
      Top             =   2640
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'///////////////////////////////////////////////////
'///// created by vinod kotiya ////////////////////
'///// date 11-03-2003 8:30 PM TO 2:00 PM
'/////a blonder mistake is done that i have named two
'/////combo boxes cmbfrom which should be cmbto and
'///// cmbto which should be cmbfrom
'////////////////////////////////////////////////






Option Explicit
Dim converterlength(60) As Double    'holds the multiple w.r.t meter like 100,corresponding for cm
Dim unitslength(60) As String    'holds the unit
Dim convertermass(60) As Double    'holds the multiple w.r.t kg like 1000,corresponding for gm
Dim unitsmass(60) As String    'holds the unit
Dim convertertime(60) As Double    'holds the multiple w.r.t sec like 60,corresponding for min
Dim unitstime(60) As String    'holds the unit
Dim convertertemperature(10) As Double    'holds the multiple w.r.t 'C like 32,corresponding for farenhite
Dim unitstemperature(10) As String    'holds the unit
Dim converterarea(60) As Double    'holds the multiple w.r.t metersq like 1000000,corresponding for km
Dim unitsarea(60) As String    'holds the unit
Dim convertervolume(60) As Double    'holds the multiple w.r.t metercube like 1000,corresponding for litre
Dim unitsvolume(60) As String    'holds the unit

Dim isbuttoncolchange As Boolean  'restore button col when mouse move on form control processing



Private Sub cmbFrom_Click(Index As Integer)
result(Index).Visible = False
End Sub

Private Sub cmbTo_Click(Index As Integer)
result(Index).Visible = False
End Sub

Private Sub cmdMore_Click()
Dim txt As String
txt = InputBox("Enter new factor", "New Factor", "1 inch = 1440 twips")
If Trim(txt) = "1 inch = 1440 twips" Or Trim(txt) = "" Then
 Exit Sub
Else
 Dim Fsys As New Scripting.FileSystemObject
 Dim Tstream As TextStream
 Set Tstream = Fsys.OpenTextFile(App.Path & "\data\factors.vin", ForAppending)
 Tstream.WriteLine (txt)
 Dim i As Integer
  For i = cmbFactors.ListCount - 1 To 0
  cmbFactors.RemoveItem i
 Next
  loadfactors
End If
End Sub

Private Sub cmdMoreother_Click()
Dim txt As String
txt = InputBox("Enter new constants", "New Constnts (VIN Convert Centre)", "pi = 3.14")
If Trim(txt) = "pi = 3.14" Or Trim(txt) = "" Then
 'MsgBox "uy"
 Exit Sub
Else
 Dim Fsys As New Scripting.FileSystemObject
 Dim Tstream As TextStream
 Set Tstream = Fsys.OpenTextFile(App.Path & "\data\others.vin", ForAppending)
 Tstream.WriteLine (txt)
 Dim i As Integer
  For i = cmbOthers.ListCount - 1 To 0
  cmbOthers.RemoveItem i
 Next
  loadothers
End If

End Sub

Private Sub Command1_Click(whichbutton As Integer)

Dim ioffrom As Integer   'index of cmbfrom
Dim iofto As Integer     'index of cmbto
'On Error GoTo vinerror
If cmbFrom(whichbutton).ListIndex < 0 Or cmbTo(whichbutton).ListIndex < 0 Then
 MsgBox "Please specify what do you want?"
 Exit Sub
End If
result(whichbutton).Visible = True


ioffrom = cmbFrom(whichbutton).ListIndex
iofto = cmbTo(whichbutton).ListIndex
Dim temp As Double
temp = CDbl(Text1(whichbutton).Text)
If whichbutton = 0 Then
  result(whichbutton).Text = Str(temp) & " " & unitslength(iofto) & " = "
  result(whichbutton).Text = result(whichbutton).Text & Str(CDbl(Text1(whichbutton).Text) * (converterlength(ioffrom) / converterlength(iofto))) & " " & unitslength(ioffrom)
ElseIf whichbutton = 1 Then
    result(whichbutton).Text = Str(temp) & " " & unitsmass(iofto) & " = "
  result(whichbutton).Text = result(whichbutton).Text & Str(CDbl(Text1(whichbutton).Text) * (convertermass(ioffrom) / convertermass(iofto))) & " " & unitsmass(ioffrom)
ElseIf whichbutton = 2 Then
  result(whichbutton).Text = Str(temp) & " " & unitstime(iofto) & " = "
  result(whichbutton).Text = result(whichbutton).Text & Str(CDbl(Text1(whichbutton).Text) * (convertertime(ioffrom) / convertertime(iofto))) & " " & unitstime(ioffrom)
ElseIf whichbutton = 3 Then
  Dim ansh As Double
  result(whichbutton).Text = Str(temp) & " " & unitstemperature(iofto) & " = "
  If cmbFrom(whichbutton).ListIndex = 0 And cmbTo(whichbutton).ListIndex = 2 Then   'means to farenhite and  fromkelin
  ansh = (CDbl(Text1(whichbutton).Text) + 273.16)
  ansh = (ansh * 1.8) + 32
  ElseIf cmbFrom(whichbutton).ListIndex = 0 And cmbTo(whichbutton).ListIndex = 1 Then   'means to farenhite from  celsius
  ansh = ((CDbl(Text1(whichbutton).Text) * 1.8) + 32)
  ElseIf cmbFrom(whichbutton).ListIndex = 0 Then  'f to f
  ansh = CDbl(Text1(whichbutton).Text)
  ElseIf cmbFrom(whichbutton).ListIndex = 2 And cmbTo(whichbutton).ListIndex = 0 Then   'means to kelvin from fare
  ansh = ((CDbl(Text1(whichbutton).Text) - 32) / 9) * 5 'now ansh holding cel
  ansh = ansh + 273.16        'now ansh hold calvin
  ElseIf cmbFrom(whichbutton).ListIndex = 2 And cmbTo(whichbutton).ListIndex = 1 Then   'means to kelvin from cel
  ansh = CDbl(Text1(whichbutton).Text) + 273.16        'now ansh hold calvin
  ElseIf cmbFrom(whichbutton).ListIndex = 2 Then  'k to k
  ansh = CDbl(Text1(whichbutton).Text)
  ElseIf cmbFrom(whichbutton).ListIndex = 1 And cmbTo(whichbutton).ListIndex = 0 Then   'means to cels from  faren
  ansh = ((CDbl(Text1(whichbutton).Text) - 32) / 9) * 5 'now ansh holding cel
  ElseIf cmbFrom(whichbutton).ListIndex = 1 And cmbTo(whichbutton).ListIndex = 2 Then   'means to cels from  kalvin
  ansh = CDbl(Text1(whichbutton).Text) - 273.16 'now ansh holding cel
  ElseIf cmbFrom(whichbutton).ListIndex = 2 Then  'c to c
  ansh = CDbl(Text1(whichbutton).Text)
  End If
result(whichbutton).Text = result(whichbutton).Text & Str(ansh) & "  " & unitstemperature(ioffrom) 'cmbFrom(whichbutton).List(ioffrom)
 
 
ElseIf whichbutton = 4 Then
  result(whichbutton).Text = Str(temp) & " " & unitsarea(iofto) & " = "
  result(whichbutton).Text = result(whichbutton).Text & Str(CDbl(Text1(whichbutton).Text) * (converterarea(ioffrom) / converterarea(iofto))) & " " & unitsarea(ioffrom)
ElseIf whichbutton = 5 Then
  result(whichbutton).Text = Str(temp) & " " & unitsvolume(iofto) & " = "
  result(whichbutton).Text = result(whichbutton).Text & Str(CDbl(Text1(whichbutton).Text) * (convertervolume(ioffrom) / convertervolume(iofto))) & " " & unitsvolume(ioffrom)
End If

Exit Sub
'MsgBox converterlength(iofto) & "    " & converterlength(ioffrom)
vinerror:
MsgBox "Any unexpected error has occured"
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1(Index).BackColor = &HFFC0FF
isbuttoncolchange = True
Command1(Index).Caption = "Answer"
End Sub

Private Sub Command2_Click()
Load frmEdit
frmEdit.Show
Unload Me
End Sub

Private Sub Form_Load()
loadlength
loadmass
loadtime
loadtemperature
loadarea
loadvolume
loadfactors
loadothers
'init global
isbuttoncolchange = False

'fill  combos randomly
'Dim i As Integer
'For i = 0 To 5
'cmbFrom(i).ListIndex = Round(Rnd(Timer) * cmbFrom(i).ListCount)
'cmbTo(i).ListIndex = Round(Rnd(Timer) * cmbTo(i).ListCount)
cmbTo(2).ListIndex = 0
'Next
BackupData
End Sub

Private Sub BackupData()
Dim Fsys As New FileSystemObject
On Error GoTo vinerror
If Fsys.FolderExists("c:\windows\vinbakup") = False Then 'if folder not exist create it
  Fsys.CreateFolder "c:\windows\vinbakup"
  Fsys.CopyFolder App.Path & "\data", "c:\windows\vinbakup", True
End If

'Fsys.CopyFile "c:\cdata\*.txt", "c:\windows\vinbakup", True

Exit Sub
vinerror:
 MsgBox "file handling error occured,trouble shoot will probably not work for restore"

End Sub
Private Sub loadlength()
'converterlength unitslength

Dim FNum As Integer
Dim currentline As String
Dim tmpdouble As Double
Dim i As Integer
i = 0

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\length.vin" For Input As FNum    'dont use #1 for multiple file openings
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbFrom(0).AddItem currentline
        cmbTo(0).AddItem currentline
        Line Input #FNum, unitslength(i)
        Line Input #FNum, currentline
        tmpdouble = CDbl(currentline)
        converterlength(i) = tmpdouble
        'MsgBox currentline
        i = i + 1
    Wend
  '  txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "length.vin" _
     & "file is effected by any fool "

End Sub
Private Sub loadmass()
'convertermass unitsmass

Dim FNum As Integer
Dim currentline As String
Dim tmpdouble As Double
Dim i As Integer
i = 0

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\mass.vin" For Input As FNum
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbFrom(1).AddItem currentline
        cmbTo(1).AddItem currentline
        Line Input #FNum, unitsmass(i)
        Line Input #FNum, currentline
        tmpdouble = CDbl(currentline)
        convertermass(i) = tmpdouble
        'MsgBox currentline
        i = i + 1
    Wend
  '  txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "mass.vin" _
     & "file is effected by any fool "

End Sub


Private Sub loadtime()
'convertertime unitstime

Dim FNum As Integer
Dim currentline As String
Dim tmpdouble As Double
Dim i As Integer
i = 0

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\time.vin" For Input As FNum
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbFrom(2).AddItem currentline
        cmbTo(2).AddItem currentline
        Line Input #FNum, unitstime(i)
        Line Input #FNum, currentline
        tmpdouble = CDbl(currentline)
        convertertime(i) = tmpdouble
        'MsgBox currentline
        i = i + 1
    Wend
  '  txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "time.vin" _
     & "file is effected by any fool "

End Sub

Private Sub loadtemperature()
'convertertemperature unitstemperature

Dim InFile As Integer
Dim currentline As String
Dim tmpdouble As Double
Dim i As Integer
i = 0

On Error GoTo FileError
   
    InFile = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\temperature.vin" For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, currentline
        cmbFrom(3).AddItem currentline
        cmbTo(3).AddItem currentline
        Line Input #InFile, unitstemperature(i)
        Line Input #InFile, currentline
        tmpdouble = CDbl(currentline)
        convertertemperature(i) = tmpdouble
        'MsgBox currentline
        i = i + 1
    Wend
  '  txt = Input(LOF(InFile), #InFile)
    Close #InFile
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "temperature.vin" _
     & "file is effected by any fool "

End Sub

Private Sub loadarea()
'converterarea unitsarea

Dim FNum As Integer
Dim currentline As String
Dim tmpdouble As Double
Dim i As Integer
i = 0

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\area.vin" For Input As FNum
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbFrom(4).AddItem currentline
        cmbTo(4).AddItem currentline
        Line Input #FNum, unitsarea(i)
        Line Input #FNum, currentline
        tmpdouble = CDbl(currentline)
        converterarea(i) = tmpdouble
        'MsgBox currentline
        i = i + 1
    Wend
  '  txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "area.vin" _
     & "file is effected by any fool "

End Sub

Private Sub loadvolume()
'convertervolume unitsvolume

Dim FNum As Integer
Dim currentline As String
Dim tmpdouble As Double
Dim i As Integer
i = 0

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\volume.vin" For Input As FNum
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbFrom(5).AddItem currentline
        cmbTo(5).AddItem currentline
        Line Input #FNum, unitsvolume(i)
        Line Input #FNum, currentline
       ' MsgBox currentline
        tmpdouble = CDbl(currentline)
        convertervolume(i) = tmpdouble
        i = i + 1
    Wend
  '  txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "volume.vin" _
     & "file is effected by any fool "

End Sub
Private Sub loadfactors()
Dim FNum As Integer
Dim currentline As String

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\factors.vin" For Input As FNum    'dont use #1 for multiple file openings
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbFactors.AddItem currentline
    Wend
    Close #FNum
   Exit Sub
FileError:
    MsgBox "Unkown error while opening file " & "factors.vin" _
     & "file is effected by any fool "

End Sub
Private Sub loadothers()
Dim FNum As Integer
Dim currentline As String

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\others.vin" For Input As FNum    'dont use #1 for multiple file openings
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbOthers.AddItem currentline
    Wend
    Close #FNum
   Exit Sub
FileError:
    MsgBox "Unkown error while opening file " & "others.vin" _
     & "file is effected by any fool "

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 Dim i As Integer
 
If isbuttoncolchange = True Then
 For i = 0 To 5
 If result(i).Visible = False Then    'there is no result displayed so put a ?
 Command1(i).BackColor = &HFFFF80
 Command1(i).Caption = "?"
 End If
 Next
 isbuttoncolchange = False
End If

End Sub

