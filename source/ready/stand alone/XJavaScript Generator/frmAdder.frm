VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdder 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Scripts"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000FF00&
      Caption         =   ": : Add this Script in to Script Collection : :"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Please Enter Default Values of variables to which you have assigned the value input1 ,2,3,4"
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtVar 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   27
         Text            =   "Color"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtVar 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   26
         Text            =   "image"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtVar 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   25
         Text            =   "Speed"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtVar 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   24
         Text            =   "eg : Msg"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   23
         Text            =   "blue"
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   22
         Text            =   "path of image.jpg"
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   21
         Text            =   "500"
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   20
         Text            =   "Type your message here"
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox checkVar 
         BackColor       =   &H00000000&
         Caption         =   "Name of variable1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox checkVar 
         BackColor       =   &H00000000&
         Caption         =   "Name of variable2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox checkVar 
         BackColor       =   &H00000000&
         Caption         =   "Name of variable3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox checkVar 
         BackColor       =   &H00000000&
         Caption         =   "Name of variable4"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Default Value"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Default Value"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Default Value"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   33
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Default Value"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   6735
      Begin VB.OptionButton opt 
         BackColor       =   &H00000000&
         Caption         =   "No"
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   6120
         TabIndex        =   18
         Top             =   -20
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00000000&
         Caption         =   "Yes"
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   0
         Left            =   5400
         TabIndex        =   17
         Top             =   -20
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Is your script contain any variables whose values can be changed by user "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   5295
      End
   End
   Begin VB.TextBox txtDiscription 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Frame frmCat 
      BackColor       =   &H00000000&
      Caption         =   "Category"
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Select any one category then choose any script from the list"
      Top             =   0
      Width           =   6855
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Mouse Cursor"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Backgrounds"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Date and Time"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Links and Menus"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Text"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Status Bar"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Image and Sound"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Various"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      _Version        =   393217
      BackColor       =   16777152
      ScrollBars      =   3
      TextRTF         =   $"frmAdder.frx":0000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Discription of Script"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Place the code of your script here"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Name of the New Script in selected Category"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "frmAdder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblInput_Click(Index As Integer)

End Sub

Private Sub checkVar_Click(Index As Integer)
If checkVar(Index).Value Then
 txtVar(Index).Enabled = True
  txtInput(Index).Enabled = True
  txtVar(Index).Text = ""
  txtInput(Index).Text = ""
 Else
 txtVar(Index).Enabled = False
  txtInput(Index).Enabled = False
 End If
End Sub

Private Sub Form_Load()
opt_Click (1)
End Sub

Private Sub opt_Click(Index As Integer)
If Index = 0 Then
 MsgBox "It means your script contain some user defined variables. ? Let us see" & vbCrLf & _
 "Suppose your script contain a sting variable" & Chr(34) & " msg " & Chr(34) & "which store any message " & vbCrLf & _
 "--->>In script it would look like this msg = " & Chr(34) & "it is a message<<---" & Chr(34) & vbCrLf & _
 "Now what you have to do is type " & Chr(34) & "input1" & Chr(34) & " (if it is 1st variable) inplace of " & Chr(34) & "it is a message" & Chr(34) & vbCrLf & _
 "----->>So the new seen look like this msg = " & Chr(34) & "input1" & Chr(34) & " inside double quotes <<-----" & vbCrLf & vbCrLf & _
 "Similarly if  script contain another  variable" & Chr(34) & " Speed " & Chr(34) & " which store any integer value " & vbCrLf & _
 "--->>In script it would look like this Speed = 50  <<---" & vbCrLf & _
 "Now what you have to do is type input2 (if it is 2nd variable) inplace of" & Chr(34) & "50" & Chr(34) & vbCrLf & _
 "----->>So the new seen like this Speed = input2" & "   without double quotes <<-----" & vbCrLf & vbCrLf & _
 "Apply  above procedure inside your script code and remember you can use only maximum 4 or less user defined variable and ther values will be " & vbCrLf & _
 "input1 , input2 , input3 , input4  for integer like variables" & vbCrLf & _
  Chr(34) & "input1" & Chr(34) & " , " & Chr(34) & "input2" & Chr(34) & " , " & Chr(34) & "input3" & Chr(34) & "  , " & Chr(34) & "input4" & Chr(34) & "  for string variables" & vbCrLf & vbCrLf & _
 "Then in textboxes enter name of your variable and its default value "
 Frame2.Visible = True
 cmdAdd.Top = Frame2.Top + Frame2.Height + 40
 
 ElseIf Index = 1 Then
 Frame2.Visible = False
 cmdAdd.Top = Frame2.Top
 
 End If
 frmAdder.Height = cmdAdd.Top + cmdAdd.Height + 600
End Sub
