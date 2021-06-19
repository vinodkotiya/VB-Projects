VERSION 5.00
Begin VB.Form frmMan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Password Manager"
   ClientHeight    =   5880
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   720
      Width           =   615
   End
   Begin VB.CheckBox chkPass 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   1320
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Modify"
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   15
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Delete"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   14
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Make"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "New"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtPass 
      Height          =   525
      Index           =   4
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   8
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ComboBox cmbPass 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VIN Password Manager For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   -30
      TabIndex        =   19
      Top             =   80
      Width           =   5415
   End
   Begin VB.Label Label 
      Caption         =   "Login Name"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Comments"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Any User Name"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "For Application / WebSite / Email"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Show Password #"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VIN Password Manager For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   5415
   End
   Begin VB.Menu user 
      Caption         =   "User"
      Begin VB.Menu mnuUser 
         Caption         =   "Create New User"
         Index           =   0
      End
      Begin VB.Menu mnuUser 
         Caption         =   "Delete Current User"
         Index           =   1
      End
      Begin VB.Menu mnuUser 
         Caption         =   "Log Off...."
         Index           =   2
      End
      Begin VB.Menu mnuUser 
         Caption         =   "Switch User..."
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colPass As New Collection
Dim colApp As New Collection
Dim colUser As New Collection
Dim colLogin As New Collection
Dim colCom As New Collection
Dim totalPass As Long


Private Sub chkPass_Click()
'chkPass.Value = Not chkPass.Value
If chkPass.Value Then
  txtPass(0).PasswordChar = "*"
Else
 txtPass(0).PasswordChar = ""
End If
 
End Sub

Private Sub cmbPass_Click()
txtPass(0).Text = colPass.Item(cmbPass.ListIndex + 1)
 txtPass(1).Text = colApp.Item(cmbPass.ListIndex + 1)
 txtPass(2).Text = colUser.Item(cmbPass.ListIndex + 1)
 txtPass(3).Text = colLogin.Item(cmbPass.ListIndex + 1)
 txtPass(4).Text = colCom.Item(cmbPass.ListIndex + 1)
End Sub

Private Sub cmdOp_Click(Index As Integer)
If Index = 0 Then  'new
 cmdOp(0).Enabled = False
 cmdOp(1).Enabled = True
 cmdOp(2).Enabled = False
 cmdOp(3).Enabled = False
 
ElseIf Index = 1 Then  'make
 cmdOp(0).Enabled = True
 cmdOp(1).Enabled = False
 cmdOp(2).Enabled = True
 cmdOp(3).Enabled = True
 'KeyCode = AscB("VINOD")
 colPass.Add txtPass(0).Text
 colApp.Add txtPass(1).Text
 colUser.Add txtPass(2).Text
 colLogin.Add txtPass(3).Text
 colCom.Add txtPass(4).Text
 totalPass = totalPass + 1
 
 cmbPass.AddItem totalPass
 cmbPass.ListIndex = totalPass - 1
ElseIf Index = 2 Then  'delete
 colPass.Remove cmbPass.ListIndex + 1
 colApp.Remove cmbPass.ListIndex + 1
 colUser.Remove cmbPass.ListIndex + 1
 colLogin.Remove cmbPass.ListIndex + 1
 colCom.Remove cmbPass.ListIndex + 1
 cmbPass.RemoveItem (totalPass - 1)
 totalPass = totalPass - 1
 cmbPass.ListIndex = totalPass - 1
 cmbPass_Click
End If
End Sub

Private Sub Command1_Click()
txtPass(3).Text = Encrypt(txtPass(0).Text, AscB("V") + 80)

End Sub

Private Sub loadData()
Dim fnum As Integer
Dim currentline As String
Dim vin As Byte
If userIndex = 1 Then
   vin = AscB("V") 'default user
Else
 vin = AscB(colUserName.Item(userIndex)) + 80
End If
   'On Error GoTo fileerror
   
    fnum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\" & colUserName.Item(userIndex) & ".vin" For Binary As fnum     'dont use #1 for multiple file openings
    'While Not EOF(FNum)
   ' MsgBox colUserName.Item(userIndex)
    Line Input #fnum, currentline  '<user>
   Do While Not "</end>" = currentline
      Line Input #fnum, currentline
      If "</end>" = Trim(currentline) Then Exit Do
      colPass.Add Decrypt(currentline, vin)
      Line Input #fnum, currentline
      colApp.Add Decrypt(currentline, vin)
      Line Input #fnum, currentline
      colUser.Add Decrypt(currentline, vin)
      Line Input #fnum, currentline
      colLogin.Add Decrypt(currentline, vin)
      Line Input #fnum, currentline
      colCom.Add Decrypt(currentline, vin)
      totalPass = totalPass + 1
      cmbPass.AddItem totalPass
      cmbPass.ListIndex = 0
   Loop
   Close #fnum
End Sub

Private Sub Form_Load()
If userIndex = 1 Then 'default user
   mnuUser(1).Enabled = False
Else
  mnuUser(1).Enabled = True
End If
 lblCap(0).Caption = "VIN Password Manager For " & colUserName.Item(userIndex)
  lblCap(1).Caption = "VIN Password Manager For " & colUserName.Item(userIndex)
totalPass = 0
loadData

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Caption = "Saving Data"
Dim vin As Byte
If userIndex = 1 Then
   Exit Sub 'default user
Else
 vin = AscB(colUserName.Item(userIndex)) + 80
End If

Dim fnum As Integer
Dim i As Integer

   'On Error GoTo fileerror
   
    fnum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\" & colUserName.Item(userIndex) & ".vin" For Output As fnum      'dont use #1 for multiple file openings
    Print #fnum, "<start>"
    For i = 1 To colPass.Count
     
      Print #fnum, Encrypt(Trim(colPass.Item(i)), vin)
      Print #fnum, Encrypt(Trim(colApp.Item(i)), vin)
      Print #fnum, Encrypt(Trim(colUser.Item(i)), vin)
      Print #fnum, Encrypt(Trim(colLogin.Item(i)), vin)
      Print #fnum, Encrypt(Trim(colCom.Item(i)), vin)
    Next
    Print #fnum, "</end>"
    Close #fnum

End Sub

Private Sub mnuUser_Click(Index As Integer)
If Index = 0 Then
 Unload Me
 Load frmNew
 frmNew.Visible = True
ElseIf Index = 1 Then
 Dim reply As Integer
  reply = MsgBox("Do you want to delete the current user " & colUserName.Item(userIndex) & vbCrLf & "Note that it will also delete all the data for " & colUserName.Item(userIndex), vbYesNo, "Prompt For Deleting User")
  If reply = vbYes Then         'yes
    Dim fsys As New FileSystemObject
    fsys.DeleteFile App.Path & "\data\" & colUserName.Item(userIndex) & ".vin", True
    colUserName.Remove (userIndex)
    colPassword.Remove (userIndex)
    userIndex = 1 'so that exit from query unload
    Unload Me
    frmLogin.ReLoadCombo
    frmLogin.Show
 End If
ElseIf Index = 2 Then
 Unload Me
 frmLogin.ReLoadCombo
 frmLogin.Show
End If
End Sub

Private Sub txtPass_Change(Index As Integer)
If Len(txtPass(Index).Text) > 0 Then
 txtPass(Index).BackColor = &HFFFF00
Else
 txtPass(Index).BackColor = vbWhite
End If
End Sub
