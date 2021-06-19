VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Password Manager Login"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "Create New User"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log In"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.ComboBox cmbUser 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Default User"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()
userIndex = cmbUser.ListIndex + 1  'used when user deletion
If colPassword.Item(userIndex) = txtPassword.Text Then
   'MsgBox "MATCHED" & userIndex
   Load frmMan
   frmMan.Visible = True
   Me.Hide
End If
End Sub

Private Sub cmdNew_Click()
Load frmNew
 frmNew.Visible = True
 Me.Hide
'colUserName.Remove (userIndex)
'Dim i As Integer
'Dim str As String
'For i = 1 To colUserName.Count
' str = str & colUserName.Item(i)
'Next
'MsgBox str
End Sub

Private Sub Form_Load()

 cmbUser.Clear
 'cmbUser.AddItem "Default User"
loadUserName
ReLoadCombo
End Sub
Public Sub ReLoadCombo()
Dim i As Integer
cmbUser.Clear
For i = 1 To colUserName.Count
 cmbUser.AddItem colUserName.Item(i)
Next
 cmbUser.ListIndex = 0
End Sub
Public Sub loadUserName()
Dim fnum As Integer
Dim currentline As String

   'On Error GoTo fileerror
   
    fnum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\users.vin" For Input As fnum    'dont use #1 for multiple file openings
    'While Not EOF(FNum)
    Line Input #fnum, currentline  '<user>
   Do While Not "</user>" = currentline
      Line Input #fnum, currentline
      If "</user>" = Trim(currentline) Then Exit Do
      colUserName.Add Decrypt(currentline, AscB("V") + 80)
      Line Input #fnum, currentline
      colPassword.Add Decrypt(currentline, AscB("V") + 80)
   Loop
   Close #fnum
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Caption = "Saving Data"
Dim vin As Byte
vin = AscB("V") + 80

Dim fnum As Integer
Dim i As Integer

   'On Error GoTo fileerror
   
    fnum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\users.vin" For Output As fnum      'dont use #1 for multiple file openings
    Print #fnum, "<user>"
    For i = 1 To colUserName.Count
     
      Print #fnum, Encrypt(colUserName.Item(i), vin)
      Print #fnum, Encrypt(colPassword.Item(i), vin)
      
    Next
    Print #fnum, "</user>"
    Close #fnum

End Sub
