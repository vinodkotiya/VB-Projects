VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   ForeColor       =   &H0000FF00&
   Icon            =   "frmDeInstaller.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmDeInstaller.frx":0ECA
   ScaleHeight     =   2370
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No"
      Height          =   255
      Index           =   1
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdOp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Yes"
      Height          =   255
      Index           =   0
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblAsk 
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Sure to uninstall the application"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirming Deletion"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colFile As New Collection

Private Sub cmdOp_Click(Index As Integer)
If Index = 0 Then
 cmdOp(0).Visible = False
 cmdOp(1).Visible = False
 lblAsk.Visible = False
 RemoveFiles
 MsgBox "Uninstallation Successfully Complete...."
 End
Else
'Dim i As Integer
'Dim str As String
'For i = 1 To colFile.Count
' str = str & colFile.Item(i) & vbCrLf
'N'ext
'MsgBox str & colFile.Count
 End
End If

End Sub
Private Sub RemoveFiles()
On Error Resume Next
Dim i As Integer
Dim fsys As New FileSystemObject
For i = 1 To colFile.Count
  fsys.DeleteFile colFile.Item(i), True
 lblStatus.Caption = "Removing File " & colFile.Item(i)
Next
fsys.DeleteFile App.Path & "\uninstall.vin"
fsys.DeleteFile App.Path & "\link.vbs"
''''delete subfolders
Dim AllFolders As Folders 'contain all subfolders of thisfolder
Dim fold As Folder  'i as integer type variable
Set fold = fsys.GetFolder(App.Path)
Set AllFolders = fold.SubFolders
For Each fold In AllFolders
  fsys.DeleteFolder fold.Path, True
  lblStatus.Caption = "Removing Folder " & fold.Name
Next
lblStatus.Caption = "Uninstallation Complete..."
End Sub
Private Sub LoadFile()
Dim Fnum As Integer
Dim currLine As String
Dim fsys As New FileSystemObject
Fnum = FreeFile
Open App.Path & "\uninstall.vin" For Input As Fnum
     ''''''''' remove redundance ''''''''''''
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     
     '''''''''''''''''''''''''removed'''''''''''''
     
     Line Input #Fnum, currLine  '<JAIMATADI>
      'MsgBox currLine
     Line Input #Fnum, currLine
     lblAsk.Caption = lblAsk.Caption & currLine
     Do While (Not "</JAIMATADI>" = Trim(currLine)) 'Or (Not EOF(FNum))
      Line Input #Fnum, currLine
      If "</JAIMATADI>" = Trim(currLine) Then Exit Do
      colFile.Add Trim(currLine)
     Loop
Close Fnum

End Sub
Private Sub cmdOp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOp(Index).BackColor = &HC0E0FF
End Sub

Private Sub Form_Load()
LoadFile
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOp(0).BackColor = vbWhite
cmdOp(1).BackColor = vbWhite
End Sub
