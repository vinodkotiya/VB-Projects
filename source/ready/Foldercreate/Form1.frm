VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create folder and copy c:\cdata"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Fsys As New FileSystemObject

MsgBox Fsys.FileExists("c:\cdata\*.txt")
'MsgBox Fsys.FolderExists("c:\aldus")
'Fsys.CopyFolder "c:\cdata", "c:\aldus"
If Fsys.FolderExists("c:\windows\vinbakup") = False Then
MsgBox "creating folder"
 Fsys.CreateFolder "c:\windows\vinbakup"
End If
'Fsys.CopyFolder "c:\cdata", "c:\windows\vinbakup", True
Fsys.CopyFile "c:\cdata\*.txt", "c:\windows\vinbakup", True
End Sub
