VERSION 5.00
Begin VB.Form frmLock 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim thisFile As File
Dim fsys As New FileSystemObject
Set thisFile = fsys.GetFile(App.Path & "\INTRO.DAT")
MsgBox thisFile.Attributes & thisFile.Path
thisFile.Attributes = Volume

End Sub

Private Sub Command2_Click()
Dim inFile As Integer, outfileNumber As Integer
    inFile = FreeFile
    Open App.Path & "\INTRO.DAT" For Binary Access Write As #inFile
'    Put #inFile, , "vinod kotiya" 'Input$(eachFileSizeBe ,infileNumber) 'outChar'put all the datasize given by user on each file
MsgBox Loc(inFile)
'Print #inFile, "vinod"
'   Lock #inFile    'lock till app running
 '   MsgBox " "
    Close inFile
End Sub
