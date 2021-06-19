VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000FF00&
      Caption         =   "Add SubMenu's"
      Height          =   1335
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H008080FF&
      Caption         =   "Delete SubMenu's"
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu smnNew 
         Caption         =   "New"
      End
      Begin VB.Menu smnOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu smnSave 
         Caption         =   "Save"
      End
      Begin VB.Menu smnSaveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu smnImport 
         Caption         =   "Import"
      End
      Begin VB.Menu smnExport 
         Caption         =   "Export"
      End
      Begin VB.Menu smnExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu smnSub 
         Caption         =   "Sub menu"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim submenu As Integer


Private Sub cmdDel_Click()
If smnNew.Visible = True Then
   smnNew.Visible = False
ElseIf smnOpen.Visible = True Then
       smnOpen.Visible = False
ElseIf smnSave.Visible = True Then
       smnSave.Visible = False
ElseIf smnSaveas.Visible = True Then
       smnSaveas.Visible = False
ElseIf smnImport.Visible = True Then
       smnImport.Visible = False
ElseIf smnExport.Visible = True Then
       smnExport.Visible = False
ElseIf smnExit.Visible = True Then
       smnExit.Visible = False
End If
If submenu = 0 Then
        MsgBox "can't delete"
    Exit Sub
    End If
    
    Unload smnSub(submenu)
    submenu = submenu - 1

End Sub

Private Sub cmdAdd_Click()
If smnNew.Visible = False Then
   smnNew.Visible = True
ElseIf smnOpen.Visible = False Then
       smnOpen.Visible = True
ElseIf smnSave.Visible = False Then
       smnSave.Visible = True
ElseIf smnSaveas.Visible = False Then
       smnSaveas.Visible = True
ElseIf smnImport.Visible = False Then
       smnImport.Visible = True
ElseIf smnExport.Visible = False Then
       smnExport.Visible = True
ElseIf smnExit.Visible = False Then
       smnExit.Visible = True
Else

submenu = submenu + 1
    If submenu = 1 Then smnSub(0).Caption = "Sub menu"
    Load smnSub(submenu)
    smnSub(submenu).Caption = "Sub Menu " & submenu
    
End If
End Sub
  

Private Sub RemoveCommand_Click()
    End Sub
