VERSION 5.00
Begin VB.Form frmShoot 
   BackColor       =   &H00FF0000&
   Caption         =   "TROUBLESHOOTER"
   ClientHeight    =   5325
   ClientLeft      =   1590
   ClientTop       =   2205
   ClientWidth     =   7860
   Icon            =   "frmShoot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7860
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Restore VIN Convert Centre to its Ideal Condition"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Restore Fone Directory Database (Yesterday)"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton cmdTrouble 
      BackColor       =   &H0080FF80&
      Caption         =   "My system has MIDI Device and speakers are turned on but i can't here any sound."
      Height          =   375
      Index           =   5
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   7095
   End
   Begin VB.CommandButton cmdTrouble 
      BackColor       =   &H0080FF80&
      Caption         =   "VIN Convert Centre Fails to convert any value or giving error on startup."
      Height          =   375
      Index           =   4
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   7095
   End
   Begin VB.CommandButton cmdTrouble 
      BackColor       =   &H0080FF80&
      Caption         =   "VIN WEB COMPILER is not changing Background or Text Color of webpages not created by it."
      Height          =   375
      Index           =   3
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   7095
   End
   Begin VB.CommandButton cmdTrouble 
      BackColor       =   &H0080FF80&
      Caption         =   "VIN Remind Me Later is not showing messages when i start computer."
      Height          =   375
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   7095
   End
   Begin VB.CommandButton cmdTrouble 
      BackColor       =   &H0080FF80&
      Caption         =   "My VIN FONE DIRECTORY is not showing fone numbers or not opening database."
      Height          =   375
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   7095
   End
   Begin VB.CommandButton cmdTrouble 
      BackColor       =   &H0080FF80&
      Caption         =   "There is no Background Music is Playing Back"
      Height          =   375
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Height          =   4215
      Left            =   120
      Top             =   120
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      DrawMode        =   15  'Merge Pen Not
      Height          =   4215
      Left            =   120
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmShoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTrouble_Click(Index As Integer)
If Index = 0 Then
 MsgBox "Make sure that your system must have a MIDI palayback Device" & Chr(13) & _
    "You can find your MIDI Device In ControlPanel->Sound and Audio Device."
ElseIf Index = 5 Then
MsgBox "Open your master volume panel from ControlPanel->Sound and Audio Device." & Chr(13) & _
    "Make the WAVE slidebar to its maximum "
ElseIf Index = 1 Then
MsgBox "This means that the fone directory database is corrupted / changed" & Chr(13) & _
    "By opening it.But don't worry VIN UTILITY KIT v1.0 creates backup of your directory database everyday" & Chr(13) & _
    "So you can restore all of your (Yesterdays) records " & Chr(13) & _
    "To do this please click on button ''RESTORE (to yesterday) FONE Directory Database'' "
ElseIf Index = 2 Then
MsgBox "Probably the shortcut of vinreminder.exe is deleted from start ->program -> startup" & Chr(13) & _
    "To resolve this problem place an shortcut of application VIN Remind Me Later in startup"
ElseIf Index = 3 Then
MsgBox "It is due to If page contain javascript hence <BODY...> tag comes after 5-6 lines " & Chr(13) & _
 "This problem can be resolved by placing the <BODY...> tag" & Chr(13) & _
 "after 5-6 lines from starting."
ElseIf Index = 4 Then
MsgBox "Probably someone disturb the data files" & Chr(13) & _
    "To resolve this problem click on ''Restore All'' or reinstall this software again "

End If
End Sub

Private Sub Command1_Click()
Dim Fsys As New FileSystemObject
Dim reply As Integer
On Error GoTo vinerror
If Fsys.FolderExists("c:\windows\vinbakup") = False Then 'if folder not exist create it
  MsgBox "Backup was not created so unable to Restore"
End If
'Fsys.CopyFolder App.Path & "\data", "c:\windows\vinbakup", True
reply = MsgBox("Are you sure to restore your yesterday's Fone Directory", vbYesNo)
If reply = vbYes Then
Fsys.CopyFile "c:\windows\vinbakup\directory.mdb", App.Path & "\data\", True
MsgBox "Restore successfully"
End If
'MsgBox reply
Exit Sub
vinerror:
 MsgBox "file handling error occured"

End Sub

Private Sub Command2_Click()
Dim Fsys As New FileSystemObject
Dim reply As Integer
On Error GoTo vinerror
If Fsys.FolderExists("c:\windows\vinbakup") = False Then 'if folder not exist create it
  MsgBox "Backup was not created so unable to Restore"
End If
'Fsys.CopyFolder App.Path & "\data", "c:\windows\vinbakup", True
reply = MsgBox("Are you sure to restore all data files to previous state", vbYesNo)
If reply = vbYes Then
Fsys.CopyFile "c:\windows\vinbakup\time.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\length.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\mass.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\area.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\temperature.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\volume.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\factors.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\others.vin", App.Path & "\data\", True
MsgBox "Restore successfully completes"
End If
Exit Sub
vinerror:
 MsgBox "file handling error occured"


End Sub
