VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Starter"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   Icon            =   "frmStarter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''////////////''''''''''''''///////////////////////////////
'''   This exe fill the file loadrem.vin so that main application can start
'othervise on direct clicking main application will check dates to display if no message than unload itself
''because loadrem is empty

Option Explicit

Private Sub Form_Load()
Dim fnum As Integer

On Error GoTo FileError
  fnum = FreeFile
  Open App.Path & "\data\loadrem.vin" For Output As #1
   Print #fnum, "loadremember Vinod Kotiya " _
               & " is calling you from a program "
 Close #fnum
 fnum = Shell(App.Path & "\vinreminder.exe", vbNormalFocus)
 End
 Exit Sub

FileError:
    
    MsgBox "Unable to write . Unkown error while filling file " & "data\loadrem.vin"

End Sub
