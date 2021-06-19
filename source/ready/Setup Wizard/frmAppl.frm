VERSION 5.00
Begin VB.Form frmAppl 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Step 3 :"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   3045
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFF80&
      Caption         =   "Preview"
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remove Selected"
      Height          =   495
      Index           =   1
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdDll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add"
      Height          =   615
      Index           =   0
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox comboDll 
      Height          =   315
      Left            =   3480
      TabIndex        =   25
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtDll 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   480
      TabIndex        =   24
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton cmdOpendll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open System/dll Files"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Open the files to be splitted"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ListBox listDll 
      BackColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   480
      MultiSelect     =   2  'Extended
      TabIndex        =   21
      ToolTipText     =   "Display the file to be splitted .Select the file here."
      Top             =   2400
      Width           =   5535
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< Back"
      Height          =   375
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next >>"
      Height          =   375
      Index           =   1
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame frLnk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create Shortcuts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   480
      TabIndex        =   7
      Top             =   3720
      Width           =   6615
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         Height          =   255
         Index           =   3
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.ListBox listLink 
         Height          =   735
         ItemData        =   "frmAppl.frx":0000
         Left            =   120
         List            =   "frmAppl.frx":0002
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   1320
         Width           =   6375
      End
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         Height          =   255
         Index           =   2
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         Height          =   255
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdLink 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create"
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtLink 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   12
         Text            =   "My App.exe"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtLink 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   10
         Text            =   "\vinsoft\My App"
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Source Dir\"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Link With File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name of Shortcut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtTarget 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Text            =   "C:\Program Files\VINSOFT"
      Top             =   3360
      Width           =   4695
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1920
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   3720
   End
   Begin VB.Label lblApp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Where To INSTALL"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   23
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Step3>>    APPLICATION INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   20
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Installation Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Your dll/ocx files And specify that where to install these files"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Folder where your Application (exe) and other data files and subfolders are placed"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   6495
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1200
      X2              =   1440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1320
      X2              =   1320
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1320
      X2              =   3360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Source Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmAppl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim editListno As Integer
Dim notWhite As Boolean 'true when not white


Private Sub cmdDir_Click(Index As Integer)
If Index = 1 Then
frmButton.imgStepOver_Click (3)
Else
 frmButton.imgStepOver_Click (1)
End If
 
End Sub

Private Sub cmdDir_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 cmdDir(Index).BackColor = &HE0E0E0
 notWhite = True
End If
End Sub

Private Sub cmdDll_Click(Index As Integer)
isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
Dim i As Integer
If Index = 0 Then
If Trim(txtDll.Text) = "" Then
 MsgBox "No system files are opened.First Open any System/dll File."
 Exit Sub
End If
If Trim(comboDll.Text) = "" Then
 MsgBox "Specify Where to install System/dll File."
 Exit Sub
End If
i = InStrRev(Trim(txtDll.Text), "\")
listDll.AddItem Trim(comboDll.Text) & "\" & Right(Trim(txtDll.Text), Len(Trim(txtDll.Text)) - i)
colDll.Add Trim(txtDll.Text)
ElseIf Index = 1 Then
If listDll.SelCount = 0 Then
 MsgBox "First Select Any Entry to delete"
  Exit Sub
End If
For i = listDll.ListCount - 1 To 0 Step -1
  If listDll.Selected(i) = True Then
  listDll.RemoveItem (i)
  colDll.Remove i + 1
  End If
 Next
 'Dim txt As String
 'For i = 1 To colDll.Count
 'txt = txt & colDll.Item(i) & vbCrLf
 'Next
 'MsgBox colDll.Count & txt
End If
End Sub

Private Sub cmdLink_Click(Index As Integer)

Dim pos As Integer
Dim posD As Integer
If Index = 0 Or Index = 3 Then 'create
    If Index = 3 And editListno <> -5 Then
     listLink.RemoveItem (editListno)
     editListno = -5     'updated
    End If
    
  If "\" <> Left(Trim(txtLink(0).Text), 1) Then txtLink(0).Text = "\" & Trim(txtLink(0).Text)
  If Left(Trim(txtLink(1).Text), 1) = "\" Then Trim(txtLink(1).Text) = Right(Trim(txtLink(1).Text), Len(Trim(txtLink(1).Text)) - 1)
 
  listLink.AddItem Combo1.Text & txtLink(0).Text & "  #  " & txtLink(1).Text
  isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
ElseIf Index = 1 Then
 pos = InStr(1, listLink.Text, "\")
 posD = InStr(1, listLink.Text, "#")
 txtLink(0).Text = Trim(Mid(listLink.Text, pos, posD - pos))
 txtLink(1).Text = Trim(Right(listLink.Text, Len(listLink.Text) - posD))
 If "D" = Left(listLink.Text, 1) Then
  Combo1.ListIndex = 0
 ElseIf "P" = Left(listLink.Text, 1) Then
  Combo1.ListIndex = 1
 ElseIf "S" = Left(listLink.Text, 1) Then
  Combo1.ListIndex = 2
 End If
 'listLink.RemoveItem (listLink.ListIndex)
 editListno = listLink.ListIndex
ElseIf Index = 2 Then
 listLink.RemoveItem (listLink.ListIndex)
 txtLink(0).Text = ""
 txtLink(1).Text = ""
 isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
End If
End Sub

Private Sub cmdOpendll_Click()

mdifrmMain.CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   mdifrmMain.CommonDialog1.fileName = ""
   mdifrmMain.CommonDialog1.Filter = "*.dll|*.dll|*.ocx|*.ocx|All Files|*.*"
   mdifrmMain.CommonDialog1.ShowOpen
   If mdifrmMain.CommonDialog1.fileName = "" Or Err.Number = cdlCancel Then
      MsgBox "No file is opened"
      Exit Sub
   End If
   
    txtDll.Text = mdifrmMain.CommonDialog1.fileName
   

End Sub

Private Sub cmdPreview_Click()
frmPrev.Visible = True
 frmPrev.step3
End Sub
Private Sub cmdPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = False Then
 notWhite = True
 cmdPreview.BackColor = &HFF00FF
End If
End Sub

Private Sub Command1_Click()
'Dim i As String
'i = CreateShortcutScript
'MsgBox i
'For i = colDll.Count To 1 Step -1
 '   colDll.Remove i
 'Next
 'Dim txt As String
 'For i = 1 To colDll.Count
 'txt = txt & colDll.Item(i) & vbCrLf
 'Next
 'MsgBox colDll.Count & txt

 'For i = 0 To listLink.ListCount - 1
 '  MsgBox listLink.List(i)
 'Next
 
End Sub

Private Sub Dir1_Change()
isCompiled = False
 ChDir Dir1.path
'Dim fsys As New FileSystemObject
'Dim thisfolder As Folder
'Dim allFolders As Folders'contain all subfolders of thisfolder
'Dim fold As Folder 'i as integer type variable
'Set thisfolder = fsys.GetFolder(Dir1.Path)
'Set allfolders = thisfolder.SubFolders
'Dim txt As String
'Dim i As Integer
'MsgBox allfolders.Count
'For Each fold In allfolder
' txt = txt & fold.Name & vbCrLf
'Next
'MsgBox txt
End Sub

Private Sub Drive1_Change()
On Error GoTo vinerror
ChDrive Dir1.path
    Dir1.path = Drive1.Drive
    Dir1.Refresh
  Exit Sub
vinerror:
  MsgBox "There is no disk in drive"
End Sub

Private Sub Form_Load()
Combo1.AddItem "Desktop"
Combo1.AddItem "Programs"
Combo1.AddItem "Startup"
Combo1.ListIndex = 0
comboDll.AddItem "c:\windows\system"
comboDll.AddItem "c:\Windows\system32"
comboDll.ListIndex = 1
Me.Picture = LoadPicture(App.path & "\data\back.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If notWhite = True Then
 cmdDir(0).BackColor = vbWhite
 cmdDir(1).BackColor = vbWhite
 cmdPreview.BackColor = 16777088
 notWhite = False
End If
End Sub

Private Sub listLink_Click()
Dim i As Integer
'MsgBox listLink.Text
For i = 0 To listLink.ListCount - 1
 If i <> listLink.ListIndex Then
  listLink.Selected(i) = False
 End If
 
Next
End Sub
Public Function MakeFile3() As String
Dim txtSave As String
Dim i As Integer
txtSave = "<<<Application>>>" & vbCrLf
txtSave = txtSave & Dir1.path & vbCrLf
txtSave = txtSave & txtTarget.Text & vbCrLf
txtSave = txtSave & " <ListLink>" & vbCrLf
For i = 0 To listLink.ListCount - 1
 txtSave = txtSave & listLink.List(i) & vbCrLf
Next
txtSave = txtSave & " </ListLink>" & vbCrLf
txtSave = txtSave & " <ListDll>" & vbCrLf
For i = 0 To listDll.ListCount - 1
 txtSave = txtSave & listDll.List(i) & vbCrLf
Next
txtSave = txtSave & " </ListDll>" & vbCrLf
txtSave = txtSave & " <colDll>" & vbCrLf
For i = 1 To colDll.Count
 txtSave = txtSave & colDll.item(i) & vbCrLf
Next
txtSave = txtSave & " </colDll>" & vbCrLf
MakeFile3 = txtSave
End Function

Public Function CreateShortcutScript() As String
If listLink.ListCount = 0 Then
 CreateShortcutScript = ""
 Exit Function
End If
Dim ShortcutScript As String 'store vbscript of shortcuts
ShortcutScript = "Set WShell = Wscript.CreateObject(""Wscript.Shell"")" & vbCrLf & _
      "pos = InStrRev(wscript.ScriptFullName, ""\"")" & vbCrLf & _
      "Path = Left(wscript.ScriptFullName, pos)" & vbCrLf & _
      "Set Fsys = CreateObject(""scripting.filesystemobject"")" & vbCrLf & _
      "strDesktop = WShell.SpecialFolders(""Desktop"")" & vbCrLf & _
      "strStartup = WShell.SpecialFolders(""Startup"")" & vbCrLf & _
      "strPrograms = WShell.SpecialFolders(""Programs"")" & vbCrLf
 Dim i As Integer
 Dim pos As Integer
 Dim posD As Integer
 Dim strLink As String  'eg \vinsoft\vinsplit.lnk  ==> vinsplit.lnk
 Dim strLinkParent As String 'eg \vinsoft
 Dim strLinkFile As String  'vinsplit.exe
 Dim pathUnin As String  'uninstallers path in programs
 pathUnin = "\" & AppName
 For i = 0 To listLink.ListCount - 1
  pos = InStr(1, listLink.List(i), "\")
  posD = InStr(1, listLink.List(i), "#")
  strLink = Trim(Mid(listLink.List(i), pos, posD - pos))
  strLinkFile = Trim(Right(listLink.List(i), Len(listLink.List(i)) - posD))
  pos = InStrRev(strLink, "\")
  strLinkParent = Left(strLink, pos - 1)
  strLink = Right(strLink, Len(strLink) - pos)
  strLink = strLink & ".lnk"

  
 If "D" = Left(listLink.List(i), 1) Then
 
  ShortcutScript = ShortcutScript & "If fSys.FolderExists(strDesktop & " & Chr(34) & strLinkParent & Chr(34) & ") = False Then fSys.CreateFolder strDesktop & " & Chr(34) & strLinkParent & Chr(34) & vbCrLf
  ShortcutScript = ShortcutScript & "Set ShellLink = WShell.CreateShortcut(strDeskTop &" & Chr(34) & strLinkParent & "\" & strLink & Chr(34) & ")" & vbCrLf
  ShortcutScript = ShortcutScript & "ShellLink.TargetPath = Path &" & Chr(34) & "\" & strLinkFile & Chr(34) & vbCrLf
 ElseIf "P" = Left(listLink.List(i), 1) Then
  ShortcutScript = ShortcutScript & "If fSys.FolderExists(strPrograms & " & Chr(34) & strLinkParent & Chr(34) & ") = False Then fSys.CreateFolder strPrograms & " & Chr(34) & strLinkParent & Chr(34) & vbCrLf
  ShortcutScript = ShortcutScript & "Set ShellLink = WShell.CreateShortcut(strPrograms &" & Chr(34) & strLinkParent & "\" & strLink & Chr(34) & ")" & vbCrLf
  ShortcutScript = ShortcutScript & "ShellLink.TargetPath = Path &" & Chr(34) & "\" & strLinkFile & Chr(34) & vbCrLf
  pathUnin = strLinkParent
 ElseIf "S" = Left(listLink.List(i), 1) Then
  ShortcutScript = ShortcutScript & "If fSys.FolderExists(strStartup & " & Chr(34) & strLinkParent & Chr(34) & ") = False Then fSys.CreateFolder strStartup & " & Chr(34) & strLinkParent & Chr(34) & vbCrLf
  ShortcutScript = ShortcutScript & "Set ShellLink = WShell.CreateShortcut(strStartup &" & Chr(34) & strLinkParent & "\" & strLink & Chr(34) & ")" & vbCrLf
  ShortcutScript = ShortcutScript & "ShellLink.TargetPath = Path &" & Chr(34) & "\" & strLinkFile & Chr(34) & vbCrLf
 End If
  ShortcutScript = ShortcutScript & "ShellLink.Save" & vbCrLf
 Next
 ''''LINK TO UNINSTALLER
  ShortcutScript = ShortcutScript & "If fSys.FolderExists(strPrograms & " & Chr(34) & strLinkParent & Chr(34) & ") = False Then fSys.CreateFolder strPrograms & " & Chr(34) & pathUnin & Chr(34) & vbCrLf
  ShortcutScript = ShortcutScript & "Set ShellLink = WShell.CreateShortcut(strPrograms &" & Chr(34) & pathUnin & "\" & "VIN unistaller.lnk" & Chr(34) & ")" & vbCrLf
  ShortcutScript = ShortcutScript & "ShellLink.TargetPath = Path &" & Chr(34) & "\vinunins.exe" & Chr(34) & vbCrLf
  ShortcutScript = ShortcutScript & "ShellLink.Save" & vbCrLf
  CreateShortcutScript = ShortcutScript
End Function

Private Sub txtTarget_Change()
isCompiled = False 'now compile again before building ha haaa hhaaaaaaaa ..
End Sub
