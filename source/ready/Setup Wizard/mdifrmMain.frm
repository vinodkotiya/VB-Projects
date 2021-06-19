VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdifrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "VIN Setup Wizard"
   ClientHeight    =   5115
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7785
   Icon            =   "mdifrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu project 
      Caption         =   "&Project"
      Begin VB.Menu mnuNew 
         Caption         =   "New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Project"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Projecct As......"
         Index           =   1
      End
      Begin VB.Menu er 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu setting 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSet 
         Caption         =   "Artificial Delay"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSet 
         Caption         =   "WallPaper"
         Index           =   1
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Export wscript.exe with Setup"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSet 
         Caption         =   "In Output Dir Create a Copy of"
         Index           =   3
         Begin VB.Menu mnuCopy 
            Caption         =   "Software Information"
            Checked         =   -1  'True
            Index           =   0
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "License Agreement"
            Checked         =   -1  'True
            Index           =   1
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "Readme.txt"
            Checked         =   -1  'True
            Index           =   2
            Shortcut        =   ^M
         End
      End
   End
   Begin VB.Menu wizard 
      Caption         =   "&Wizard"
      Begin VB.Menu mnuWiz 
         Caption         =   "Welcome Screen"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuWiz 
         Caption         =   "License And Agreement"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuWiz 
         Caption         =   "Application Information"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuWiz 
         Caption         =   "System Information"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuWiz 
         Caption         =   "Finishing Setup"
         Index           =   4
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu compile 
      Caption         =   "&Compile"
      Begin VB.Menu mnuCompile 
         Caption         =   "Compile"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuBuild 
         Caption         =   "Buid Setup"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Run Setup"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuRunfull 
         Caption         =   "Run with Full Compile"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "About"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "About Me"
         Index           =   2
      End
   End
End
Attribute VB_Name = "mdifrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***      *** ***   *****  ***   *******    *******
'  ***    ***  ***   *****  ***  ***   ***   ***  ****
'   ***  ***   ***   *** ** ***  ***   ***   ***   ****
'    ******    ***   ***  *****  ***   ***   ***  ****
'     ****     ***   ***   ****   *******    *******
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Programmer : VINOD KOTIYA
'  B.E. (Information Technology)
'  Semester V
'  University Institute of Technology
'  Rajeev Gandhi Prodyogiki Vishwavidyalaya Bhopal.
'  Address: S-2 ShreeMaya Apartment Sector-B/363
'           Sarvdharm Colony Bhopal-42 (India)
'  Email: vinodkotiya24@rediffmail.com
'  Web : http://vinodkotiya.tripod.com
'  cell: +91-9827394994
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Date of Starting:Wednesday,06,Aug 2003, 10:59:20 AM
'  Completion Date :Monday,05,Aug 2003, 11:43:33 AM
'  Associated Projects:1. Main Installer 2. VIN Uninstaller
'
'  First Modification : 10-aug-2003
'                       Debugging feature in compilation
'     window was added.
'  Second Modification : 15-aug-2003
'                       Settings option added.
'  Third Modification : 24-aug-2003
'                       Path creation bug was fixed by
'     CreatePath algorithm.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim OpenProjectName As String
Dim isCompiling As Boolean 'true when compiling 'if already cumoiling then dont enter there
Public iserr As Boolean 'true if error in compile
Dim frmTop As Integer 'store top value of all forms
Dim frmLeft As Integer 'store left value of all forms
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_OTHERUNZOOM = 4



Private Function MakeFile() As String
Dim txtRet As String
txtRet = "<<<<VINOD KOTIYA's VIN Setup Wizard>>>>" & vbCrLf
txtRet = txtRet & frmStart.MakeFile1
txtRet = txtRet & frmAgree.MakeFile2
txtRet = txtRet & frmAppl.MakeFile3
txtRet = txtRet & frmSys.MakeFile4
txtRet = txtRet & frmEnd.MakeFile5
MakeFile = txtRet
End Function

Private Sub MDIForm_Click()
frmPrev.Visible = False
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
Dim i As Integer
i = Second(Time)
If (i > 15 And i < 30) Or (i > 45 And i < 60) Then
 mdifrmMain.Picture = LoadPicture(App.path & "\data\wallpaper\drievin.jpg")
Else
 mdifrmMain.Picture = LoadPicture(App.path & "\data\wallpaper\fishvin.jpg")
End If
If Screen.Width = 12000 Then
  frmLeft = 1800
  frmTop = 700  '// 800x600
  frmPrev.Top = frmTop
  frmPrev.Left = frmLeft
  frmButton.Left = frmLeft - 500
  mdifrmMain.Caption = "VIN Setup Wizard v1.0 (Please set your Monitors resolution to 1024 X 768)"
Else
frmLeft = 3000
frmTop = 1000

End If

frmStart.Top = frmTop
frmSys.Top = frmTop
frmEnd.Top = frmTop
frmAppl.Top = frmTop
frmAgree.Top = frmTop
frmStart.Left = frmLeft
frmSys.Left = frmLeft
frmEnd.Left = frmLeft
frmAppl.Left = frmLeft
frmAgree.Left = frmLeft
frmCredit.Left = frmLeft

Load frmStart
Load frmAgree
Load frmAppl
Load frmSys
Load frmEnd
Load frmAppl
Load frmButton
'frmCredit.Hide
frmButton.Top = mdifrmMain.Top + 70
frmStart.Visible = True
frmSys.Visible = False
frmEnd.Visible = False
frmAppl.Visible = False
frmAgree.Visible = False
frmPrev.Visible = False
frmButton.Visible = True
frmStart.txtMsg(1).Text = App.path & "\data\images\wel.gif"
frmStart.txtMsg(3).Text = App.path & "\data\images\left.jpg"
artificialDelay = 0.5
mnuCopy(0).Checked = True
mnuCopy(1).Checked = True
mnuCopy(2).Checked = True
mnuSet(2).Checked = False
frmButton.imgStepOver_Click (0)
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu wizard
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reply As Integer
reply = MsgBox("Do you want to save the VIN Setup Project before closing....", vbYesNoCancel, "Prompt For Save")
If reply = vbYes Then         'yes
  mnuSave_Click (0)
  End
ElseIf reply = vbCancel Then
 Cancel = 1
Else
 End
End If
End Sub

Public Sub mnuBuild_Click()
Dim i As Boolean
mnuCompile_Click
If iserr = False Then
  i = frmCompile.StartBuild
Else
MsgBox "There are some errors.First Eliminate them."
End If
End Sub

Public Sub mnuCompile_Click()
If isCompiling = True Then Exit Sub 'if already cumoiling then dont enter here
If frmCompile.Visible Then Unload frmCompile
 Load frmCompile
 frmCompile.Top = frmAgree.Top + frmAgree.ScaleHeight
 frmCompile.Left = frmStart.Left
 frmCompile.Show
 isCompiling = True
 iserr = frmCompile.StartCompilation
 isCompiling = False
End Sub

Private Sub mnuCopy_Click(Index As Integer)
mnuCopy(Index).Checked = Not mnuCopy(Index).Checked
End Sub

Public Sub mnuHelp_Click(Index As Integer)
If Index = 0 Then
 Dim b As Boolean
b = ShowHelp("setuphelp.chm", True)
ElseIf Index = 1 Then
MsgBox "''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
"  Programmer : VINOD KOTIYA" & vbCrLf & "  B.E. (Information Technology)" & vbCrLf & _
"  Semester V" & vbCrLf & "  University Institute of Technology" & vbCrLf & _
"  Rajeev Gandhi Prodyogiki Vishwavidyalaya Bhopal." & vbCrLf & _
"  Address: S-2 ShreeMaya Apartment Sector-B/363" & vbCrLf & _
"           Sarvdharm Colony Bhopal-42 (India)" & vbCrLf & _
"  Email: vinodkotiya24@rediffmail.com" & vbCrLf & "  Web : http://vinodkotiya.tripod.com" & vbCrLf & _
"'''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
"  Date of Starting:Wednesday,06,Aug 2003, 10:59:20 AM" & vbCrLf & _
"  Completion Date :Monday,05,Aug 2003, 11:43:33 AM" & vbCrLf & _
"  Associated Projects:1. Main Installer 2. VIN Uninstaller" & vbCrLf & _
"  First Modification : 10-aug-2003" & vbCrLf & "         Debugging feature in compilation window was added." & vbCrLf & _
"  Second Modification : 15-aug-2003" & vbCrLf & "        Settings option added." & vbCrLf & "  Third Modification : 24-aug-2003" & vbCrLf & _
"                       Path creation bug was fixed by" & vbCrLf & _
"     CreatePath algorithm." & vbCrLf & _
"''''''''''''''''''''''''''''''''''''''''''''''''''''''''''" & vbCrLf & _
"''''''''''''''''''''''''''''''''''''''''''''''''''''''''''"""
ElseIf Index = 2 Then
Load frmCredit
frmCredit.Visible = True
End If
End Sub
Public Function ShowHelp(strTopic As String, bIsLocal As Boolean) As Boolean
Dim strDir As String
If bIsLocal Then

' Get registry entry pointing to Help
strDir = App.path + "\data\"

End If

' Launch topic
Dim hinst As Long
hinst = ShellExecute(Me.hwnd, vbNullString, strTopic, vbNullString, strDir, SW_SHOWNORMAL)

' Handle less than 32 indicates failure
ShowHelp = hinst > 32

End Function

Public Sub mnuNew_Click()
On Error Resume Next
Unload frmStart
Unload frmAgree
Unload frmAppl
Unload frmSys
Unload frmEnd
Unload frmAppl

Load frmStart
Load frmAgree
Load frmAppl
Load frmSys
Load frmEnd
Load frmAppl
frmStart.Top = frmTop
frmSys.Top = frmTop
frmEnd.Top = frmTop
frmAppl.Top = frmTop
frmAgree.Top = frmTop
frmStart.Left = frmLeft
frmSys.Left = frmLeft
frmEnd.Left = frmLeft
frmAppl.Left = frmLeft
frmAgree.Left = frmLeft
frmStart.Visible = True
frmSys.Visible = False
frmEnd.Visible = False
frmAppl.Visible = False
frmAgree.Visible = False
frmStart.txtMsg(1).Text = App.path & "\data\images\wel.gif"
frmStart.txtMsg(3).Text = App.path & "\data\images\left.jpg"

  frmButton.imgStepOver_Click (0)
''''refreshing list of dll files
Dim i As Integer
For i = colDll.Count To 1 Step -1
    colDll.Remove i
Next
For i = frmAppl.listDll.ListCount - 1 To 0 Step -1
    frmAppl.listDll.RemoveItem (i)
Next
For i = frmAppl.listLink.ListCount - 1 To 0 Step -1
    frmAppl.listLink.RemoveItem (i)
Next
For i = frmSys.listLink.ListCount - 1 To 0 Step -1
    frmSys.listLink.RemoveItem (i)
   Next
End Sub

Public Sub mnuOpen_Click()

On Error GoTo fileerror
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.DefaultExt = "VKP"
    CommonDialog1.Filter = "VIN Setup Projects (*.vkp)|*.vkp"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.fileName = "" Then Exit Sub
    OpenProjectName = CommonDialog1.fileName
    
  frmButton.imgStepOver_Click (0)
  
    ''''refreshing list of dll files
   Dim i As Integer
   For i = colDll.Count To 1 Step -1
     colDll.Remove i
    Next
    For i = frmAppl.listDll.ListCount - 1 To 0 Step -1
    frmAppl.listDll.RemoveItem (i)
    Next
   
   For i = frmAppl.listLink.ListCount - 1 To 0 Step -1
    frmAppl.listLink.RemoveItem (i)
   Next
   For i = frmSys.listLink.ListCount - 1 To 0 Step -1
    frmSys.listLink.RemoveItem (i)
   Next

    OpenProject
    
    
    Exit Sub

fileerror:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & CommonDialog1.fileName
    OpenProjectName = ""

End Sub
Private Sub OpenProject()
Dim fnum As Integer
Dim currentline As String

On Error GoTo fileerror
   
    fnum = FreeFile    'getting file no for futures referance
    Open OpenProjectName For Input As fnum    'dont use #1 for multiple file openings
    'While Not EOF(FNum)
    Line Input #fnum, currentline  '<<<<VINOD KOTIYA's VIN Setup Wizard>>>>
    If currentline <> "<<<<VINOD KOTIYA's VIN Setup Wizard>>>>" Then
      MsgBox "This file is not valid VIN Setup Wizard Project File."
      Close #fnum
    
    Exit Sub
    End If
    Line Input #fnum, currentline  '<<<Starting Form>>>
    Line Input #fnum, currentline  'application name
    frmStart.txtInfo(0).Text = currentline
    Line Input #fnum, currentline  'application ver
    frmStart.txtInfo(1).Text = currentline
    Line Input #fnum, currentline  'compony name
    frmStart.txtInfo(2).Text = currentline
     Line Input #fnum, currentline ' <Welcome>
     Line Input #fnum, currentline 'opmsg(0 or 1)
     frmStart.optMsg(currentline).Value = True
      If currentline = 1 Then frmStart.cmdBrowse(0).Enabled = True
    Line Input #fnum, currentline  'txtmsg(0)
    frmStart.txtMsg(0).Text = currentline
     Line Input #fnum, currentline 'backcol
     frmStart.lblCol(0).BackColor = currentline
     Line Input #fnum, currentline 'chkback(0 or -1) if -1 then skip
      If currentline = "0" Then
        frmStart.chkBack(0).Value = Checked
      Else
        frmStart.chkBack(0).Value = Unchecked
      End If
     Line Input #fnum, currentline 'forecol
     frmStart.lblCol(1).BackColor = currentline
     Line Input #fnum, currentline 'image path
     frmStart.txtMsg(1).Text = currentline
     Line Input #fnum, currentline 'opmsg(2 or 3)
     frmStart.optMsg(currentline).Value = True
     If currentline = 3 Then frmStart.cmdBrowse(1).Enabled = True
    Line Input #fnum, currentline  'txtmsg(2)
    frmStart.txtMsg(2).Text = currentline
     Line Input #fnum, currentline 'backcol
     frmStart.lblCol(2).BackColor = currentline
     Line Input #fnum, currentline 'chkback(1 or -1) if -1 then skip
      If currentline = "1" Then
        frmStart.chkBack(1).Value = Checked
      Else
        frmStart.chkBack(1).Value = Unchecked
      End If
     Line Input #fnum, currentline 'forecol
     frmStart.lblCol(3).BackColor = currentline
     Line Input #fnum, currentline 'image path
     frmStart.txtMsg(3).Text = currentline
     
     '----------------End of frmstart ------------------
     DoEvents
     '----------------frmAgree--------------------------
     Line Input #fnum, currentline '<<<Agreement>>>
     Line Input #fnum, currentline 'chkagree(0 or -1) if -1 then skip
      If currentline = "0" Then frmAgree.chkAgree(0).Value = Checked
     Line Input #fnum, currentline 'chkagree(1 or -1)
      If currentline = "1" Then frmAgree.chkAgree(1).Value = Checked
     Line Input #fnum, currentline 'chkagree(2 or -1)
      If currentline = "2" Then frmAgree.chkAgree(2).Value = Checked
     Line Input #fnum, currentline ' <Software Info>
     Do While Not "</Software Info>" = Trim(currentline)
      Line Input #fnum, currentline
      If "</Software Info>" = Trim(currentline) Then Exit Do
      frmAgree.txtAgree(0).Text = frmAgree.txtAgree(0).Text & currentline & vbCrLf
     Loop
     DoEvents
     Line Input #fnum, currentline   '  <Lisence>
     Do While (Not "</Lisence>" = Trim(currentline)) 'Or (Not EOF(FNum))
      Line Input #fnum, currentline
      If "</Lisence>" = Trim(currentline) Then Exit Do
      frmAgree.txtAgree(1).Text = frmAgree.txtAgree(1).Text & currentline & vbCrLf
     Loop
     DoEvents
     Line Input #fnum, currentline ' <Read Me>
     Do While (Not "</Read Me>" = Trim(currentline)) ' Or (Not EOF(FNum))
      Line Input #fnum, currentline
      If "</Read Me>" = Trim(currentline) Then Exit Do
      frmAgree.txtAgree(2).Text = frmAgree.txtAgree(2).Text & currentline & vbCrLf
     Loop
     '---------------------- End of frmagree -------------
     DoEvents
     '------------------ start of frm appl ---------------
     
     Line Input #fnum, currentline '<<<Application>>>
     Line Input #fnum, currentline 'dir1.path
      frmAppl.Dir1.path = currentline
      ChDrive frmAppl.Dir1.path
      frmAppl.Drive1.Refresh
     Line Input #fnum, currentline
     frmAppl.txtTarget.Text = currentline
     Line Input #fnum, currentline '  <ListLink>
     Do While (Not "</ListLink>" = Trim(currentline)) ' Or (Not EOF(FNum))
      Line Input #fnum, currentline
      If "</ListLink>" = Trim(currentline) Then Exit Do
      If Trim(currentline) <> "" Then frmAppl.listLink.AddItem currentline
     Loop
     Line Input #fnum, currentline '  <ListDll>
     Do While (Not "</ListDll>" = Trim(currentline)) ' Or (Not EOF(FNum))
      Line Input #fnum, currentline
      If "</ListDll>" = Trim(currentline) Then Exit Do
      If Trim(currentline) <> "" Then frmAppl.listDll.AddItem currentline
     Loop
      Line Input #fnum, currentline '  <colDll>
     Do While (Not "</colDll>" = Trim(currentline)) ' Or (Not EOF(FNum))
      Line Input #fnum, currentline
      If "</colDll>" = Trim(currentline) Then Exit Do
      If Trim(currentline) <> "" Then colDll.Add currentline
     Loop
     '---------------------End of frmAppl ---------------------
     DoEvents
     '---------------------Start of frm Sys -------------------
     Line Input #fnum, currentline '<<<System Information>>>
     Line Input #fnum, currentline 'chkSYs(0 or -1) if -1 then skip
      If currentline = "0" Then frmSys.chkSys(0).Value = Checked
     Line Input #fnum, currentline 'chkSYs(1 or -1)
      If currentline = "1" Then frmSys.chkSys(1).Value = Checked
     Line Input #fnum, currentline 'chkSYs(2 or -1)
      If currentline = "2" Then frmSys.chkSys(2).Value = Checked
     Line Input #fnum, currentline
      frmSys.txtReg(0).Text = currentline
     Line Input #fnum, currentline
      frmSys.txtReg(1).Text = currentline
     Line Input #fnum, currentline
      frmSys.txtReg(2).Text = currentline
     Line Input #fnum, currentline
      frmSys.txtReg(3).Text = currentline
     Line Input #fnum, currentline '  <ListLink>
     Do While (Not "</ListLink>" = Trim(currentline)) ' Or (Not EOF(FNum))
      Line Input #fnum, currentline
      If "</ListLink>" = Trim(currentline) Then Exit Do
      If Trim(currentline) <> "" Then frmSys.listLink.AddItem currentline
     Loop
      '-----------------End of frmSys -----------------
      DoEvents
      '----------------Start of frmEnd ----------------
      Line Input #fnum, currentline '<<<Finishing Form>>>
      Line Input #fnum, currentline 'chkSYs(0 or -1) if -1 then skip
      If currentline = "0" Then frmEnd.chkSys(0).Value = Checked
       Line Input #fnum, currentline 'optchk(0 or 1)
        frmEnd.optChk(currentline).Value = True
      Line Input #fnum, currentline 'chkSYs(1 or -1)
      If currentline = "1" Then frmEnd.chkSys(1).Value = Checked
       Line Input #fnum, currentline 'optchk(2 or 3)
        frmEnd.optChk(currentline).Value = True
      Line Input #fnum, currentline  'launch app
       frmEnd.txtTarget.Text = currentline
      Line Input #fnum, currentline 'chkSYs(2 or -1) 'reboot
      If currentline = "2" Then frmEnd.chkSys(2).Value = Checked
       Line Input #fnum, currentline 'chkSYs(3 or -1) 'run in back
      If currentline = "3" Then frmEnd.chkSys(3).Value = Checked
     Line Input #fnum, currentline  'back app
       frmEnd.txtRunBack.Text = currentline
      
      Line Input #fnum, currentline ' <Alvida>
      Line Input #fnum, currentline 'optmsg(0 or 1)
      frmEnd.optMsg(currentline).Value = True
       If currentline = 1 Then frmEnd.cmdBrowse(0).Enabled = True
      Line Input #fnum, currentline 'txtmsg0
      frmEnd.txtMsg(0).Text = currentline
      Line Input #fnum, currentline 'backcolor
      frmEnd.lblCol(0).BackColor = currentline
      Line Input #fnum, currentline 'chkback(0 or -1) if -1 then skip
      If currentline = "0" Then
       frmEnd.chkBack.Value = Checked
      Else
        frmEnd.chkBack.Value = Unchecked
      End If
      Line Input #fnum, currentline 'forcolor
      frmEnd.lblCol(1).BackColor = currentline
      Line Input #fnum, currentline 'txtmsg1
      frmEnd.txtMsg(1).Text = currentline
      Line Input #fnum, currentline  'txtoutput
       frmEnd.txtOutput.Text = currentline
   
      '-------------- End of frmEnd -------------
      DoEvents
    Close #fnum
   Exit Sub
fileerror:
    MsgBox "Unkown error while Loading Project " & OpenProjectName


End Sub

Public Sub mnuRun_Click()
Dim fsys As New FileSystemObject
If OutputDir = "" Then OutputDir = Trim(frmEnd.txtOutput.Text)  'when any saved project is opened and run button is pressed
If AppName = "" Then AppName = frmStart.txtInfo(0).Text
If fsys.FileExists(OutputDir & "\" & Trim(AppName) & "\setup.exe") = False Then
 MsgBox "Setup File is not exist.Please First Compile and Build to Run The Setup."
 Exit Sub
End If
 
Dim hinst As Long
  If Trim(OutputDir) = "" Then OutputDir = frmEnd.txtOutput.Text
  hinst = ShellExecute(Me.hwnd, vbNullString, "setup.exe", vbNullString, OutputDir & "\" & Trim(AppName), SW_SHOWNORMAL And SW_OTHERUNZOOM)
'  MsgBox "yo"

End Sub

Private Sub mnuRunfull_Click()
mnuBuild_Click
mnuRun_Click
End Sub

Public Sub mnuSave_Click(Index As Integer)
Dim fnum As Integer
Dim txtSave As String
On Error GoTo vinerror
    If OpenProjectName = "" Or Index = 1 Then
     CommonDialog1.DefaultExt = "VKP"
     CommonDialog1.Filter = "VIN Setup Projects(*.vkp)|*.vkp"
     CommonDialog1.ShowSave
     If CommonDialog1.fileName = "" Then Exit Sub
      OpenProjectName = CommonDialog1.fileName
    End If
    txtSave = MakeFile
    fnum = FreeFile    'getting file no for futures referance
    If Index = 1 Then
      Open CommonDialog1.fileName For Output As fnum     'dont use #1 for multiple file openings
    Else
      Open OpenProjectName For Output As fnum     'dont use #1 for multiple file openings
    End If
     Print #fnum, txtSave
     Close fnum
    
    mdifrmMain.Caption = "VIN Setup Wizard v1.0  ( " & OpenProjectName & " )"
   Exit Sub
vinerror:
     If Err.Number = cdlCancel Then Exit Sub
    
End Sub

Private Sub mnuSet_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
mnuSet(0).Checked = Not mnuSet(0).Checked
  If mnuSet(0).Checked Then
   artificialDelay = 0.5
  Else
    artificialDelay = 0.1
  End If
ElseIf Index = 1 Then
  CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
  CommonDialog1.InitDir = App.path & "\data\wallpaper"
  CommonDialog1.fileName = ""
  CommonDialog1.Filter = "*.jpg|*.jpg|*.bmp|*.bmp|*.wmf|*.wmf|*.gif|*.gif|All Files|*.*"
  CommonDialog1.ShowOpen
  If CommonDialog1.fileName = "" Or Err.Number = cdlCancel Then
      MsgBox "No Background image file is opened"
      Exit Sub
   End If
  mdifrmMain.Picture = LoadPicture(CommonDialog1.fileName)
ElseIf Index = 2 Then
 mnuSet(2).Checked = Not mnuSet(2).Checked
  If mnuSet(2).Checked Then
   MsgBox "Wscript.exe is used only to run vbscript to create " & _
   " Shortcuts at installation time. Probability that end user have wscript.exe " & _
   " is higher. So you can save 150 kb if not exporting."
  End If
End If
End Sub

Public Sub mnuWiz_Click(Index As Integer)
If Index = 0 Then frmStart.Visible = True
If Index = 3 Then frmSys.Visible = True
If Index = 4 Then frmEnd.Visible = True
If Index = 2 Then frmAppl.Visible = True
If Index = 1 Then frmAgree.Visible = True
If Index <> 0 Then frmStart.Visible = False
If Index <> 3 Then frmSys.Visible = False
If Index <> 4 Then frmEnd.Visible = False
If Index <> 2 Then frmAppl.Visible = False
If Index <> 1 Then frmAgree.Visible = False

End Sub


