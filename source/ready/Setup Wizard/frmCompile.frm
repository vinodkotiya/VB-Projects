VERSION 5.00
Begin VB.Form frmCompile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compilation Window"
   ClientHeight    =   2235
   ClientLeft      =   2565
   ClientTop       =   7845
   ClientWidth     =   8010
   Icon            =   "frmCompile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox listErr 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   2205
      ItemData        =   "frmCompile.frx":058A
      Left            =   0
      List            =   "frmCompile.frx":0591
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmCompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ***      *** ***   *****  ***   *******    *******
'  ***    ***  ***   *****  ***  ***   ***   ***  ****
'   ***  ***   ***   *** ** ***  ***   ***   ***   ****
'    ******    ***   ***  *****  ***   ***   ***  ****
'     ****     ***   ***   ****   *******    *******
'

Private Sub delay(t As Double)
Dim i As Double
i = Timer()
While Timer() - i < t
DoEvents
Wend
End Sub




Private Sub Form_Load()
frmCompile.Visible = True

End Sub
Private Sub CreatePath(path As String)
''''''''''''''''''''''''''''''''''''''''
'''' This Procedure is written by vinod kotiya
''''' 24-08-2003 - 9:30 pm
''''''''''''''''''''''''''''''''''''''''''''
Dim fsys As New FileSystemObject
Dim tempPath As String
'path = "e:\vin\tin\min\ho\"
Dim pos As Integer
pos = 1
While pos > 0
pos = InStr(pos + 1, path, "\", vbBinaryCompare)
tempPath = Left(path, pos) 'e:\
If fsys.FolderExists(tempPath) = False And tempPath <> "" Then fsys.CreateFolder (tempPath)
Wend

End Sub
Public Function StartBuild() As Boolean
'On Error GoTo vinerror
'On Error Resume Next
Me.Caption = "Compilation Window (Building " & AppName & " )"
listErr.Clear
listErr.ForeColor = vbBlue
Dim fsys As New FileSystemObject
Dim ThisFolder As Folder
Dim AllFolders As Folders 'contain all subfolders of thisfolder
Dim fold As Folder  'i as integer type variable

Dim i As Integer
'create o/p dir

listErr.AddItem "Creating output directory"
'If fsys.FolderExists(OutputDir & "\" & Trim(AppName)) = False Then
 'fsys.CreateFolder (OutputDir & "\" & Trim(AppName))
'End If
 CreatePath (OutputDir & "\" & Trim(AppName) & "\")
 
'create application dir
delay (artificialDelay)
listErr.AddItem "Creating application directory"
'If fsys.FolderExists(OutputDir & "\" & Trim(AppName) & "\data") = False Then fsys.CreateFolder
 CreatePath (OutputDir & "\" & Trim(AppName) & "\data" & "\")
'If fsys.FolderExists(OutputDir & "\" & Trim(AppName) & "\images") = False Then fsys.CreateFolder
 CreatePath (OutputDir & "\" & Trim(AppName) & "\images" & "\")
listErr.AddItem "Copying Appliction files in data directory."
delay (artificialDelay)
'''''' copying setup files and uninstaller
fsys.CopyFile App.path & "\data\setup.exe", OutputDir & "\" & Trim(AppName) & "\"
fsys.CopyFile App.path & "\data\vinunins.exe", OutputDir & "\" & Trim(AppName) & "\data\"
If mdifrmMain.mnuSet(2).Checked Then
fsys.CopyFile App.path & "\data\vinscript.exe", OutputDir & "\" & Trim(AppName) & "\data\"
End If
fsys.CopyFile App.path & "\data\zz.bat", OutputDir & "\" & Trim(AppName) & "\data\"

fsys.CopyFile frmAppl.Dir1.path & "\*.*", OutputDir & "\" & Trim(AppName) & "\data\"
'fsys.CopyFolder frmAppl.Dir1.Path, OutputDir & "\data\"
Set ThisFolder = fsys.GetFolder(frmAppl.Dir1.path)
Set AllFolders = ThisFolder.SubFolders
For Each fold In AllFolders
  fsys.CopyFolder frmAppl.Dir1.path & "\" & fold.Name, OutputDir & "\" & Trim(AppName) & "\data\"
  listErr.AddItem "Copying Folder " & fold.Name
Next

'create system dir
If colDll.Count > 0 Then
 listErr.AddItem "Creating system directory"
 If fsys.FolderExists(OutputDir & "\" & Trim(AppName) & "\system") = False Then fsys.CreateFolder (OutputDir & "\" & Trim(AppName) & "\system")
 listErr.AddItem "Copying system files in system directory."
 For i = 1 To colDll.Count
  fsys.CopyFile colDll.item(i), OutputDir & "\" & Trim(AppName) & "\system\"
 Next
End If

listErr.AddItem "Creating Welcome Screen"
delay (artificialDelay)
Dim fnum As Integer

fnum = FreeFile
 Open OutputDir & "\" & Trim(AppName) & "\VINSetup1.vin" For Output As fnum
 'Width #Fnum, 60
  '///////redundancy
  Print #fnum, "Donot alter anything in this file."
  Print #fnum, "This File contains information about VIN Setup Wizard."
  Print #fnum, "Author: VINOD KOTIYA."
  Print #fnum, "Address: S-2 shrimaya apartment sector-b/363"
  Print #fnum, "        Sarvdharm colony bhopal-42 india ."
  Print #fnum, "Fone : +91-0755-2794428  ."
  Print #fnum, "Email : vinodkotiya24@rediffmail.com"
  Print #fnum, "Web : http:\\vinodkotiya.tripod.com"
  Print #fnum, "Student: B.E. 3rd year "
  Print #fnum, "        Information Technology"
  Print #fnum, "   University Institute of Technology"
  Print #fnum, "   Rajeev Gandhi Produogiki Vishwavidyalaya, BHOPAL"
  Print #fnum, "***************************************"
  '///////redundancy
 Print #fnum, "<WELLCOME> "
 Print #fnum, AppName
 Print #fnum, Version
 Print #fnum, Company
 Print #fnum, WelMessage 'true for image else false
 Print #fnum, frmStart.txtMsg(0).Text
 Print #fnum, frmStart.chkBack(0).Value  'true for tranceparency
 Print #fnum, frmStart.lblCol(0).BackColor
 Print #fnum, frmStart.lblCol(1).BackColor
 Print #fnum, WelImage
 Print #fnum, Trim(frmStart.txtTime.Text)
 
 Print #fnum, DispMessage 'true for image else false
 Print #fnum, frmStart.txtMsg(2).Text
 Print #fnum, frmStart.chkBack(1).Value  'true for tranceparency
 Print #fnum, frmStart.lblCol(2).BackColor
 Print #fnum, frmStart.lblCol(3).BackColor
 Print #fnum, DispImage
 '''''''''''' frmstart complete ..............
 listErr.AddItem "Creating Lisence and agreement files."
delay (artificialDelay)

 Print #fnum, "</WELLCOME> "
 Print #fnum, "<AGREEMENT>"
 Print #fnum, "<Software>"
 If Len(Trim(frmAgree.txtAgree(0).Text)) < 10 Then
  Print #fnum, "There is no software information " & vbCrLf & "For this Application"
 Else
  Print #fnum, frmAgree.txtAgree(0).Text
 End If
 Print #fnum, "</Software>"
 Print #fnum, "<Lisence>"
 If Len(Trim(frmAgree.txtAgree(1).Text)) < 10 Then
  Print #fnum, "There is no Lisence Agreement " & vbCrLf & "For this Application"
 Else
  Print #fnum, frmAgree.txtAgree(1).Text
 End If
 Print #fnum, "</Lisence>"
 Print #fnum, "<ReadME>"
 If Len(Trim(frmAgree.txtAgree(2).Text)) < 10 Then
  Print #fnum, "There is no Read Me File " & vbCrLf & "For this Application"
 Else
  Print #fnum, frmAgree.txtAgree(2).Text
 End If
 Print #fnum, "</ReadMe>"
 Print #fnum, "</AGREEMENT>"
 ''''''''''''''' frmagree complete
 listErr.AddItem "Creating Apllication Information."
delay (artificialDelay)

Dim ret As String
 Print #fnum, "<APPLICATION>"
  Print #fnum, InstallDir
  Print #fnum, "<SystemFiles>"
  Print #fnum, SystemFiles
  Print #fnum, "</SystemFiles>"
  Print #fnum, "<Shortcutvbscript>"
   ret = frmAppl.CreateShortcutScript
  Print #fnum, ret
  Print #fnum, "</Shortcutvbscript>"
  Print #fnum, "</APPLICATION>"
 '''''''''''''''' End of frmAppl
 listErr.AddItem "Creating System Information."
delay (artificialDelay)

 Print #fnum, "<SYSTEM>"
 Print #fnum, SysAd 'true for display user name
 Print #fnum, SysCompany
 Print #fnum, RegCode
 Print #fnum, frmSys.txtReg(0).Text & frmSys.txtReg(1).Text & frmSys.txtReg(2).Text & frmSys.txtReg(3).Text
 Print #fnum, "</SYSTEM>"
 '''''''''''''' End of frmsys
 listErr.AddItem "Creating Finishing Screen."
delay (artificialDelay)

 Print #fnum, "<END>"
 Print #fnum, frmEnd.chkSys(0).Value 'prompt for readme
 Print #fnum, frmEnd.optChk(0).Value 'true for checked else false
 Print #fnum, frmEnd.chkSys(1).Value 'prompt for launch
 Print #fnum, frmEnd.optChk(2).Value 'true for checked else false
 Print #fnum, frmEnd.txtTarget.Text 'launch application
 Print #fnum, frmEnd.chkSys(2).Value 'prompt for boot
 Print #fnum, frmEnd.chkSys(3).Value 'run in back if 1
 Print #fnum, frmEnd.txtRunBack.Text 'launch application in background
 
 Print #fnum, EndMessage 'true for image else false
 Print #fnum, frmEnd.txtMsg(0).Text
 Print #fnum, frmEnd.chkBack.Value  'true for tranceparency
 Print #fnum, frmEnd.lblCol(0).BackColor
 Print #fnum, frmEnd.lblCol(1).BackColor
 Print #fnum, EndImage
 Print #fnum, Trim(frmEnd.txtTime.Text)
 Print #fnum, "</END>"
 '///////redundancy
  Print #fnum, "Donot alter anything in this file."
  Print #fnum, "This File contains information about VIN Setup Wizard."
  Print #fnum, "Author: VINOD KOTIYA."
  Print #fnum, "Address: S-2 shrimaya apartment sector-b/363"
  Print #fnum, "        Sarvdharm colony bhopal-42 india ."
  Print #fnum, "Fone : +91-0755-2794428  ."
  Print #fnum, "Email : vinodkotiya24@rediffmail.com"
  Print #fnum, "Web : http:\\vinodkotiya.tripod.com"
  Print #fnum, "Student: B.E. 3rd year "
  Print #fnum, "        Information Technology"
  Print #fnum, "   University Institute of Technology"
  Print #fnum, "   Rajeev Gandhi Produogiki Vishwavidyalaya, BHOPAL"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  Print #fnum, "***************************************"
  '///////redundancy
 Close #fnum, fnum
 listErr.AddItem "Copying Neccessary Files."
 delay (artificialDelay)
If frmStart.optMsg(1).Value Then fsys.CopyFile Trim(frmStart.txtMsg(1).Text), OutputDir & "\" & Trim(AppName) & "\images\", True
If frmStart.optMsg(3).Value Then fsys.CopyFile Trim(frmStart.txtMsg(3).Text), OutputDir & "\" & Trim(AppName) & "\images\", True
If frmEnd.optMsg(1).Value Then fsys.CopyFile Trim(frmEnd.txtMsg(1).Text), OutputDir & "\" & Trim(AppName) & "\images\", True


If mdifrmMain.mnuCopy(0).Checked Then
 listErr.AddItem "Checking and Creating Software Information."
 delay (artificialDelay)

 fnum = FreeFile
 Open OutputDir & "\" & Trim(AppName) & "\softinfo.txt" For Output As fnum
 Print #fnum, frmAgree.txtAgree(0).Text
 Close #fnum, fnum
End If
If mdifrmMain.mnuCopy(1).Checked Then
 listErr.AddItem "Checking and Creating License Agreement."
 delay (artificialDelay)
 fnum = FreeFile
 Open OutputDir & "\" & Trim(AppName) & "\EULA.txt" For Output As fnum
 Print #fnum, frmAgree.txtAgree(1).Text
 Close #fnum, fnum
End If
If mdifrmMain.mnuCopy(0).Checked Then
 listErr.AddItem "Checking and Creating readme.txt."
 delay (artificialDelay)
 fnum = FreeFile
 Open OutputDir & "\" & Trim(AppName) & "\readme.txt" For Output As fnum
 Print #fnum, frmAgree.txtAgree(2).Text
 Close #fnum, fnum
End If

delay (artificialDelay)
listErr.Clear
listErr.AddItem "         CONGRATULATIONS!!!!!!"
listErr.AddItem "Your Setup has been Succsessfully Created..."
listErr.AddItem "You can get it from Directory " & OutputDir & "\" & AppName
listErr.AddItem "Or execute it by RUN Command <Ctrl+R>"
Me.Caption = "Compilation Window"
StartBuild = True



End Function

Public Function StartCompilation() As Boolean
On Error Resume Next

Me.Caption = "Compilation Window (Compiling " & AppName & " )"
If isCompiled Then
 listErr.Clear
 listErr.AddItem "Already Compiled......."
  Exit Function 'if already compiled then why compile again
End If
Dim iserr As Boolean ' tue if error
Dim fsys As New FileSystemObject
Dim i As Integer
Dim comErr As New Collection
Dim comWar As New Collection
Dim pos As Integer 'find pos of '\' from reverse
'listErr.RemoveItem 0
listErr.Clear
listErr.ForeColor = vbBlue
listErr.AddItem "Checking Application Name............"
If Trim(frmStart.txtInfo(0).Text) = "" Then
 comErr.Add "Error1 : Application Name was not entered in Welcome Screen."
 iserr = True
Else
 AppName = frmStart.txtInfo(0).Text
End If

listErr.AddItem "Checking Application Version..........."
If Trim(frmStart.txtInfo(1).Text) = "" Then
 comWar.Add "Warning1 : Application Version was not entered in Welcome Screen."
Else
 Version = frmStart.txtInfo(1).Text
End If
delay (artificialDelay)
listErr.AddItem "Checking Company Name..........."
If Trim(frmStart.txtInfo(2).Text) = "" Then
 comWar.Add "Warning1 : Company Name was not entered in Welcome Screen."
 Else
 Company = frmStart.txtInfo(2).Text
End If
delay (artificialDelay)
listErr.AddItem "Checking Welcome Message..........."
If frmStart.optMsg(0).Value Then
 If Trim(frmStart.txtMsg(0).Text) = "" Then
 comErr.Add "Error1 : Welcome Message was not entered in Welcome Screen."
 iserr = True
 End If
Else
 If Trim(frmStart.txtMsg(1).Text) = "" Then
 comErr.Add "Error1 : Image File was not Selected for Welcome Screen."
 iserr = True
 
 ElseIf fsys.FileExists(Trim(frmStart.txtMsg(1).Text)) = False Then
 comErr.Add "Error1 : Image File for Welcome Screen is not Exist."
 iserr = True
 Else
 pos = InStrRev(Trim(frmStart.txtMsg(1).Text), "\")
 WelImage = Right(Trim(frmStart.txtMsg(1).Text), Len(Trim(frmStart.txtMsg(1).Text)) - pos)
 End If
 
End If

delay (artificialDelay)
listErr.AddItem "Checking Time to Display............"
If Trim(frmStart.txtTime.Text) = "" Then
  comErr.Add "Error1 : Time Field should not blank."
  iserr = True
 End If
 If IsNumeric(frmStart.txtTime.Text) = False Then
  comErr.Add "Error1 : Please Enter a numeric value for time."
  iserr = True
 End If
 delay (artificialDelay)
listErr.AddItem "Checking Left Side Display............"
If frmStart.optMsg(2).Value Then
 If Trim(frmStart.txtMsg(2).Text) = "" Then
 comErr.Add "Error1 : Welcome Message was not entered in Left Side Display."
 iserr = True
 End If
Else
 If Trim(frmStart.txtMsg(3).Text) = "" Then
  comErr.Add "Error1 : Image File was not Selected for Left Side Display."
  iserr = True
 ElseIf fsys.FileExists(Trim(frmStart.txtMsg(3).Text)) = False Then
  comErr.Add "Error1 : Image File for Left Side Display is not Exist."
  iserr = True
 Else
 pos = InStrRev(Trim(frmStart.txtMsg(3).Text), "\")
 DispImage = Right(Trim(frmStart.txtMsg(3).Text), Len(Trim(frmStart.txtMsg(3).Text)) - pos)
 End If
End If
 delay (artificialDelay)
 
 
listErr.AddItem "Checking License and Agreements............"
If frmAgree.chkAgree(0).Value And Trim(frmAgree.txtAgree(0).Text) = "" Then
  comErr.Add "Error2 : Field of Software Information is empty."
  iserr = True
End If
listErr.AddItem "Checking License Agreements............"
If frmAgree.chkAgree(1).Value And Trim(frmAgree.txtAgree(1).Text) = "" Then
  comErr.Add "Error2 : Field of Lisence Agreement is empty."
  iserr = True
End If
listErr.AddItem "Checking Read Me............"
If frmAgree.chkAgree(2).Value And Trim(frmAgree.txtAgree(2).Text) = "" Then
  comErr.Add "Error2 : Field of Read ME is empty."
  iserr = True
End If


delay (artificialDelay)
listErr.AddItem "Checking Application Information............"
If Right(Trim(frmAppl.txtTarget.Text), 1) = "\" Then
 comWar.Add "Warning3 : A backslash is found at end in installation directory name.  Autocorrected..."
 frmAppl.txtTarget.Text = Left(Trim(frmAppl.txtTarget.Text), Len(Trim(frmAppl.txtTarget.Text)) - 1)
End If

If Trim(frmAppl.txtTarget.Text) = "" Then
 InstallDir = "C:\Program Files\VINSOFT"
 comWar.Add "Warning3 : There is no Installation directory for you application."
Else
 InstallDir = Trim(frmAppl.txtTarget.Text)
End If
delay (artificialDelay)
listErr.AddItem "Checking system files............"
SystemFiles = ""
 For i = 1 To colDll.Count
  If fsys.FileExists(colDll.item(i)) = False Then
    comErr.Add "Error3 : System File " & colDll.item(i) & " does not exist."
    iserr = True
  Else
   SystemFiles = SystemFiles & vbCrLf & frmAppl.listDll.List(i - 1)
  End If
 Next
delay (artificialDelay)

listErr.AddItem "Checking System Information............"
If frmSys.chkSys(2).Value Then
 If 4 > Len(frmSys.txtReg(0).Text) Or 4 > Len(frmSys.txtReg(1).Text) Or 4 > Len(frmSys.txtReg(2).Text) Or 4 > Len(frmSys.txtReg(3).Text) Then
  comErr.Add "Error4 : Registration Code is wrong."
  iserr = True
 End If
End If

delay (artificialDelay)
listErr.AddItem "Checking Application Launch............"
If frmEnd.chkSys(1).Value Then
 If Trim(frmEnd.txtTarget.Text) = "" Then
 comErr.Add "Error5 : Launching Application was not entered."
 iserr = True
 End If
 If fsys.FileExists(frmAppl.Dir1.path & "\" & Trim(frmEnd.txtTarget.Text)) = False Then
 comErr.Add "Error5 : Launching Application " & Trim(frmEnd.txtTarget.Text) & " is not exist."
 iserr = True
 End If
End If
delay (artificialDelay)
If frmEnd.chkSys(3).Value Then
 If Trim(frmEnd.txtRunBack.Text) = "" Then
 comErr.Add "Error5 : Launching Application was not entered."
 iserr = True
 ElseIf fsys.FileExists(frmAppl.Dir1.path & "\" & Trim(frmEnd.txtRunBack.Text)) = False Then
 comErr.Add "Error5 : Launching Application " & Trim(frmEnd.txtRunBack.Text) & " is not exist."
 iserr = True
 End If
End If
delay (artificialDelay)

listErr.AddItem "Checking Finishing Screen............"
If frmEnd.optMsg(0).Value Then
 If Trim(frmEnd.txtMsg(0).Text) = "" Then
 comErr.Add "Error5 : Finishing Message was not entered."
 iserr = True
 End If
Else
 If Trim(frmEnd.txtMsg(1).Text) = "" Then
 comErr.Add "Error5 : Image File was not Selected for Finishing Screen"
 iserr = True
 End If
 If fsys.FileExists(Trim(frmEnd.txtMsg(1).Text)) = False Then
 comErr.Add "Error5 : Image File for Finishing Screen is not Exist."
 iserr = True
 Else
 pos = InStrRev(Trim(frmEnd.txtMsg(1).Text), "\")
 EndImage = Right(Trim(frmEnd.txtMsg(1).Text), Len(Trim(frmEnd.txtMsg(1).Text)) - pos)
 End If
End If
delay (artificialDelay)
listErr.AddItem "Checking Time to Display............"
If Trim(frmEnd.txtTime.Text) = "" Then
  comErr.Add "Error5 : Time Field should not blank."
  iserr = True
 End If
 If IsNumeric(frmEnd.txtTime.Text) = False Then
  comErr.Add "Error5 : Please Enter a numeric value for time."
  iserr = True
 End If

delay (artificialDelay)
listErr.AddItem "Checking Output Folder............"
If Right(Trim(frmEnd.txtOutput.Text), 1) = "\" Then
 comWar.Add "Warning5 : A backslash is found at end in output directory name. Autocorrected..."
 frmEnd.txtOutput.Text = Left(Trim(frmEnd.txtOutput.Text), Len(Trim(frmEnd.txtOutput.Text)) - 1)
End If
If Trim(frmEnd.txtOutput.Text) = "" Then
  comWar.Add "Warning5 : The Output Directory for the Setup Not Entered."
  OutputDir = "C:\VIN Setups"
  frmEnd.txtOutput.Text = OutputDir
Else
OutputDir = Trim(frmEnd.txtOutput.Text)
End If
 '----------------- Checking Complete ----------------
 '----------------- Now Display Error and Warning -----
 
 

 listErr.Clear
 listErr.ForeColor = vbBlue
 listErr.AddItem "Compilation Complete ...........", 0
 delay (artificialDelay)
 If iserr = True Then listErr.AddItem "Click on error to debugg .."
 listErr.AddItem "Warnings : " & comWar.Count & " Errors : " & comErr.Count, 1
 delay (artificialDelay)
 
 For i = 1 To comErr.Count
 listErr.ForeColor = vbRed
 listErr.AddItem comErr.item(i)
 delay (artificialDelay)
 Next
 For i = 1 To comWar.Count
 listErr.ForeColor = vbRed
 listErr.AddItem comWar.item(i)
 delay (artificialDelay)
 Next
 StartCompilation = iserr
 isCompiled = True


End Function
Private Sub listErr_Click()
On Error Resume Next
Dim i As Integer
'MsgBox listLink.Text
For i = 0 To listErr.ListCount - 1
 If i <> listErr.ListIndex Then
  listErr.Selected(i) = False
 End If
Next
Dim Index As Integer
'initialize
If frmStart.Visible = True Then Index = 0
If frmSys.Visible = True Then Index = 3
If frmEnd.Visible = True Then Index = 4
If frmAppl.Visible = True Then Index = 2
If frmAgree.Visible = True Then Index = 1

If Left(listErr.Text, 6) = "Error1" Or Left(listErr.Text, 8) = "Warning1" Then
 Index = 0
ElseIf Left(listErr.Text, 6) = "Error2" Or Left(listErr.Text, 8) = "Warning2" Then
 Index = 1
ElseIf Left(listErr.Text, 6) = "Error3" Or Left(listErr.Text, 8) = "Warning3" Then
 Index = 2
ElseIf Left(listErr.Text, 6) = "Error4" Or Left(listErr.Text, 8) = "Warning4" Then
 Index = 3
ElseIf Left(listErr.Text, 6) = "Error5" Or Left(listErr.Text, 8) = "Warning5" Then
 Index = 4
End If

frmButton.imgStepOver_Click (Index)

End Sub
