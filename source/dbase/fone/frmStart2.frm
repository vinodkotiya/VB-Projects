VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fone Directory"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7335
   Icon            =   "frmStart2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "frmStart2.frx":1CCA
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      Height          =   495
      Index           =   2
      Left            =   6960
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modify"
      Height          =   495
      Index           =   1
      Left            =   1800
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ComboBox txtSearch 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   $"frmStart2.frx":4F8E
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4080
      Top             =   1920
   End
   Begin VB.ComboBox cmbFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmStart2.frx":5022
      Left            =   1320
      List            =   "frmStart2.frx":5024
      TabIndex        =   0
      ToolTipText     =   "Select the field in which you want to search."
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start Searching"
      Height          =   735
      Index           =   0
      Left            =   5040
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Search In"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "To Search"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Menu mnuTask 
      Caption         =   "&Task"
      Begin VB.Menu mnuStart 
         Caption         =   "Start Search"
         Shortcut        =   ^S
      End
      Begin VB.Menu submnuExtend 
         Caption         =   "Extend the Search In Options"
         Shortcut        =   ^E
      End
      Begin VB.Menu display 
         Caption         =   "Display the whole Fone Directory"
         Shortcut        =   ^D
      End
      Begin VB.Menu ree 
         Caption         =   "-"
      End
      Begin VB.Menu submnuClear 
         Caption         =   "Clear Previous Search List"
         Shortcut        =   ^C
      End
      Begin VB.Menu submnuModify 
         Caption         =   "Add/Delete Records"
         Shortcut        =   ^N
      End
      Begin VB.Menu restore 
         Caption         =   "Restore Directory to Last Use"
         Shortcut        =   ^R
      End
      Begin VB.Menu ww1 
         Caption         =   "-"
      End
      Begin VB.Menu vinrem 
         Caption         =   "VIN Remind Me Later"
         Shortcut        =   ^V
      End
      Begin VB.Menu ww 
         Caption         =   "-"
      End
      Begin VB.Menu submnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&Help"
      Begin VB.Menu submnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu submnuCredit 
         Caption         =   "Credit"
         Shortcut        =   {F2}
      End
      Begin VB.Menu submnuAbout 
         Caption         =   "About"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuApp 
      Caption         =   "&VIN UTILITY KIT"
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
'///////////////// VIN FONE DIRECTORY  ///////////////////////////////////
'//////////        Created By : - VINOD KOTIYA             ///////////////////////////
'/////////          free on http://vinodkotiya.tripod.com ////////////////////
'/////////          help to promote the site if you want     ///////////////////////////////
'/////////          tell to your friend.         ///////////////////////////////////////
'/////////          provide advertizer or online job ///////
'/////////          Proudly releasing version 1.0 ////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////

'/////////////////   use project FONE 2 for latest coding

Option Explicit
Dim OpenFile As String
Dim txt As String
Dim madhya As Integer     'contain splitted scroll
Dim unloadmdi As Boolean
Dim scroll As String    'used scrolling text
Dim VIN As Byte

Private Sub Command_Click(Index As Integer)

If Index = 0 Then
 '//unloadmdi form before next searching
If txtSearch.Text = "" Then
  MsgBox "Please  Enter what you wanna search in 'To Search' "
  Exit Sub
End If
If unloadmdi = True Then      ' unload after first clicking over txtsearch
   Unload frmShow
   Unload MDIForm1
End If
unloadmdi = True
'/// now load
Load MDIForm1
MDIForm1.Left = Screen.Width - MDIForm1.Width
Load frmShow
frmShow.Visible = True
'///////ADDING TO LIST  ////////////////////////
Dim j As Integer
'scan whole txtsearchlist to prevent duplicasy of new entery
For j = 0 To txtSearch.ListCount
   If txtSearch.List(j) = txtSearch.Text Then
        Exit Sub   'item already exist so quit
   End If
Next                 'item not exist so add it
    If Trim(txtSearch.Text) <> "" Then
        txtSearch.AddItem txtSearch.Text
    End If
'/////////////////////////////////////////////

ElseIf Index = 1 Then
 Load Form1
 Form1.Visible = True
ElseIf Index = 2 Then
 submnuHelp_Click
End If
End Sub








Private Sub Command_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Command(Index).BackColor = &HF0E7D7
End Sub

Private Sub display_Click()
txtSearch.Text = "*"
Command_Click (0)
End Sub

Private Sub Form_Load()
On Error Resume Next
Form1.Datafone.DatabaseName = App.Path & "\data\DIRECTORY.mdb"
Form1.Datafone.RecordSource = "Fone"
Form1.Datafone.Refresh
'////////////////////////////
'''''''for splash /
Dim displayTime As Integer



'MsgBox Environ("OS")
If Environ("OS") = "Windows_NT" Then
displayTime = 1000   'in milliseconds
VIN = 0
Timer1.Interval = Int(displayTime / 255)
'''if os is not nt then do something else
Else
' MsgBox "Not windows nt"
VIN = 255
Me.Show

End If
 
 




'\////////////////////////
Form1.Datafone.DatabaseName = App.Path & "\data\DIRECTORY.mdb"
BackupDbase        'BACKUP THE FONE DIRECTORY DATABASE ON STARTUPS
    cmbFields.AddItem "Name"
      cmbFields.AddItem "Surname"
      cmbFields.AddItem "Area"
      cmbFields.AddItem "City"
      cmbFields.AddItem "FoneResident"
      cmbFields.AddItem "FoneOffice"
      cmbFields.AddItem "Mobile"
      
     cmbFields.ListIndex = 0
'load textsearchbox
     loadtxtsearch
Dim parts() As String
Dim i As Integer

   parts = Split(txt, "^")
 'split txt and save it to arry
 'then add to serchbox
  For i = 1 To UBound(parts)
    
        txtSearch.AddItem parts(i - 1)
    
  Next
'initializing global variables
 madhya = 1
 scroll = "       Fone Directory By Vinod Kotiya     *********"
 unloadmdi = False      'dont unload on first clicking over txtsearch
 
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuTask
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
Command(1).BackColor = vbWhite
Command(0).BackColor = vbWhite
Command(2).BackColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
 savethesearchtext
Unload Form1
Unload Form3
Unload frmShow
'Unload MDIForm1


Unload frmHelp
Unload Me

End Sub

Private Sub loadtxtsearch()


Dim FNum As Integer



On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\fonesearch.vin" For Input As #1
    txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & "fonesearch.vin" _
     & "To eliminate problem Open Your TextEditor(NotePad) " _
     & "Click on File->Save and Save the file as 'fonesearch.vin' in FoneDirectory\data\vinod.vin "
     
    OpenFile = ""
    
End Sub


Private Sub savethesearchtext()
Dim FNum As Integer
Dim txt As String

Dim i As Integer
On Error GoTo FileError
    FNum = FreeFile
    Open App.Path & "\data\fonesearch.vin" For Output As #1
     For i = txtSearch.ListCount - 1 To 0 Step -1
       txt = txt & txtSearch.List(i) & "^"
      
      Next
      Print #FNum, txt
      Print #FNum, "fone directory By Vinod Kotiya" _
                & "save your searches to hard disk"
    Close #FNum
    'OpenFile = "c:\vin.vin" 'CommonDialog1.FileName
    Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file " & "fonesearch.vin" 'CommonDialog1.FileName
    OpenFile = ""
End Sub


Private Sub subMenuPicker_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "indradhanush.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'indradhanush' is not found in its " _
  & "Default directory fonedirectory\data\vinclock.exe "

Exit Sub
End Sub

Private Sub mnuApp_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell("vin_utility.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'vin_utility.EXE' is not found in its " _
  & "Default directory "
End Sub

Private Sub mnuStart_Click()
Command_Click (0)
End Sub

Private Sub restore_Click()
Dim Fsys As New FileSystemObject
Dim reply As Integer
On Error GoTo vinerror
If Fsys.FolderExists("c:\windows\vinbakup") = False Then  'if folder not exist create it
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
 MsgBox "Sorry| The Restoration may fail"
End Sub
Private Sub BackupDbase()
Dim Fsys As New FileSystemObject
Dim reply As Integer
On Error GoTo vinerror
If Fsys.FolderExists("c:\windows\vinbakup") = False Then Exit Sub  'if folder not exist

Fsys.CopyFile App.Path & "\data\directory.mdb", "c:\windows\vinbakup\", True

Exit Sub
vinerror:

End Sub
Private Sub submnuAbout_Click()
MsgBox ("Fone Directory is dedicated to Nishikant Naveen for his valuable support.")
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\about.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'about.EXE' is not found in its " _
  & "Default directory about.exe "
End Sub

Private Sub submnuClear_Click()
'deleting  previous search LIST
Dim i As Integer
 For i = txtSearch.ListCount - 1 To 0 Step -1
       txtSearch.RemoveItem (i)
 Next
    
End Sub


Private Sub submnuCredit_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\credit.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'CREDIT.EXE' is not found in its " _
  & "Default directory CREDIT.exe "
End Sub

Private Sub submnuExit_Click()
Unload Me
End Sub

Private Sub submnuExtend_Click()
Dim i As Integer
submnuExtend.Checked = Not submnuExtend.Checked 'toggle checked
 If submnuExtend.Checked = False Then
     For i = cmbFields.ListCount - 1 To 7 Step -1
         cmbFields.RemoveItem (i) ''when unchecked remove
     Next
     
 Else       ''when checked add more
       cmbFields.AddItem "Address"
      cmbFields.AddItem "Designation"
      cmbFields.AddItem "Email"
      cmbFields.AddItem "Website"
      cmbFields.AddItem "Dateofbirth"
      cmbFields.AddItem "Remarks"
  
 End If
 cmbFields.ListIndex = 0   'always show name
End Sub

Private Sub submnuHelp_Click()
Load frmHelp
frmHelp.Show
End Sub

Private Sub submnuModify_Click()
Command_Click (1)
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
''' for splash
If VIN = 255 Then
  Timer1.Interval = 200
    GoTo skiper
 End If
VIN = VIN + 1
Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED)
Call GetWindowLong(Me.hwnd, GWL_EXSTYLE)
Call SetLayeredWindowAttributes(Me.hwnd, 0, VIN, LWA_ALPHA)
If VIN = 1 Then frmStart.Visible = True
If VIN < 255 Then Exit Sub 'dont proceed next
''///////////////////////
skiper:
'Dim temp As String


 frmStart.Caption = Mid$(scroll, madhya, Len(scroll) - madhya)
 'temp = Mid$(scroll, 1, madhya)
  frmStart.Caption = frmStart.Caption & Mid$(scroll, 1, madhya) 'temp
 madhya = madhya + 1
 If madhya > Len(scroll) Then
  madhya = 1
 End If
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command_Click (0)
End Sub

Private Sub vinrem_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell("vinreminder.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'vinreminder.EXE' is not found in its " _
  & "Default directory "

End Sub
