VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fone Directory"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7335
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "frmStart.frx":0442
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtSearch 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4080
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Modify"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
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
      ItemData        =   "frmStart.frx":3706
      Left            =   1320
      List            =   "frmStart.frx":3708
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Start Searching"
      Height          =   735
      Left            =   5040
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
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
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "To Search"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Menu mnuTask 
      Caption         =   "&Task"
      Begin VB.Menu submnuExtend 
         Caption         =   "Extend the Search In Options"
         Shortcut        =   ^E
      End
      Begin VB.Menu submnuClear 
         Caption         =   "Clear Previous Search List"
         Shortcut        =   ^C
      End
      Begin VB.Menu submnuModify 
         Caption         =   "Add/Delete Records"
         Shortcut        =   ^N
      End
      Begin VB.Menu submnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu submnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu submnuCredit 
         Caption         =   "&Credit"
      End
      Begin VB.Menu submnuAbout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OpenFile As String
Dim txt As String
Dim madhya As Integer
Dim unloadmdi As Boolean

Private Sub Command1_Click()
Load Form1
Form1.Visible = True
End Sub

Private Sub Command2_Click()
'//unloadmdi form before next searching
If unloadmdi = True Then      ' unload after first clicking over txtsearch
   Unload frmShow
   Unload MDIForm1
End If
unloadmdi = True
'/// now load
Load MDIForm1
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
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
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
 unloadmdi = False      'dont unload on first clicking over txtsearch
End Sub

Private Sub Form_Unload(Cancel As Integer)
 savethesearchtext
Unload Form1
Unload Form3
Unload frmShow
'Unload MDIForm1
Unload frmCredit
Unload frmAbout
Unload frmHelp
Unload Me

End Sub

Private Sub loadtxtsearch()


Dim FNum As Integer



On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open "data\vinod.vin" For Input As #1
    txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & "vinod.vin" _
     & "To eliminate problem Open Your TextEditor(NotePad) " _
     & "Click on File->Save and Save the file as 'vinod.vin' in FoneDirectory\data\vinod.vin "
     
    OpenFile = ""
    
End Sub


Private Sub savethesearchtext()
Dim FNum As Integer
Dim txt As String

Dim i As Integer
On Error GoTo FileError
    FNum = FreeFile
    Open "data\vinod.vin" For Output As #1
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
    MsgBox "Unkown error while saving file " & "vinod.vin" 'CommonDialog1.FileName
    OpenFile = ""
End Sub


Private Sub submnuAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub submnuClear_Click()
'deleting  previous search LIST
Dim i As Integer
 For i = txtSearch.ListCount - 1 To 0 Step -1
       txtSearch.RemoveItem (i)
 Next
    
End Sub

Private Sub submnuCredit_Click()
Load frmCredit
frmCredit.Show
End Sub

Private Sub submnuExit_Click()
Unload Me
End Sub

Private Sub submnuExtend_Click()
submnuExtend.Checked = Not submnuExtend.Checked
 If submnuExtend.Checked = False Then
'  cmbFields.RemoveItem (cmbFields.ListCount)  'email
  cmbFields.RemoveItem (cmbFields.ListCount - 1) 'Address
  cmbFields.RemoveItem (cmbFields.ListCount - 2) 'Address
 Else
 cmbFields.AddItem "Emails"
  If cmbFields.ListCount > 6 Then
  cmbFields.RemoveItem (6)
   cmbFields.AddItem "Address"
  End If
 
 cmbFields.AddItem "Mobile"
 cmbFields.AddItem "Post"
 End If
End Sub

Private Sub submnuHelp_Click()
Load frmHelp
frmHelp.Show
End Sub

Private Sub submnuModify_Click()
Load Form1
Form1.Visible = True
End Sub

Private Sub Timer1_Timer()
Dim scroll As String
Dim temp As String

scroll = "       Fone Directory By Vinod Kotiya     *********"
 frmStart.Caption = Mid$(scroll, madhya, Len(scroll) - madhya)
 temp = Mid$(scroll, 1, madhya)
  frmStart.Caption = frmStart.Caption & temp
 madhya = madhya + 1
 If madhya > Len(scroll) Then
  madhya = 1
 End If
End Sub

