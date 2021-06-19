VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALL2HTML CONVERTER V2.0"
   ClientHeight    =   5370
   ClientLeft      =   1935
   ClientTop       =   1995
   ClientWidth     =   5505
   DrawStyle       =   1  'Dash
   FillStyle       =   5  'Downward Diagonal
   Icon            =   "HTMLCONVERTER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5370
   ScaleWidth      =   5505
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Any script (java/VB) in pages to be converted"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2760
      TabIndex        =   18
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Frame frBack 
      BackColor       =   &H00FFE7D9&
      Caption         =   "Background color of webpage"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   2655
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   120
         Shape           =   3  'Circle
         Top             =   240
         Width           =   150
      End
      Begin VB.Image imgPreviewSet 
         Height          =   225
         Left            =   120
         Top             =   240
         Width           =   240
      End
      Begin VB.Label frText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR of text"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Include Remark at End of File"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Prompt For File Name"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   240
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtDefault 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   "C:\ALL2HTML\"
      ToolTipText     =   "All the converted files will be saved in this default folder specified by you"
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&HTML Coding  "
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
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton optMin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Minimum  (fast)"
         Height          =   495
         Left            =   2400
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optMod 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moderate"
         Height          =   495
         Left            =   1320
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optMax 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Maximum  (less fast)  "
         Height          =   495
         Left            =   120
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F8F8E4&
      Caption         =   "CONVERT SELECTED FILES"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "First Select the files from above and then press to convert them"
      Top             =   4320
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   2760
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   105
      TabIndex        =   3
      Top             =   960
      Width           =   2640
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1815
      Left            =   2520
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"HTMLCONVERTER.frx":1CCA
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Converted Files in "
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
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   5295
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default Folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "All the converted files will be saved in the default folder specified by you"
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Same Folder"
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
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "All the converted files will be saved in the same folder"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:- No files Selected ......."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   5295
   End
   Begin VB.Menu setting 
      Caption         =   "&Settings"
      Begin VB.Menu filename 
         Caption         =   "Prompt for Each filename"
         Checked         =   -1  'True
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu smax 
         Caption         =   "Maximum Html Coding"
         Shortcut        =   {F3}
      End
      Begin VB.Menu smod 
         Caption         =   "Moderate Html Coding"
         Shortcut        =   {F4}
      End
      Begin VB.Menu smin 
         Caption         =   "Minimum Html Coding"
         Shortcut        =   {F5}
      End
      Begin VB.Menu er 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetpreview 
         Caption         =   "Change View of webpage"
         Shortcut        =   ^V
      End
      Begin VB.Menu fold 
         Caption         =   "Save in Same Folder"
         Shortcut        =   ^S
      End
      Begin VB.Menu default 
         Caption         =   "Save In Default Folder"
         Shortcut        =   ^D
      End
      Begin VB.Menu sf 
         Caption         =   "-"
      End
      Begin VB.Menu sconvert 
         Caption         =   "Convert Selected Files"
         Shortcut        =   {F11}
      End
      Begin VB.Menu tre 
         Caption         =   "-"
      End
      Begin VB.Menu end 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu Credit 
         Caption         =   "Credit"
         Shortcut        =   ^C
      End
      Begin VB.Menu hlp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu adrishya 
      Caption         =   "convert"
      Visible         =   0   'False
      Begin VB.Menu vconvert 
         Caption         =   "Convert Selected Files"
      End
   End
   Begin VB.Menu preview 
      Caption         =   "preview"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "Change background color of webpage"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Change Text color"
         Index           =   1
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ******************************
'   BY VINOD KOTIYA
 ' Created By VINOD KOTIYA Created On 01-March-2003 during INDIA-PAK Match
'Dedicted to Indian Teams Victory "
' version 1.0  released
'  FIRST MODIFICATION :- 06-06-2003 adding webpage preview
' second modification :- 07-06-2003 adding script facility
'     version 2.0 released

Dim InitialFolder
Dim totalFiles As Integer
Dim optmaxonetime As Boolean  'show maximum coding warning only one time if multiple files are selected
                   'for this option
Dim promptfilename As Boolean
Dim txtRemark_at_end As String 'store the user's remark that will be added at end of every converted file
Dim hcodeBACK As String    'store hex code of background color of page
Dim hcodeTEXT As String    'store hex code of text color of page
Dim trimfileName As String    'store pure filename

Private Sub about_Click()
MsgBox "ABOUT" & Chr(13) & _
"*************************************" & Chr(13) & Chr(13) & _
"ALL2HTML CONVERTER v2.0" & Chr(13) & "             is a part of VIN UTILITY KIT " & Chr(13) _
& "Created By VINOD KOTIYA " & Chr(13) & "Created On 01-March-2003 during INDIA-PAK Match " & Chr(13) & _
"Dedicted to Indian Teams Victory " & vbCrLf & _
 "version 1.0  released " & vbCrLf & _
"FIRST MODIFICATION :- 06-06-2003 adding webpage preview" & vbCrLf & _
"second modification :- 07-06-2003 adding script facility" & vbCrLf & _
"version 2.0 released"

End Sub


Private Sub Check1_Click()
filename.Checked = Not filename.Checked

If promptfilename = True Then
 promptfilename = False

Else
 promptfilename = True

End If
'

End Sub

Private Sub Check2_Click()
If Check2.Value Then
  txtRemark_at_end = InputBox("Enter the Remark you want to add at the end of all file : ", "User's information")
 Else
 txtRemark_at_end = "" 'user  wants no info added at end of file
 End If
End Sub

Private Sub Check3_Click()
If Check3.Value Then
  
  frm2.Visible = True
End If
End Sub

Private Sub Combo1_Click()
 Dir1_Change
End Sub

Private Sub Combo1_LostFocus()
 Dir1_Change
End Sub


Sub ScanFolders()
Dim subFolders As Integer

    totalFiles = totalFiles + File1.ListCount
    subFolders = Dir1.ListCount
    If subFolders > 0 Then
        For i = 0 To subFolders - 1
            ChDir Dir1.List(i)
            Dir1.Path = Dir1.List(i)
            File1.Path = Dir1.List(i)
            Form1.Refresh
            ScanFolders
        Next
    End If
    File1.Path = Dir1.Path
    MoveUp
End Sub

Sub MoveUp()
    If Dir1.List(-1) <> InitialFolder Then
        ChDir Dir1.List(-2)
        Dir1.Path = Dir1.List(-2)
    End If
End Sub










Private Sub Command2_Click()
Dim i As Integer
'check wheather any file is selected or not
'if no files selected than i = file1.listcount
For i = 0 To File1.ListCount - 1
 If File1.Selected(i) = True Then
   Exit For
 End If
Next

If i = File1.ListCount Then    'no files selected
 MsgBox "Please First select any file for conversion " & Chr(13) _
  & " Use Ctrl key for multiple sellection "
 Exit Sub
End If

'yaa any file is selected so come in to action

Dim filename As String    'store htm file name to be saved
Dim dirname As String   'edit drive name
dirname = Dir1.Path
If Len(dirname) < 4 Then    'If "F:\" = 3
 dirname = Left(dirname, 2)  'if "F:\" it return "F:"
End If
 
'check default folder is valid or not
   If Option2.Value = True Then
    IsDefaultValid
   End If
'Screen.MousePointer = vb

For i = 0 To File1.ListCount - 1
lblStatus.Caption = " Status : Converting File Please wait for a while..."
 If File1.Selected(i) = True Then    'only choose the selected file
   trimfileName = File1.List(i)    'store pure file name globally used in convertmin or convertmidmax function
   lblStatus.Caption = "Status:-  Converting file " & trimfileName & " in to webpage"
   loadrtf             ''load rtf box with initial data
     If optMod.Value = True Or optMax.Value = True Then
    converttohtmlModMax (dirname & "\" & File1.List(i))
   'RichTextBox1.LoadFile Dir1.Path & "\" & File1.List(i), rtfText
   ElseIf optMin.Value = True Then
    converttohtmlMIN (dirname & "\" & File1.List(i))
   End If
   filename = Left(File1.List(i), InStr(File1.List(i), ".")) 'return file name with '.' eg test.
   filename = filename & "html"  'now make test.html
   If promptfilename = True Then
     filename = InputBox("Do you want to enter any other file name to be saved ", "File Name Confirmation", filename)
     If vbCancel = True Then
      MsgBox "Cancelled"
     End If
   End If
   
   saveashtml (filename)
 End If
Next
lblStatus.Caption = "Status:-  Convertion complete"
MsgBox "All selected files are converted and saved in " & Chr(13) & _
 "Default directory " & txtDefault.Text & " Or In Same Folder "
 
End Sub



Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
MsgBox "You have selected Maximum HTML Coding option " _
 & "This will take a few minutes to convert multiple files (more than 3 )", vbYesNo
If vbYes Then
 MsgBox "yes"
End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.BackColor = &HFFC0FF
End Sub

Private Sub Credit_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\credit.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'credit.EXE' is not found in its " _
  & "Default directory  "

End Sub

Private Sub default_Click()
default.Checked = Not default.Checked
fold.Checked = False
Option2.Value = True
End Sub

Private Sub Dir1_Change()
    ChDir Dir1.Path
    File1.Path = Dir1.Path
    File1.Pattern = Combo1.Text    '"*.jpg ;*.bmp"
    lblStatus.Caption = "Status:- No files Selected........"

End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = "Switch to the directory where the text file you want to convert is placed"
End Sub

Private Sub Drive1_Change()
On Error GoTo vinerror
    ChDrive Dir1.Path
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    Exit Sub
vinerror:
   MsgBox "There is no disk in drive " & Drive1.Drive
End Sub


Private Sub end_Click()
End
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
For i = 0 To File1.ListCount - 1
 If File1.Selected(i) = True Then
   Exit For
 End If
Next

    'if files selected

If Button = 2 And i <> File1.ListCount Then PopupMenu adrishya
If i <> File1.ListCount Then lblStatus.Caption = "Status :- File Selected "
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = "You can use Ctrl or Shift Key for multiple file selection."
End Sub

Private Sub filename_Click()

Option3.Value
End Sub

Private Sub fold_Click()
fold.Checked = Not fold.Checked
Option1.Value = True
default.Checked = False
End Sub

Private Sub Form_Load()
'Load frm2
Dim fsys As New FileSystemObject
On Error GoTo vinerror
'crete folder if not exist
If fsys.FolderExists("C:\ALL2HTML") = False Then
 fsys.CreateFolder ("C:\ALL2HTML")
End If

    ChDrive "c:\" 'App.Path
    'ChDir "\VINOD DOC\resume\" '\App.Path
   
' initialize globals
optmaxonetime = True
promptfilename = True
hcodeBACK = "#D9E7FF"
hcodeTEXT = "#0000FF"

 Combo1.AddItem "*.TXT"
 Combo1.AddItem "*.*"
 Combo1.ListIndex = 0
 File1.Pattern = Combo1.Text
 'get default folder
 RichTextBox1.LoadFile App.Path & "\data\all2html.vin", rtfText
 txtDefault.Text = RichTextBox1.Text
 RichTextBox1.Text = ""
 
 Exit Sub
vinerror:
 MsgBox "Error during creating folder c:\all2html "
End Sub

Private Sub converttohtmlModMax(filename As String)
'RichTextBox1.LoadFile filename, rtfText
Dim InFile As Integer   ' Descriptor for file.
Dim currentline As String   'take a line of file to be converted
Dim messagefile As String
Dim j As Integer
j = 1
'OPEN THE FILE DATE.VIN AND STORE EACH LINE IN NEXTDATE TILL END
'AND CHECK THE MESSAGE IS ON,AFTER OR BEFORE TODAYS DATE
On Error GoTo FileError
InFile = FreeFile
Open filename For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, currentline
        RichTextBox1.Text = RichTextBox1.Text & currentline & "<br>" & Chr(13)
        
        'MsgBox currentline
        
    Wend
  Close InFile
  RichTextBox1.Text = RichTextBox1.Text & "<h4>" & txtRemark_at_end & "</h4></body>" & Chr(13) & "</html>" & Chr(13) _
   & "<hr><h6><marquee>Created by vinod kotiya's ALL2HTML converter</marquee></h6> "
If optMax = True Then
   replacespace  'with &nbsp;
End If
   Exit Sub
FileError:
  '  If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file data\dates.vin" & "" _
     & "The file may not exist on your harddrive "
End Sub
Private Sub converttohtmlMIN(filename As String)
'minimum coding only used <PRE> after "text"
'first directly load in rtfbox then add </body></html> and any remark
'now goto start and add starting tags <html>....
'also add scripts b/n head or body according to value set

Dim pos As Long

 RichTextBox1.LoadFile filename, rtfText
 RichTextBox1.Text = RichTextBox1.Text & "<PRE> <H4>" & txtRemark_at_end & _
 "</H4> <hr> <marquee> <h6> created by VINOD KOTIYA's ALL2HTML Converter </h6></marquee></body></html>"
If frm2.optHead.Value = True Then
RichTextBox1.SelStart = 0
 RichTextBox1.SelText = "<!-- By VINOD KOTIYA S-2 SHRIMAYA APARTMENT SECTOR-B/363 SARVDHARM COLONY BHOPAL FONE : 2794428---->" & Chr(13) & _
  "<HTML>" & Chr(13) & "<HEAD> <TITLE> converted from " & trimfileName & " by  vinod kotiya's ALL2HTML converter </TITLE> " & frm2.txtScript.Text & _
  " </HEAD>" & Chr(13) _
 & "<BODY BGCOLOR = " & Chr(34) & hcodeBACK & Chr(34) & "TEXT = " & Chr(34) & hcodeTEXT & Chr(34) & "> " & Chr(13) & _
 "<PRE>"

ElseIf frm2.optBody.Value = True Then
RichTextBox1.SelStart = 0
  RichTextBox1.SelText = "<!--By VINOD KOTIYA S-2 SHRIMAYA APARTMENT SECTOR-B/363 SARVDHARM COLONY BHOPAL FONE : 2794428---->" & Chr(13) & _
 "<HTML>" & Chr(13) & "<HEAD> <TITLE> converted from " & trimfileName & " by vinod kotiya's ALL2HTML converter </TITLE> " & _
  "</HEAD>" & Chr(13) _
& "<BODY BGCOLOR = " & Chr(34) & hcodeBACK & Chr(34) & "TEXT = " & Chr(34) & hcodeTEXT & Chr(34) & "> " & Chr(13) & frm2.txtScript.Text
 RichTextBox1.SelStart = 0
'====RichTextBox1.SelLength = 1
pos = RichTextBox1.Find("</script>", 0, , 0)
  If pos > 0 Then
     RichTextBox1.SelStart = pos + 9
     '=====RichTextBox1.SelLength = 1
     RichTextBox1.SelText = "<PRE>"
  Else
  pos = RichTextBox1.Find("TEXT = ", 0, , 0)
     RichTextBox1.SelStart = pos + 9
    
   RichTextBox1.SelText = "<PRE>"
   End If
End If
End Sub
Private Sub replacespace()

Dim temp As Long
temp = 2
Screen.MousePointer = vbHourglass
 While temp > 1
   temp = RichTextBox1.Find(Chr(32), temp + 180, , Binary)
   RichTextBox1.SelText = "&nbsp;"
 Wend
Screen.MousePointer = vbNormal
End Sub
Private Sub saveashtml(filename As String)
Dim destination As String
On Error GoTo vinerror
If Option1.Value = True Then
  destination = Dir1.Path & "\" & filename
Else
  destination = txtDefault.Text & filename
End If
 RichTextBox1.SaveFile destination, rtfText
 Exit Sub
 
vinerror:
 MsgBox "Please check the destination Default directory " _
 & destination
End Sub
Private Sub loadrtf()


RichTextBox1.Text = "<!--By VINOD KOTIYA S-2 SHRIMAYA APARTMENT SECTOR-B/363 SARVDHARM COLONY BHOPAL FONE : 2794428---->" & Chr(13)
RichTextBox1.Text = RichTextBox1.Text + "<HTML>" & Chr(13) & "<HEAD> <TITLE> converted from " & trimfileName & " by vinod kotiya's All2Html converter </TITLE>"
If frm2.optHead.Value = True Then
  RichTextBox1.Text = RichTextBox1.Text & frm2.txtScript.Text & " </HEAD>" & Chr(13) _
   & "<BODY BGCOLOR = " & Chr(34) & hcodeBACK & Chr(34) & "TEXT = " & Chr(34) & hcodeTEXT & Chr(34) & "> " & Chr(13)
ElseIf frm2.optBody.Value = True Then
  RichTextBox1.Text = RichTextBox1.Text & " </HEAD>" & Chr(13) _
   & "<BODY BGCOLOR = " & Chr(34) & hcodeBACK & Chr(34) & "TEXT = " & Chr(34) & hcodeTEXT & Chr(34) & "> " & Chr(13) & frm2.txtScript.Text
End If
End Sub





Private Sub Form_Unload(Cancel As Integer)
'save default folders name
On Error GoTo vinerror
RichTextBox1.Text = ""
RichTextBox1.Text = txtDefault.Text
RichTextBox1.SaveFile App.Path & "\data\all2html.vin", rtfText
 Unload frm2
 Exit Sub
vinerror:
 MsgBox "unable to write on disk path not found "
 Unload frm2
 Unload Me
 End Sub






Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)
lblStatus.Caption = "Determines how much coding will be done on each webpage"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.BackColor = &HF8F8E4
lblStatus.Caption = "Determines where to save the converted files"
End Sub

Private Sub frBack_DblClick()
PopupMenu preview
End Sub



Private Sub frBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu preview
End Sub

Private Sub frBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = "Determines the background and foreground color of output webpage."
End Sub

Private Sub hlp_Click()
MsgBox "HELP1:-  " & Chr(13) & _
"**************" & Chr(13) & Chr(13) & _
 "To convert all files (generally text) to HTML:" & Chr(13) & _
 "*************************************" & Chr(13) & Chr(13) & _
 "Change to the drive and folder containing the text files." & _
   "Select all files (using Ctrl ) you wish to convert." & Chr(13) & _
   "Click CONVERT ALL FILES button or press the F11 key." & Chr(13) & _
   "----------------------------------" & Chr(13) & Chr(13) & _
  "To Enable/Disable Prompt for filename" & Chr(13) & _
  "*************************************" & Chr(13) & Chr(13) & _
  "By default, a dialog will pop up to allowing you to change the name of every file converted. To turn this feature off, choose settings|Prompt for FileName Or press Ctrl + P from the main menu. When this menu item is unchecked no dialog box will be seen, and converted files will have the same name as the text file but with a .Html extension." & Chr(13) & _
  "It will Prevent the overwriting of any existing file." & Chr(13) & _
  "Please do not change the default extention .html otherwise file will not converted in html format " & Chr(13) & _
  "------------------------------------" & Chr(13) & Chr(13) & _
"Conversion of DOC files is not possible. Your WORD PROCESSOR also provide export in HTML option."
MsgBox "HELP2:-  " & Chr(13) & _
"*******" & Chr(13) & Chr(13) & _
 "HTML Coding option: " & Chr(13) & _
 "*****************" & Chr(13) & _
 "There are 3 options Minimum (fast) which provide less coding on your webpages and very fast" & _
   "Moderate option provide midium coding and Maximum(least fast) option provide maximum coding" & Chr(13) & _
   "Hence it will take too much time to convert file morethan 100 kb" & Chr(13) & Chr(13) & _
   "Change Settings of Page to be Created" & Chr(13) & _
  "****************************" & Chr(13) & _
  "To change the background color or Text color of webpage to be created Click on the Green Circle or press Ctrl + V" & Chr(13) & Chr(13) & _
  "Include Remark at end of Web page" & Chr(13) & _
  "**************************" & Chr(13) & _
  "You can add any remark at end of each webpage to be converted like Backword Link or any information." & Chr(13) & Chr(13) & _
  "Add the script in Web Pages" & Chr(13) & _
  "********************" & Chr(13) & _
  "A window will popup where you can type your vb or java script tobe added Also specify the portion <HEAD> or <BODY> where Script to be added. Some spectacular inbuilt scripts are also provided press Alt + C " & Chr(13) & Chr(13) & _
 "NOTE :- Include Remark and Add Script will merge same remarks and script in all files when Multiconversion is done."
  
  End Sub

Private Sub imgPreviewSet_Click()
PopupMenu preview
End Sub



Private Sub mnuPreview_Click(Index As Integer)
' since ullu ka pattha vb accept BGR color
'so you have to do some kasrat to get RGB color for web pages
'show the BGR color in VB and store RGB color for webpage

Dim CDFlags As Long
Dim Lal As Integer, Hara As Integer, Nila As Integer
Dim Rang As Long

On Error GoTo ColorError

    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)

    CommonDialog1.Flags = CDFlags
    CommonDialog1.CancelError = True
        CommonDialog1.ShowColor
   If Index = 0 Then
 
    Rang& = CommonDialog1.Color      'obtained BGR color
    'now convert it in to RGB color
    frBack.BackColor = Rang&    'long value of color
   Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    hcodeBACK = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)

   ElseIf Index = 1 Then
   Rang& = CommonDialog1.Color
   frText.ForeColor = Rang&
   Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    hcodeTEXT = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)
   End If
     
    
    
      'txtDefault.Text = hcodeBACK & "  " & hcodeTEXT ' (frText.ForeColor)
    Exit Sub
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
    
    Else
        MsgBox "An error occured"
    End If

End Sub

Private Sub mnuSetpreview_Click()
imgPreviewSet_Click
End Sub

Private Sub Option1_Click()
txtDefault.Enabled = False
End Sub

Private Sub Option2_Click()
txtDefault.Enabled = True
End Sub


Private Sub optMax_Click()
If optmaxonetime = True Then
  MsgBox "You have selected Maximum HTML Coding option " & Chr(13) & _
  "This will take a few minutes to convert files more than 100 KB )" & Chr(13) & _
  "Recomended for file having size less than 100 kb"
   optmaxonetime = False
 End If
  
End Sub
Private Sub IsDefaultValid()
 If Trim(txtDefault.Text) = "" Then
  MsgBox "Please specify your default folder name where " _
  & "You want to generate the html files " & Chr(13) _
  & " Like C:\MyHTML\ "
 End If
 Dim temp As String
 temp = Right(txtDefault.Text, 1)
 If StrComp(temp, "\") <> 0 Then
  txtDefault.Text = txtDefault.Text & "\"
 End If
 
' Dim fsys As New FileSystemObject
'On Error GoTo vinerror
'crete folder if not exist
'If fsys.FolderExists(txtDefault.Text) = False Then
' MsgBox Not "Default Folder" & txtDefault.Text _
' & "does not exist. Do you want to create it ?"
'End If
 Exit Sub
 
vinerror:
 MsgBox "An filesystem error occured"
End Sub

Private Sub sconvert_Click()
Command2_Click
End Sub

Private Sub smax_Click()
smax.Checked = Not smax.Checked
optMax = smax.Checked
smin.Checked = False
smod.Checked = False
End Sub

Private Sub smin_Click()
smin.Checked = Not smin.Checked
optMin = smin.Checked
smax.Checked = False
smod.Checked = False
End Sub


Private Sub smod_Click()
smod.Checked = Not smod.Checked
optMod = smod.Checked
smin.Checked = False
smax.Checked = False
End Sub

Private Sub vconvert_Click()
sconvert_Click
End Sub

