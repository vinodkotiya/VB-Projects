VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileScan"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "HTML Coding  "
      Height          =   975
      Left            =   1680
      TabIndex        =   8
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton optMin 
         Caption         =   "Minimum  (fast)"
         Height          =   495
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optMod 
         Caption         =   "Moderate"
         Height          =   495
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMax 
         Caption         =   "Maximum   (slow)"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   -120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CONVERT"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan Now"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5655
      TabIndex        =   3
      Top             =   3855
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   2880
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1110
      Width           =   2715
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   120
      TabIndex        =   1
      Top             =   1380
      Width           =   2640
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   1080
      Width           =   2760
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3495
      Left            =   5640
      TabIndex        =   5
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"HTMLCONVERTER1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ******************************
'   BY VINOD KOTIYA
'  CREATED 26 FEB 2003
'  FIRST MODIFICATION
Dim InitialFolder
Dim totalFiles As Integer
Dim optmaxonetime As Boolean  'show maximum coding warning only one time if multiple files are selected
                   'for this option

Private Sub Command1_Click()
    ChDrive Drive1.Drive
    ChDir Dir1.Path
    InitialFolder = CurDir
    ScanFolders
    MsgBox "There are " & totalFiles & " under the " & InitialFolder & " folder"
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
Dim filename As String    'store htm file name to be saved
Dim dirname As String   'edit drive name
dirname = Dir1.Path
If Len(dirname) < 4 Then    'If "F:\" = 3
 dirname = Left(dirname, 2)  'if "F:\" it return "F:"
End If
 
For i = 0 To File1.ListCount - 1
 If File1.Selected(i) = True Then    'only choose the selected file
  loadrtf             ''load rtf box with initial data
   If optMod.Value = True Or optMax.Value = True Then
    converttohtmlModMax (dirname & "\" & File1.List(i))
   'RichTextBox1.LoadFile Dir1.Path & "\" & File1.List(i), rtfText
   ElseIf optMin.Value = True Then
    converttohtmlMIN (dirname & "\" & File1.List(i))
   End If
   filename = Left(File1.List(i), InStr(File1.List(i), ".")) 'return file name with '.' eg test.
   filename = filename & "html"  'now make test.html
   filename = InputBox("Do you want to enter any other file name to be saved ", , filename)
   saveashtml (filename)
 End If
Next
End Sub

Private Sub Command3_Click()
 'RichTextBox1.SaveFile Dir1.Path & "\" & "yes.html", rtfText
Dim filename As String
 filename = Left(File1.List(i), InStr(File1.List(i), "."))
 MsgBox filename
End Sub

Private Sub Command4_Click()
Dim what As String

'While what > 1
what = RichTextBox1.SelText
MsgBox Asc(what)
'what = RichTextBox1.Find(Chr(255), what + 1, , Binary)
'RichTextBox1.SelText = "x"
End Sub

Private Sub Command5_Click()
MsgBox "You have selected Maximum HTML Coding option " _
 & "This will take a few minutes to convert multiple files (more than 3 )", vbYesNo
If vbYes Then
 MsgBox "yes"
End If
End Sub

Private Sub Dir1_Change()
    ChDir Dir1.Path
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    ChDrive Dir1.Path
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
End Sub


Private Sub File1_DblClick()

End Sub

Private Sub Form_Load()
 optMod.Value = True
    ChDrive "f:\" 'App.Path
    ChDir "\VINOD DOC\resume\" '\App.Path
   
' initialize globals
optmaxonetime = True
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
'On Error GoTo FileError
InFile = FreeFile
Open filename For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, currentline
        RichTextBox1.Text = RichTextBox1.Text & currentline & "<br>"
        
        'MsgBox currentline
        
    Wend
  Close InFile
  RichTextBox1.Text = RichTextBox1.Text & "</body>" & Chr(13) & "</html>" & Chr(13) _
   & "Creted by vinod kotiya's * to html converter "
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
'directly load in rtfbox
Dim pos As Long
RichTextBox1.LoadFile filename, rtfText
RichTextBox1.Text = RichTextBox1.Text & "</PRE>" & Chr(13) _
 & "</BODY>" & Chr(13) & "</HTML>"
 RichTextBox1.SelStart = 0
 'RichTextBox1.SelLength = 1
 RichTextBox1.SelText = "<!--DocType HTML By VINOD KOTIYA's TEXT2HTML CONVERTER" & Chr(13) & _
  "<HTML>" & Chr(13) & "<HEAD> <TITLE> converted by vinod kotiya's * to html converter </TITLE> </HEAD>" & Chr(13) _
   & "<BODY BGCOLOR = " & Chr(34) & "#0099FF" & Chr(34) & "TEXT = >   "

 pos = RichTextBox1.Find("TEXT =", 0, , 0)
 RichTextBox1.SelStart = pos + 8
 'RichTextBox1.SelLength = 1
 RichTextBox1.SelText = "<PRE>"
End Sub
Private Sub replacespace()

Dim temp As Long
temp = 2
Screen.MousePointer = vbHourglass
 While temp > 1
   temp = RichTextBox1.Find(Chr(32), temp + 1, , Binary)
   RichTextBox1.SelText = "&nbsp;"
 Wend
Screen.MousePointer = vbNormal
End Sub
Private Sub saveashtml(filename As String)
 RichTextBox1.SaveFile Dir1.Path & "\" & filename, rtfText
End Sub
Private Sub loadrtf()
RichTextBox1.Text = ""
RichTextBox1.Text = "<!--DocType HTML By VINOD KOTIYA's TEXT2HTML CONVERTER" & Chr(13)
RichTextBox1.Text = RichTextBox1.Text + "<HTML>" & Chr(13) & "<HEAD> <TITLE> converted by vinod kotiya's * to html converter </TITLE> </HEAD>" & Chr(13) _
   & "<BODY BGCOLOR = " & Chr(34) & "#0099FF" & Chr(34) & "TEXT = >"

End Sub

Private Sub optMax_Click()
  MsgBox "You have selected Maximum HTML Coding option " _
   & "This will take a few minutes to convert multiple files (more than 3 )"
   
  
End Sub
