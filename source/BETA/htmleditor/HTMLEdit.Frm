VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form HTMLEdit 
   Caption         =   "HTML Editor"
   ClientHeight    =   6285
   ClientLeft      =   390
   ClientTop       =   1590
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6285
   ScaleWidth      =   10245
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   1560
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BEKAR"
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3180
      Left            =   90
      TabIndex        =   0
      Top             =   2955
      Width           =   9960
      ExtentX         =   17568
      ExtentY         =   5609
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2775
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   16777152
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"HTMLEdit.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7005
      Top             =   -135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.17485e-38
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu FileNew 
         Caption         =   "New Document"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "Open Document"
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save Document"
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save As ..."
      End
      Begin VB.Menu NavigateTo 
         Caption         =   "Open URL"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "Edit"
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu EditPaste 
         Caption         =   "Paste"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu EditClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu EditAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu ShorcutSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu EditRender 
         Caption         =   "Render Document"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu shortcutMenu 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu ShortcutCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu ShortcutCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu ShortcutPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu ShortcutClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu ShortcutAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu ShortcutSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu ShortcutRender 
         Caption         =   "Render Document"
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "Format"
      Begin VB.Menu submnuBold 
         Caption         =   "Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu submnuUnderline 
         Caption         =   "Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu submnubig 
         Caption         =   "Big"
      End
   End
End
Attribute VB_Name = "HTMLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RestoreText
Dim OpenFileName
Dim ResizeWindows As Boolean
Dim DragStartY As Integer
Dim HTMLHeight As Integer, WebHeight As Integer
Dim position As Long   'FOR SEARCHING AND COLOR CHANGING
Dim fromhere As Long
Dim tothere As Long
'Dim isrtfboxloded As Boolean

Private Sub Command1_Click()
'MsgBox Chr(156) & Chr(171) & Chr(172) & Chr(173) & Chr(174)
If tothere > fromhere Then
MsgBox tothere & " DFDD" & fromhere
  RichTextBox1.SelStart = fromhere
  RichTextBox1.SelLength = tothere - fromhere + 1
  'RichTextBox1.SelText = "" 'RichTextBox1.SelText
  'MsgBox RichTextBox1.SelText
 RichTextBox1.SelColor = vbBlue
 'MsgBox tothere & " DFDD" & fromhere
End If

End Sub

Private Sub Command2_Click()


End

''
position = RichTextBox1.Find("<U>", position + 1, , 0)
  While position > 0
     MakeBlack (2)
      position = RichTextBox1.Find("<U>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("</U>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("</U>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("<H1>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<H1>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("</H1>", position + 1, , 0)
  While position > 0
      MakeBlack (4)
      position = RichTextBox1.Find("</H1>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("<H2>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<H2>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("</H2>", position + 1, , 0)
  While position > 0
     MakeBlack (4)
      position = RichTextBox1.Find("</H2>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("<H3>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<H3>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("</H3>", position + 1, , 0)
  While position > 0
      MakeBlack (4)
      position = RichTextBox1.Find("</H3>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("<H4>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<H4>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("</H4>", position + 1, , 0)
  While position > 0
      MakeBlack (4)
      position = RichTextBox1.Find("</H4>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("<H5>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<H5>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("</H5>", position + 1, , 0)
  While position > 0
      MakeBlack (4)
      position = RichTextBox1.Find("</H5>", position + 1, , 0)
  Wend

position = RichTextBox1.Find("<H6>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<H6>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</H6>", position + 1, , 0)
  While position > 0
      MakeBlack (4)
      position = RichTextBox1.Find("</H6>", position + 1, , 0)
  Wend
        
position = RichTextBox1.Find("<P>", position + 1, , 0)
  While position > 0
      MakeBlack (2)
      position = RichTextBox1.Find("<P>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</P>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("</P>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<I>", position + 1, , 0)
  While position > 0
      MakeBlack (2)
      position = RichTextBox1.Find("<I>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</I>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("</I>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<BIG>", position + 1, , 0)
  While position > 0
      MakeBlack (4)
      position = RichTextBox1.Find("<BIG>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</BIG>", position + 1, , 0)
  While position > 0
      MakeBlack (5)
      position = RichTextBox1.Find("</BIG>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<SMALL>", position + 1, , 0)
  While position > 0
      MakeBlack (6)
      position = RichTextBox1.Find("<SMALL>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</SMALL>", position + 1, , 0)
  While position > 0
      MakeBlack (7)
      position = RichTextBox1.Find("</SMALL>", position + 1, , 0)
  Wend
  
position = RichTextBox1.Find("<PRE>", position + 1, , 0)
  While position > 0
      MakeBlack (4)
      position = RichTextBox1.Find("<PRE>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</PRE>", position + 1, , 0)
  While position > 0
      MakeBlack (5)
      position = RichTextBox1.Find("</PRE>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<STRIKE>", position + 1, , 0)
  While position > 0
      MakeBlack (7)
      position = RichTextBox1.Find("<STRIKE>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</STRIKE>", position + 1, , 0)
  While position > 0
      MakeBlack (8)
      position = RichTextBox1.Find("</STRIKE>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<CENTER>", position + 1, , 0)
  While position > 0
      MakeBlack (7)
      position = RichTextBox1.Find("<CENTER>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</CENTER>", position + 1, , 0)
  While position > 0
      MakeBlack (8)
      position = RichTextBox1.Find("</CENTER>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<FONT>", position + 1, , 0)
  While position > 0
      MakeBlack (5)
      position = RichTextBox1.Find("<FONT>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</FONT>", position + 1, , 0)
  While position > 0
      MakeBlack (6)
      position = RichTextBox1.Find("</FONT>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<TABLE>", position + 1, , 0)
  While position > 0
      MakeBlack (6)
      position = RichTextBox1.Find("<TABLE>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("</TABLE>", position + 1, , 0)
  While position > 0
      MakeBlack (7)
      position = RichTextBox1.Find("</TABLE>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<HR>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<HR>", position + 1, , 0)
  Wend
position = RichTextBox1.Find("<BR>", position + 1, , 0)
  While position > 0
      MakeBlack (3)
      position = RichTextBox1.Find("<BR>", position + 1, , 0)
  Wend

End Sub

Private Sub Command3_Click()
Dim i As Integer
i = 20
RichTextBox1.SelStart = 5
RichTextBox1.SelLength = i
RichTextBox1.SelColor = vbBlue
End Sub

Private Sub EditAll_Click()

    HTMLPad.RichTextBox1.SelStart = 0
    HTMLPad.RichTextBox1.SelLength = Len(HTMLPad.RichTextBox1.Text)
    
End Sub

Private Sub EditClear_Click()

    HTMLPad.RichTextBox1.Text = ""
    
End Sub

Private Sub EditCopy_Click()

    Clipboard.Clear
    Clipboard.SetText RichTextBox1.SelText
    
End Sub

Private Sub EditCut_Click()

    Clipboard.Clear
    Clipboard.SetText RichTextBox1.SelText
    RichTextBox1.SelText = ""

End Sub

Private Sub EditPaste_Click()

    temp = Clipboard.GetText(vbCFText)
    SelTextStart = RichTextBox1.SelStart
    SelTextEnd = Len(temp)
    RichTextBox1.SelText = temp
    RichTextBox1.SelFontName = RichTextBox1.Font
    RichTextBox1.SelBold = False
    RichTextBox1.SelItalic = False
    RichTextBox1.SelUnderline = False
    RichTextBox1.SelColor = vbBlack
    
    
End Sub

Private Sub EditRender_Click()

    RenderDocument
    
End Sub

Private Sub EditUndo_Click()

    RichTextBox1.Text = RestoreText
    
End Sub

Private Sub EditSelect_Click()

    RichTextBox1.SelStart = 0
    RichTextBox1.SelLength = Len(RichTextBox1.Text)
    
End Sub

Private Sub FileNew_Click()

    RichTextBox1.Text = ""
    OpenFileName = ""
    
End Sub

Private Sub FileOpen_Click()
On Error Resume Next

    CommonDialog1.Filter = "HML Documents|*.htm;*.html|ActiveX Documents|*.vbd|All Files|*.*"
    CommonDialog1.ShowOpen
    If Trim(CommonDialog1.FileName) = "" Then Exit Sub
    dPos = InStr(CommonDialog1.FileName, ".")
    If dPos > 0 Then ext = Mid$(CommonDialog1.FileName, dPos + 1)
    If UCase$(ext) = "HTM" Or UCase$(ext) = "HTML" Or UCase$(ext) = "TXT" Then
        RichTextBox1.LoadFile CommonDialog1.FileName, 1
        WebBrowser1.Navigate CommonDialog1.FileName
        OpenFileName = CommonDialog1.FileName
    End If
' The following lines handle non-HTML file types
' like sounds and images

    WebBrowser1.Navigate CommonDialog1.FileName
Rang_De

End Sub

Private Sub FileSave_Click()

    If OpenFileName <> "" Then
        RichTextBox1.SaveFile OpenFileName, 1
    Else
        FileSaveAs_Click
    End If
    
End Sub

Private Sub FileSaveAs_Click()

    CommonDialog1.DefaultExt = "htm"
    CommonDialog1.Filter = "HTML Documents|*.htm|All Files|*.*"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    RichTextBox1.SaveFile CommonDialog1.FileName, 1
    OpenFile = CommonDialog1.FileName
    
End Sub

Private Sub Form_Load()
MsgBox "VINOD"
RichTextBox1.Text = "<HTML> " & Chr(13) & "<HEAD><TITLE>WEB PAGE BY VINOD KOTIYA</TITLE></HEAD>" _
     & Chr(13) & "<BODY BGCOLOR = WHITE TEXT = BLUE ALINK = GREEN VLINK = RED>" & Chr(13) & _
      "<!-- start typing from here inside body -->" & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & _
     "</BODY>" & Chr(13) & "</HTML>"
 Rang_De
 'isrtfboxloded = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 And (Y > RichTextBox1.Top + RichTextBox1.Height) And (Y < WebBrowser1.Top) Then
        Screen.MousePointer = vbSizeNS
        ResizeWindows = True
        DragStartY = Y
        WebHeight = WebBrowser1.Height
        HTMLHeight = RichTextBox1.Height
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    If ResizeWindows Then
        RichTextBox1.Height = HTMLHeight + (Y - DragStartY)
        WebBrowser1.Move WebBrowser1.Left, RichTextBox1.Top + RichTextBox1.Height + 120, WebBrowser1.Width, WebHeight - (Y - DragStartY)
        HTMLEdit.Refresh
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResizeWindows = False
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Resize()

    RichTextBox1.Width = HTMLEdit.Width - RichTextBox1.Left - 200
    WebBrowser1.Width = RichTextBox1.Width
    RichTextBox1.Height = 0.3 * HTMLEdit.Height
    WebBrowser1.Move WebBrowser1.Left, RichTextBox1.Top _
    + RichTextBox1.Height + 120, WebBrowser1.Width, HTMLEdit.Height _
    - RichTextBox1.Top - RichTextBox1.Height - 940
    WebBrowser1.Navigate App.Path & "\empty.htm"
    
End Sub


Private Sub NavigateTo_Click()

    URL = InputBox("Enter URL to navigate to")
    If URL <> "" Then
        WebBrowser1.Navigate URL
    End If
    
End Sub

Private Sub RichTextBox1_Change()
' RenderDocument
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)

tothere = -10
If KeyCode = 188 Then
  RichTextBox1.SelText = Chr(156) ' "*"
  fromhere = RichTextBox1.Find(Chr(156), 0, , 0)
  RichTextBox1.SelText = ""
End If
If KeyCode = 190 Then
 'If Shift And vbShiftMask Then
 
 RichTextBox1.SelText = Chr(156) '"*"
 'Rang_De_limited
 'MsgBox "her"
 tothere = RichTextBox1.Find(Chr(156), fromhere - 1, , 0)
 'MsgBox "vif"
 RichTextBox1.SelText = ""
 'MsgBox "vif"
 
End If
If KeyCode = 13 Then
temp
End If
'If tothere > fromhere Then
 ' RichTextBox1.SelStart = fromhere
  'RichTextBox1.SelLength = tothere - fromhere
 ' RichTextBox1.SelText = RichTextBox1.SelText
  'MsgBox RichTextBox1.SelText
 'RichTextBox1.SelColor = vbBlue
 'MsgBox tothere & " DFDD" & fromhere
'End If
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu shortcutMenu
    End If
    
End Sub

Private Sub ShortcutAll_Click()
    Call EditAll_Click
End Sub


Private Sub ShortCutClear_Click()
    HTMLEdit.RichTextBox1.Text = ""
End Sub

Private Sub ShortcutCopy_Click()
    Call EditCopy_Click
End Sub


Private Sub ShortcutCut_Click()
    Call EditCut_Click
End Sub

Private Sub ShortcutPaste_Click()
    Call EditPaste_Click
End Sub

Private Sub ShortcutRender_Click()
    RenderDocument
End Sub

Private Sub submnubig_Click()
RichTextBox1.SelFontSize = 20
End Sub

Private Sub submnuBold_Click()
'   submnuBold.Checked = Not submnuBold.Checked
    RichTextBox1.SelBold = 1 'submnuBold.Checked
   RichTextBox1.SelText = "<B>" & RichTextBox1.SelText & "<B>"
End Sub

Private Sub submnuUnderline_Click()
 submnuUnderline.Checked = Not submnuUnderline.Checked
    RichTextBox1.SelUnderline = submnuUnderline.Checked
    RichTextBox1.SelText = "<U>" & RichTextBox1.SelText & "</U>"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox KeyCode
End Sub

Private Sub Timer1_Timer()
'If prevfromhere = fromhere Then

'Else
 If tothere > fromhere Then

  RichTextBox1.SelStart = fromhere
  RichTextBox1.SelLength = tothere - fromhere + 1
  'RichTextBox1.SelText = "" 'RichTextBox1.SelText
  'MsgBox RichTextBox1.SelText
 RichTextBox1.SelColor = vbBlue
 'MsgBox tothere & " DFDD" & fromhere
 End If
 'prevfromhere = fromhere
'End If
RenderDocument
End Sub

Private Sub WebBrowser1_Validate(Cancel As Boolean)
Debug.Print "Error"
End Sub
Private Sub MakeBlack(Lamba As Long)
 RichTextBox1.SelColor = vbBlue
 'LAMBA = 2 FOR 1CHAR  = 3 FOR 2CHAR  = 4 FOR 3CHAR
      RichTextBox1.SelStart = position + Lamba
      RichTextBox1.SelLength = 1   'SELECT ANY BLANK SPACE AND SET TO BLACK
      RichTextBox1.SelText = RichTextBox1.SelText '& Chr(32)
      RichTextBox1.SelBold = 0
      RichTextBox1.SelColor = vbBlack
      RichTextBox1.SelText = "   "
End Sub
Private Sub Rang_De()
On Error GoTo HI
'position = RichTextBox1.Find("<B>", position + 1, , 0)
 ' While position > 0
      'RichTextBox1.SelText = " " & RichTextBox1.SelText
     ' RichTextBox1.SelColor = vbBlue
     ' RichTextBox1.SelStart = position + 2
     ' RichTextBox1.SelLength = 1   'SELECT ANY BLANK SPACE AND SET TO BLACK
     ' RichTextBox1.SelText = RichTextBox1.SelText '& Chr(32)
     ' RichTextBox1.SelBold = 0
      'RichTextBox1.SelColor = vbBlack
      'RichTextBox1.SelText = "   "
  '    MakeBlack (2)
   '   position = RichTextBox1.Find("<B>", position + 1, , 0)
      
   'FINDING AND HIGHLIGTED THE TXT TO BE SEARCHED RETURN
   ' -VE IF NOT FOUND
   'DO OPERATIONS ON HIGHLITED TEXT
 ' Wend
  
'position = RichTextBox1.Find("</B>", position + 1, , 0)
 ' While position > 0
  '   MakeBlack (3)    'B/ IT HAS 2 CHARACTER
   '   position = RichTextBox1.Find("</B>", position + 1, , 0)
  'Wend

  Dim L As Long        'store latest position of >
  L = 1
position = 1

Do While position > 0      'exit when < not found
 position = RichTextBox1.Find("<", position + 1, , 0)
 
 If position < 0 Or RichTextBox1.SelColor = vbBlue Then
 'MsgBox "LESS"
   Exit Do          'exit when < not found
 End If
 L = RichTextBox1.Find(">", position + 1, , 0)
 RichTextBox1.SelStart = position
 RichTextBox1.SelLength = L - position + 1
'MakeBlack (2)
RichTextBox1.SelFontSize = 12
RichTextBox1.SelColor = vbBlue
 'LAMBA = 2 FOR 1CHAR  = 3 FOR 2CHAR  = 4 FOR 3CHAR
      RichTextBox1.SelStart = L
      RichTextBox1.SelLength = 1   'SELECT ANY BLANK SPACE AND SET TO BLACK
      RichTextBox1.SelText = RichTextBox1.SelText '& Chr(32)
      RichTextBox1.SelBold = 0
      RichTextBox1.SelColor = vbBlack
      RichTextBox1.SelText = "   "
    'position = RichTextBox1.Find("<", position + 1, , 0)
  Loop
  Exit Sub
HI:
MsgBox position & "eRROR"

End Sub

Private Sub Rang_De_limited()
Dim starpos As Long
position = 1
position = RichTextBox1.Find("*", position + 1, , 0)
starpos = position - 200
position = position - 200
Do While position > 0 And position < starpos + 2    'exit when < not found
 position = RichTextBox1.Find("<", position + 1, , 0)
 
 If position < 0 Or RichTextBox1.SelColor = vbBlue Then
 'MsgBox "LESS"
   Exit Do          'exit when < not found
 End If
 L = RichTextBox1.Find(">", position + 1, , 0)
 RichTextBox1.SelStart = position
 RichTextBox1.SelLength = L - position + 1
'MakeBlack (2)
RichTextBox1.SelFontSize = 12
RichTextBox1.SelColor = vbBlue
 'LAMBA = 2 FOR 1CHAR  = 3 FOR 2CHAR  = 4 FOR 3CHAR
      RichTextBox1.SelStart = L
      RichTextBox1.SelLength = 1   'SELECT ANY BLANK SPACE AND SET TO BLACK
      RichTextBox1.SelText = RichTextBox1.SelText '& Chr(32)
      RichTextBox1.SelBold = 0
      RichTextBox1.SelColor = vbBlack
      RichTextBox1.SelText = "   "
    'position = RichTextBox1.Find("<", position + 1, , 0)
  Loop

End Sub
Private Sub temp()
If tothere > fromhere Then
MsgBox tothere & " DFDD" & fromhere

RichTextBox1.SelStart = fromhere
  RichTextBox1.SelLength = tothere - fromhere + 1
  'RichTextBox1.SelText = "" 'RichTextBox1.SelText
  'MsgBox RichTextBox1.SelText
 RichTextBox1.SelColor = vbBlue
 End If
End Sub
