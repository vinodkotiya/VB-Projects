VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form HTMLEdit 
   BackColor       =   &H00FF0000&
   Caption         =   "VIN WEB-COMPILER V3.0"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   11880
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H000000FF&
   Icon            =   "HTMLEdit.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8130
   ScaleWidth      =   11880
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3240
      Top             =   1560
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5220
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Right Click for all Internet explorer option's to browse.- *VIN Web Compiler"
      Top             =   3000
      Width           =   11880
      ExtentX         =   20955
      ExtentY         =   9208
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
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Right click for inserting tags,or select your text to applly other tags."
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   16777152
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"HTMLEdit.frx":030A
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
      Caption         =   "&File"
      Begin VB.Menu FileNew 
         Caption         =   "New WebPage"
         Shortcut        =   ^N
      End
      Begin VB.Menu sample 
         Caption         =   "Sample WebPage with tutorial"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "Open WebPage"
         Shortcut        =   ^O
      End
      Begin VB.Menu NavigateTo 
         Caption         =   "Open Website"
      End
      Begin VB.Menu rter 
         Caption         =   "-"
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save Document"
         Shortcut        =   ^S
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save As ..."
      End
      Begin VB.Menu bc 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu EditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu EditAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditUndo 
         Caption         =   "UnDo"
         Visible         =   0   'False
      End
      Begin VB.Menu TYU 
         Caption         =   "-"
      End
      Begin VB.Menu find 
         Caption         =   "Find"
         Shortcut        =   {F3}
      End
      Begin VB.Menu rep 
         Caption         =   "Replace"
         Shortcut        =   {F4}
      End
      Begin VB.Menu fgf 
         Caption         =   "-"
      End
      Begin VB.Menu EditRender 
         Caption         =   "REAL TIME RENDERING"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu update 
         Caption         =   "Refresh WebPage"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu temp1 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu submnuStyle 
      Caption         =   "&Style"
      Begin VB.Menu submnuBold 
         Caption         =   "Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Underline"
         Index           =   1
         Shortcut        =   ^U
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Italic"
         Index           =   2
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Strikethrough"
         Index           =   3
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Big"
         Index           =   4
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Small"
         Index           =   5
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "SuperScript"
         Index           =   6
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Subscript"
         Index           =   7
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "StrongEmphasis"
         Index           =   8
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Emphasis"
         Index           =   9
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Defination"
         Index           =   10
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Code"
         Index           =   11
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Sample Output"
         Index           =   12
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Variable"
         Index           =   13
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Acronym"
         Index           =   14
      End
      Begin VB.Menu smnuStyles 
         Caption         =   "Abbriviation"
         Index           =   15
      End
   End
   Begin VB.Menu div 
      Caption         =   "&Align"
      Begin VB.Menu Aligns 
         Caption         =   "Left"
         Index           =   1
      End
      Begin VB.Menu Aligns 
         Caption         =   "Centre"
         Index           =   2
      End
      Begin VB.Menu Aligns 
         Caption         =   "Right"
         Index           =   3
      End
      Begin VB.Menu Aligns 
         Caption         =   "Justify"
         Index           =   4
      End
   End
   Begin VB.Menu submnuLine 
      Caption         =   "&Header/ Para"
      Begin VB.Menu ssmnuHeader 
         Caption         =   "H1"
         Index           =   1
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu ssmnuHeader 
         Caption         =   "H2"
         Index           =   2
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu ssmnuHeader 
         Caption         =   "H3"
         Index           =   3
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu ssmnuHeader 
         Caption         =   "H4"
         Index           =   4
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu ssmnuHeader 
         Caption         =   "H5"
         Index           =   5
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu ssmnuHeader 
         Caption         =   "H6"
         Index           =   6
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu separator2 
         Caption         =   "-"
      End
      Begin VB.Menu submnuPara 
         Caption         =   "Paragraph"
         Shortcut        =   ^P
      End
      Begin VB.Menu pre 
         Caption         =   "Preformat"
      End
      Begin VB.Menu Blockquote 
         Caption         =   "Blockquote"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "&Insert"
      Begin VB.Menu rule 
         Caption         =   "Horizontal Rule"
         Begin VB.Menu hrType 
            Caption         =   "Left Align"
            Index           =   1
         End
         Begin VB.Menu hrType 
            Caption         =   "Centre"
            Index           =   2
         End
         Begin VB.Menu hrType 
            Caption         =   "Right Align"
            Index           =   3
         End
      End
      Begin VB.Menu image 
         Caption         =   "Image"
         Begin VB.Menu imgType 
            Caption         =   "Image Left"
            Index           =   1
         End
         Begin VB.Menu imgType 
            Caption         =   "Image Center"
            Index           =   2
         End
         Begin VB.Menu imgType 
            Caption         =   "Image Right"
            Index           =   3
         End
      End
      Begin VB.Menu link 
         Caption         =   "Link"
         Shortcut        =   ^L
      End
      Begin VB.Menu anchor 
         Caption         =   "NamedAnchor"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu fontback 
      Caption         =   "Font / &Page"
      Begin VB.Menu font 
         Caption         =   "Font"
         Shortcut        =   ^F
      End
      Begin VB.Menu bgcolor 
         Caption         =   "BackGround Color"
      End
      Begin VB.Menu txtcol 
         Caption         =   "Text Color"
         Begin VB.Menu wincolor 
            Caption         =   "Choose Color"
            Shortcut        =   ^T
         End
         Begin VB.Menu indra 
            Caption         =   "IndraDhanush(color picker)"
         End
      End
   End
   Begin VB.Menu effect 
      Caption         =   "Effe&cts"
      Begin VB.Menu smarq 
         Caption         =   "Marquee"
         Shortcut        =   ^Q
      End
      Begin VB.Menu sblink 
         Caption         =   "Blink"
      End
      Begin VB.Menu ssoun 
         Caption         =   "Sound"
      End
      Begin VB.Menu vdo 
         Caption         =   "Vedio"
      End
   End
   Begin VB.Menu sscript 
      Caption         =   "Sc&ripts"
      Begin VB.Menu scripts 
         Caption         =   "TextArea"
         Index           =   1
      End
      Begin VB.Menu scripts 
         Caption         =   "Input"
         Index           =   2
      End
      Begin VB.Menu scripts 
         Caption         =   "Select"
         Index           =   3
      End
   End
   Begin VB.Menu symb 
      Caption         =   "Sy&mbol"
      Begin VB.Menu symbol 
         Caption         =   "Special Symbol"
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "space"
            Index           =   0
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "Copyright"
            Index           =   1
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "Registered"
            Index           =   2
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "Power1"
            Index           =   3
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "Square"
            Index           =   4
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "Cube"
            Index           =   5
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "Half"
            Index           =   6
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "oneforth"
            Index           =   7
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "threeforth"
            Index           =   8
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "micro"
            Index           =   9
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "para"
            Index           =   10
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "middot"
            Index           =   11
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "yen"
            Index           =   12
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "Pound"
            Index           =   13
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "cent"
            Index           =   14
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "iexcl"
            Index           =   15
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "curren"
            Index           =   16
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "brvbar"
            Index           =   17
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "sect"
            Index           =   18
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "uml"
            Index           =   19
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "ordf"
            Index           =   20
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "laquo"
            Index           =   21
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "not"
            Index           =   22
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "macr"
            Index           =   23
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "deg"
            Index           =   24
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "plusminus"
            Index           =   25
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "acute"
            Index           =   26
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "iquest"
            Index           =   27
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "times"
            Index           =   28
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "ordm"
            Index           =   29
         End
         Begin VB.Menu ssmnuSymbols 
            Caption         =   "raquo"
            Index           =   30
         End
      End
      Begin VB.Menu ass 
         Caption         =   "Accent Case"
      End
      Begin VB.Menu other 
         Caption         =   "Other Case"
         Begin VB.Menu ssmnuOthers 
            Caption         =   "beta"
            Index           =   0
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "phi"
            Index           =   1
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "ae"
            Index           =   2
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "eth"
            Index           =   3
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "thorn"
            Index           =   4
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "oelig"
            Index           =   5
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "AE"
            Index           =   6
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "ETH"
            Index           =   7
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "PHI"
            Index           =   8
         End
         Begin VB.Menu ssmnuOthers 
            Caption         =   "THORN"
            Index           =   9
         End
      End
   End
   Begin VB.Menu window 
      Caption         =   "&Window"
      Begin VB.Menu win 
         Caption         =   "Full Source Code View"
         Index           =   1
         Shortcut        =   {F8}
      End
      Begin VB.Menu win 
         Caption         =   "Full Webpage View"
         Index           =   2
         Shortcut        =   {F9}
      End
      Begin VB.Menu win 
         Caption         =   "Split Both "
         Checked         =   -1  'True
         Index           =   3
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu dfd 
      Caption         =   ""
   End
   Begin VB.Menu aboutit 
      Caption         =   "A&bout"
      Begin VB.Menu help 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu about 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
      Begin VB.Menu credit 
         Caption         =   "Credit"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu vin 
      Caption         =   "&Vin Utility Kit"
   End
   Begin VB.Menu adrishya1 
      Caption         =   "InsertAdrishya"
      Visible         =   0   'False
      Begin VB.Menu vinop 
         Caption         =   "-----(Insertion Option)-----"
      End
      Begin VB.Menu vsound 
         Caption         =   "Sound"
      End
      Begin VB.Menu vvedio 
         Caption         =   "Vedio"
      End
      Begin VB.Menu vimg 
         Caption         =   "Image"
         Begin VB.Menu vimgType 
            Caption         =   "Image Left"
            Index           =   1
         End
         Begin VB.Menu vimgType 
            Caption         =   "Image Centre"
            Index           =   2
         End
         Begin VB.Menu vimgType 
            Caption         =   "Image Right"
            Index           =   3
         End
      End
      Begin VB.Menu vhr 
         Caption         =   "Horizontal Rule"
         Begin VB.Menu vhrType 
            Caption         =   "Left Align"
            Index           =   1
         End
         Begin VB.Menu vhrType 
            Caption         =   "Centre Align"
            Index           =   2
         End
         Begin VB.Menu vhrType 
            Caption         =   "Right Align"
            Index           =   3
         End
      End
      Begin VB.Menu uyui 
         Caption         =   "-"
      End
      Begin VB.Menu vlink 
         Caption         =   "Link"
      End
      Begin VB.Menu vanch 
         Caption         =   "Named Anchor"
      End
      Begin VB.Menu vcrtr 
         Caption         =   "-"
      End
      Begin VB.Menu vscripts 
         Caption         =   "Text Area"
         Index           =   1
      End
      Begin VB.Menu vscripts 
         Caption         =   "Input"
         Index           =   2
      End
      Begin VB.Menu vscripts 
         Caption         =   "Select"
         Index           =   3
      End
      Begin VB.Menu oiop 
         Caption         =   "-"
      End
      Begin VB.Menu vbgc 
         Caption         =   "Background Color"
      End
      Begin VB.Menu vtxtcol 
         Caption         =   "Text Color"
      End
      Begin VB.Menu vpaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu vreal 
         Caption         =   "REAL TIME RENDERING"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu adrishya2 
      Caption         =   "Selective"
      Visible         =   0   'False
      Begin VB.Menu vHeader 
         Caption         =   "H1"
         Index           =   1
      End
      Begin VB.Menu vHeader 
         Caption         =   "H2"
         Index           =   2
      End
      Begin VB.Menu vHeader 
         Caption         =   "H3"
         Index           =   3
      End
      Begin VB.Menu vHeader 
         Caption         =   "H4"
         Index           =   4
      End
      Begin VB.Menu vHeader 
         Caption         =   "H5"
         Index           =   5
      End
      Begin VB.Menu vHeader 
         Caption         =   "H6"
         Index           =   6
      End
      Begin VB.Menu vpara 
         Caption         =   "Paragraph"
      End
      Begin VB.Menu vpre 
         Caption         =   "PreFormat"
      End
      Begin VB.Menu vblock 
         Caption         =   "BlockQuote"
      End
      Begin VB.Menu fdg 
         Caption         =   "-"
      End
      Begin VB.Menu vbold 
         Caption         =   "Bold"
      End
      Begin VB.Menu vunder 
         Caption         =   "Underline"
      End
      Begin VB.Menu vfont 
         Caption         =   "Font"
      End
      Begin VB.Menu vmark 
         Caption         =   "Marquee"
      End
      Begin VB.Menu ghj 
         Caption         =   "-"
      End
      Begin VB.Menu vcut 
         Caption         =   "Cut"
      End
      Begin VB.Menu vcopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "HTMLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                                           VIN WEB COMPILER
'//////////////////////////////////////////////////////////////////////////////////////////
'/////////////////   programmer:-      vinod kotiya //////////////////////////////////////
'////////////////     starting:- 15thJan 2003            //////////////////////////////////////
'////////////////     completed :- 12-04-2003       //////////////////////////////////////
'////////////////      v1.0 - 26-jan-2003   (not very efficient ////////////////////////////
'///////////////       v2.0  - 15-feb-2003   bug1:- don't show the relative paths data///
'///////////////       v3.0 - 12-04-2003     bug1 removed
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////    fone:- +91-0755-2794428 , vinodkotiya24@rediffmail.com///////
'////////////////////////////////////////////////////////////////////////////////////////

Dim RestoreText
Dim OpenFileName
Dim ResizeWindows As Boolean
Dim DragStartY As Integer
Dim HTMLHeight As Integer, WebHeight As Integer
Dim Position As Long   'FOR SEARCHING AND COLOR CHANGING
Dim fromhere As Long
Dim tothere As Long
Dim isrtfboxloded As Boolean
Dim isleftarrowpress As Boolean
Dim isRealTime As Boolean    'do rendering as typing when true
Dim txtModifySave As Boolean     'true when rtbox updated


Private Sub about_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell(App.Path & "\about.exe", vbNormalFocus)
MsgBox "Created on 15 th january to 31 st january 2003"
Exit Sub
Exeerror:
 MsgBox "Application 'about.EXE' is not found in its " _
  & "Default directory about.exe "
End Sub







Private Sub Aligns_Click(Index As Integer)
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
 If Index = 1 Then
     RichTextBox1.SelText = Chr(170) & "<DIV ALIGN = LEFT>  " & RichTextBox1.SelText & " </DIV> "
    Call MakemeBlue("<DIV ALIGN = LEFT>", "</DIV>", 5)
 ElseIf Index = 2 Then
     RichTextBox1.SelText = Chr(170) & "<DIV ALIGN = CENTER>  " & RichTextBox1.SelText & " </DIV> "
   Call MakemeBlue("<DIV ALIGN = CENTER>", "</DIV>", 5)
 ElseIf Index = 3 Then
          RichTextBox1.SelText = Chr(170) & "<DIV ALIGN = RIGHT>  " & RichTextBox1.SelText & " </DIV> "
   Call MakemeBlue("<DIV ALIGN = RIGHT>", "</DIV>", 5)
 ElseIf Index = 4 Then
    RichTextBox1.SelText = Chr(170) & "<DIV ALIGN = JUSTIFY>  " & RichTextBox1.SelText & " </DIV> "
   Call MakemeBlue("<DIV ALIGN = JUSTIFY>", "</DIV>", 5)

End If
End Sub

Private Sub anchor_Click()
RichTextBox1.SelText = Chr(170) & "<A NAME= >  " & RichTextBox1.SelText & " </A> "
Call MakemeBlue("<A NAME= >", "</A>", 8)

End Sub

Private Sub bgcolor_Click()
Dim CDFlags As Long
Dim Lal As Integer, Hara As Integer, Nila As Integer
Dim Rang As Long
Dim hcode As String
On Error GoTo ColorError

    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)

    CommonDialog1.Flags = CDFlags
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    
    Rang& = CommonDialog1.Color
    Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    hcode = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)
    Dim startpos As Integer
    startpos = RichTextBox1.find("BODY", , , 0)
    Rang = RichTextBox1.find("bgcolor", startpos, , 0)
    CDFlags = RichTextBox1.find("TEXT", Rang, Rang + 30, 0)
    RichTextBox1.SelStart = Rang + 10
    RichTextBox1.SelLength = CDFlags - Rang - 11

    RichTextBox1.SelText = hcode
    Exit Sub
    
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
    
    Else
        MsgBox "An error occured"
    End If

End Sub

Private Sub Blockquote_Click()
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
RichTextBox1.SelText = Chr(170) & "<BLOCKQUOTE>  " & RichTextBox1.SelText & " </BLOCKQUOTE> "
Call MakemeBlue("<BLOCKQUOTE>", "</BLOCKQUOTE>", 12)
End Sub









Private Sub Command1_Click()
 RichTextBox1.SelLength = 5
  RichTextBox1.SelLength = RichTextBox1.SelLength + 5
   
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



Private Sub Command3_Click()
Dim i As Integer
i = 20
RichTextBox1.SelStart = 5
RichTextBox1.SelLength = i
RichTextBox1.SelColor = vbBlue
End Sub



Private Sub credit_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell("credit.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'CREDIT.EXE' is not found in its " _
  & "Default directory CREDIT.exe "
End Sub







Private Sub EditAll_Click()

    RichTextBox1.SelStart = 0
    RichTextBox1.SelLength = Len(RichTextBox1.Text)
    
End Sub

Private Sub EditClear_Click()

    
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
    RichTextBox1.SelFontName = RichTextBox1.font
    RichTextBox1.SelBold = False
    RichTextBox1.SelItalic = False
    RichTextBox1.SelUnderline = False
    RichTextBox1.SelColor = vbBlack
    
    
End Sub

Private Sub EditRender_Click()
 EditRender.Checked = Not EditRender.Checked
 If EditRender.Checked = True Then
    RenderDocument
    isRealTime = True
 Else
   MsgBox "     REAL TIME RENDERING is turned off." & Chr(13) & _
   "Now you will not see the Real Time Preview of your webpage." & Chr(13) & _
   "You can press F7 to Refresh your webpage manually."
   isRealTime = False
 End If

End Sub

Private Sub EditUndo_Click()

    RichTextBox1.Text = RestoreText
    
End Sub

Private Sub EditSelect_Click()

    RichTextBox1.SelStart = 0
    RichTextBox1.SelLength = Len(RichTextBox1.Text)
    
End Sub







Private Sub FileExit_Click()
If txtModifySave = True Then
Dim reply As Integer
reply = MsgBox("Do you want to save the webpage", vbYesNo, "Prompt For Save")
If reply = 6 Then         'yes
  FileSave_Click
End If
End If

End Sub

Private Sub FileNew_Click()
If txtModifySave = True Then
Dim reply As Integer
reply = MsgBox("Do you want to save the webpage", vbYesNo, "Prompt For Save")
If reply = 6 Then         'yes
  FileSave_Click
End If
End If
HTMLEdit.Caption = "VIN Web Compiler v3.0 : New WebPage"
isrtfboxloded = False 'prevent rendering because of rtftext change event at time of loading
    loadnewdocument
    OpenFileName = ""
    Rang_De
 isrtfboxloded = True
 RenderDocument
 txtModifySave = False
End Sub

Private Sub FileOpen_Click()
If txtModifySave = True Then
Dim reply As Integer
reply = MsgBox("Do you want to save the webpage", vbYesNo, "Prompt For Save")
If reply = 6 Then         'yes
  FileSave_Click
End If
End If
isrtfboxloded = False 'prevent rendering because of rtftext change event at time of loading
On Error Resume Next
   isrtfboxloded = False
    CommonDialog1.Filter = "HTML Documents|*.htm;*.html|ActiveX Documents|*.vbd|All Files|*.*"
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
HTMLEdit.Caption = "VIN WEB-COMPILER v3.0  ( " & OpenFileName & " )"
Rang_De
    WebBrowser1.Navigate CommonDialog1.FileName

  isrtfboxloded = True
  RenderDocument
  txtModifySave = False
End Sub

Private Sub FileSave_Click()

    If OpenFileName <> "" Then
        RichTextBox1.SaveFile OpenFileName, 1
    Else
        FileSaveAs_Click
    End If
    
End Sub

Private Sub FileSaveAs_Click()

    CommonDialog1.DefaultExt = "html"
    CommonDialog1.Filter = "HTML Documents|*.html|VIN Files|*.vin|All Files|*.*"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    RichTextBox1.SaveFile CommonDialog1.FileName, 1
    OpenFileName = CommonDialog1.FileName
    HTMLEdit.Caption = "VIN WEB-COMPILER v3.0  ( " & OpenFileName & " )"
    End Sub

Private Sub find_Click()
 SearchForm.ReplaceButton.Visible = False
SearchForm.ReplaceAllButton.Visible = False
SearchForm.Label2.Visible = False
SearchForm.Text2.Visible = False
  SearchForm.Width = SearchForm.ReplaceAllButton.Left
  SetWindowPos SearchForm.hwnd, HWND_TOPMOST, 80, 250, 370, 135, SWP_SHOWWINDOW
End Sub

Private Sub font_Click()
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
RichTextBox1.SelText = Chr(170) & "<FONT COLOR =#FF3300 FACE= ARIAL,HELVATICA SIZE = 5>  " & RichTextBox1.SelText & " </FONT> "
Call MakemeBlue("<FONT COLOR =#FF3300 FACE= ARIAL,HELVATICA SIZE = 5>", "</FONT>", 8)

End Sub

Private Sub Form_GotFocus()
If isRealTime = True Then RenderDocument

End Sub

Private Sub Form_Load()
'WebBrowser1.StatusBar = True
On Error GoTo chupchap
isrtfboxloded = False
txtModifySave = False
'loadnewdocument
' Rang_De
' isrtfboxloded = True
' RenderDocument
FileNew_Click
 isleftarrowpress = False
 isRealTime = True
 EditRender.Checked = True
 vreal.Checked = True
 HTMLEdit.Height = Screen.Height
 HTMLEdit.Width = Screen.Width
 Exit Sub
chupchap:
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 And (y > RichTextBox1.Top + RichTextBox1.Height) And (y < WebBrowser1.Top) Then
        Screen.MousePointer = vbSizeNS
        ResizeWindows = True
        DragStartY = y
        WebHeight = WebBrowser1.Height
        HTMLHeight = RichTextBox1.Height
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    
    If ResizeWindows Then
        RichTextBox1.Height = HTMLHeight + (y - DragStartY)
        WebBrowser1.Move WebBrowser1.Left, RichTextBox1.Top + RichTextBox1.Height + 70, WebBrowser1.Width, WebHeight - (y - DragStartY)
        HTMLEdit.Refresh
    End If
     If (y > RichTextBox1.Top + RichTextBox1.Height) And (y < WebBrowser1.Top) Then
        Screen.MousePointer = vbSizeNS
        End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ResizeWindows = False
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Paint()
'If isRealTime = True Then RenderDocument
End Sub

Private Sub Form_Resize()
On Error GoTo vinerror
If isRealTime = True Then RenderDocument
If win(3).Checked = True Then
    RichTextBox1.Width = HTMLEdit.Width - RichTextBox1.Left - 200
    WebBrowser1.Width = RichTextBox1.Width
    RichTextBox1.Height = 0.3 * HTMLEdit.Height
    WebBrowser1.Move WebBrowser1.Left, RichTextBox1.Top _
    + RichTextBox1.Height + 70, WebBrowser1.Width, HTMLEdit.Height _
    - RichTextBox1.Top - RichTextBox1.Height - 940
    'WebBrowser1.Navigate App.Path & "\empty.htm"

  End If
    Exit Sub
vinerror:
    'MsgBox "An unexpected error occured !! sorry aborting"
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
If txtModifySave = True Then
Dim reply As Integer
reply = MsgBox("Do you want to save the webpage", vbYesNo, "Prompt For Save")
If reply = 6 Then         'yes
  FileSave_Click
End If
End If
 
End Sub







Private Sub help_Click()
MsgBox "1.Whenever you type, the preview of webpage will show below by REAL TIME RENDERING." & Chr(13) & _
"  If you don't see preview ,click on Edit->REAL TIME RENDERING (F6)" & Chr(13) & _
"2.To Learn about HTML, click on File->OpenSampleWebPage.If it will not open" & Chr(13) & _
"  Then Open it manually, by File->Open and then browse the applications default folder\data\sample.html" & Chr(13) & _
"3. To apply the tags, just select your text and click on appropriate menu" & Chr(13)

End Sub



Private Sub hrType_Click(Index As Integer)
 RichTextBox1.SelColor = vbBlue
 If Index = 1 Then
  RichTextBox1.SelText = "<HR width = 80% align = left> " & Chr(13)
 ElseIf Index = 2 Then
   RichTextBox1.SelText = "<HR width = 70% align = center> " & Chr(13)
 ElseIf Index = 3 Then
   RichTextBox1.SelText = "<HR width = 60% align = right> " & Chr(13)
 End If
 RichTextBox1.SelColor = vbBlack
End Sub



Private Sub imgType_Click(Index As Integer)
If Index = 1 Then
 RichTextBox1.SelText = Chr(170) & "<IMG SRC=  ALIGN=LEFT ALT= >  " & RichTextBox1.SelText & " </IMG> "
 Call MakemeBlue("<IMG SRC=  ALIGN=LEFT ALT= >", "</IMG>", 8)
ElseIf Index = 2 Then
     RichTextBox1.SelText = Chr(170) & "<IMG SRC=  ALIGN=center ALT= >  " & RichTextBox1.SelText & " </IMG> "
Call MakemeBlue("<IMG SRC=  ALIGN=center ALT= >", "</IMG>", 8)

ElseIf Index = 3 Then
   
   RichTextBox1.SelText = Chr(170) & "<IMG SRC=  ALIGN=RIGHT ALT= >  " & RichTextBox1.SelText & " </IMG> "
   Call MakemeBlue("<IMG SRC=  ALIGN=RIGHT ALT= >", "</IMG>", 8)

End If
End Sub






Private Sub link_Click()
RichTextBox1.SelText = Chr(170) & "<A HREF= >  " & RichTextBox1.SelText & " </A> "
Call MakemeBlue("<A HREF= >", "</A>", 8)
End Sub






Private Sub NavigateTo_Click()
Dim reply As Integer
reply = MsgBox("Do you want to save the webpage", vbYesNo, "Prompt For Save")
If reply = 6 Then      'yes
  FileSave_Click
End If
isrtfboxloded = False 'prevent rendering because of rtftext change event at time of loading
    URL = InputBox("Enter The Web Address Or URL to navigate to")
    If URL <> "" Then
        WebBrowser1.Navigate URL
    End If
    Rang_De
isrtfboxloded = True
RenderDocument
End Sub





Private Sub pre_Click()
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
RichTextBox1.SelText = Chr(170) & "<PRE>  " & RichTextBox1.SelText & " </PRE> "
Call MakemeBlue("<PRE>", "</PRE>", 6)
End Sub





Private Sub rep_Click()
SearchForm.ReplaceButton.Visible = True
SearchForm.ReplaceAllButton.Visible = True
SearchForm.Label2.Visible = True
SearchForm.Text2.Visible = True
SetWindowPos SearchForm.hwnd, HWND_TOPMOST, 106, 250, 470, 135, SWP_SHOWWINDOW
 
End Sub

Private Sub RichTextBox1_Change()
If isrtfboxloded = True And isRealTime = True Then 'at form loading time called then give a error so first load rtfbox
 RenderDocument
End If
 txtModifySave = True
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
 
tothere = -10

If KeyCode = 188 Then
 If Shift And vbShiftMask Then   'only activate when shift key down for <
  RichTextBox1.SelText = Chr(156) ' "*"
  fromhere = RichTextBox1.find(Chr(156), 0, , 0)
  RichTextBox1.SelText = ""
  isleftarrowpress = True
  End If
End If
If KeyCode = 190 Then
 If Shift And vbShiftMask Then   'only activate when shift key down for >
  If isleftarrowpress = True Then  'if < is already pressed
 
  RichTextBox1.SelText = Chr(156) '"*"
   'Rang_De_limited
  'MsgBox "her"
  tothere = RichTextBox1.find(Chr(156), fromhere - 1, , 0)
 'MsgBox "vif"
  RichTextBox1.SelText = ""
  isleftarrowpress = False
  End If
 End If
End If
If KeyCode = 13 Then
RichTextBox1.SelColor = vbBlue
RichTextBox1.SelFontSize = 12
 RichTextBox1.SelText = "<br>" 'RichTextBox1.SelText
  
 RichTextBox1.SelColor = vbBlack
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

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 And Trim(RichTextBox1.SelText) = "" Then
   PopupMenu adrishya1
   ElseIf Button = 2 Then
   PopupMenu adrishya2
    End If

End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = vbDefault
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
   ' MsgBox "IAM IN SHORTCUT RENDER"
End Sub

Private Sub sample_Click()
If txtModifySave = True Then
Dim reply As Integer
reply = MsgBox("Do you want to save the webpage", vbYesNo, "Prompt For Save")
If reply = 6 Then         'yes
  FileSave_Click
End If
End If

isrtfboxloded = False 'prevent rendering because of rtftext change event at time of loading
On Error GoTo vinerror
RichTextBox1.LoadFile App.Path & "\data\sample.html", rtfText
Rang_De
 WebBrowser1.Navigate App.Path & "\data\sample.html"
isrtfboxloded = True
RenderDocument
txtModifySave = False
Exit Sub
vinerror:
 MsgBox "Sample page not found in defaultdirectory\data\sample.html"
End Sub

Private Sub sblink_Click()
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
RichTextBox1.SelText = Chr(170) & "<BLINK>  " & RichTextBox1.SelText & " </BLINK> "
Call MakemeBlue("<BLINK>", "</BLINK>", 7)
End Sub



Private Sub scripts_Click(Index As Integer)
 If Index = 1 Then
    RichTextBox1.SelText = Chr(170) & "<TEXTAREA NAME = ROWS = 15 COLS = 50 >  " & RichTextBox1.SelText & " </TEXTAREA> "
    Call MakemeBlue("<TEXTAREA NAME = ROWS = 15 COLS = 50 >", "</TEXTAREA>", 10)

 ElseIf Index = 2 Then
   RichTextBox1.SelText = Chr(170) & "<INPUT NAME = TYPE = VALUE = SIZE = 30 MAXLENGTH = ALIGN = JUSTIFY>  " & RichTextBox1.SelText & " </INPUT> "
   Call MakemeBlue("<INPUT NAME = TYPE = VALUE = SIZE = 30 MAXLENGTH = ALIGN = JUSTIFY>", "</INPUT>", 10)

 ElseIf Index = 3 Then
   RichTextBox1.SelText = Chr(170) & "<SELECT NAME =  SIZE = 4 >  " & RichTextBox1.SelText & " </SELECT> "
Call MakemeBlue("<SELECT NAME =  SIZE = 4 >", "</SELECT>", 8)
 End If
End Sub

Private Sub smarq_Click()
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
RichTextBox1.SelText = Chr(170) & "<MARQUEE>  " & RichTextBox1.SelText & " </MARQUEE> "
Call MakemeBlue("<MARQUEE>", "</MARQUEE>", 9)
End Sub





Private Sub smnuStyles_Click(Index As Integer)
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
 If Index = 1 Then
  RichTextBox1.SelUnderline = 1 'submnuUnderline.Checked
    RichTextBox1.SelText = Chr(170) & "<U>  " & RichTextBox1.SelText & " </U> "
   Call MakemeBlue("<U>", "</U>", 4)
   RichTextBox1.SelUnderline = 0
 ElseIf Index = 2 Then
  RichTextBox1.SelText = Chr(170) & "<I>  " & RichTextBox1.SelText & " </I> "
  Call MakemeBlue("<I>", "</I>", 4)
  ElseIf Index = 3 Then
  RichTextBox1.SelText = Chr(170) & "<STRIKE>  " & RichTextBox1.SelText & " </STRIKE> "
  Call MakemeBlue("<STRIKE>", "</STRIKE>", 8)
ElseIf Index = 4 Then
  RichTextBox1.SelFontSize = 13
  RichTextBox1.SelText = Chr(170) & "<BIG>  " & RichTextBox1.SelText & " </BIG> "
  Call MakemeBlue("<BIG>", "</BIG>", 6)
ElseIf Index = 5 Then
  RichTextBox1.SelText = Chr(170) & "<SMALL>  " & RichTextBox1.SelText & " </SMALL> "
  Call MakemeBlue("<SMALL>", "</SMALL>", 6)
 ElseIf Index = 6 Then
  RichTextBox1.SelText = Chr(170) & "<SUP>  " & RichTextBox1.SelText & " </SUP> "
  Call MakemeBlue("<SUP>", "</SUP>", 6)
ElseIf Index = 7 Then
  RichTextBox1.SelText = Chr(170) & "<SUB>  " & RichTextBox1.SelText & " </SUB> "
Call MakemeBlue("<SUB>", "</SUB>", 6)
ElseIf Index = 8 Then
  RichTextBox1.SelText = Chr(170) & "<STRONG>  " & RichTextBox1.SelText & " </STRONG> "
Call MakemeBlue("<STRONG>", "</STRONG>", 8)
ElseIf Index = 9 Then
 RichTextBox1.SelText = Chr(170) & "<EM>  " & RichTextBox1.SelText & " </EM> "
Call MakemeBlue("<EM>", "</EM>", 5)
ElseIf Index = 10 Then
 RichTextBox1.SelText = Chr(170) & "<DFN>  " & RichTextBox1.SelText & " </DFN> "
Call MakemeBlue("<DFN>", "</DFN>", 5)
ElseIf Index = 11 Then
 RichTextBox1.SelText = Chr(170) & "<CODE>  " & RichTextBox1.SelText & " </CODE> "
 Call MakemeBlue("<CODE>", "</CODE>", 7)
ElseIf Index = 12 Then
  RichTextBox1.SelText = Chr(170) & "<SAMP>  " & RichTextBox1.SelText & " </SAMP> "
  Call MakemeBlue("<SAMP>", "</SAMP>", 7)
ElseIf Index = 13 Then
  RichTextBox1.SelText = Chr(170) & "<VAR>  " & RichTextBox1.SelText & " </VAR> "
   Call MakemeBlue("<VAR>", "</VAR>", 6)
ElseIf Index = 14 Then
  RichTextBox1.SelText = Chr(170) & "<ACRONYM>  " & RichTextBox1.SelText & " </ACRONYM> "
Call MakemeBlue("<ACRONYM>", "</ACRONYM>", 10)
ElseIf Index = 15 Then
 RichTextBox1.SelText = Chr(170) & "<ABBR>  " & RichTextBox1.SelText & " </ABBR> "
Call MakemeBlue("<ABBR>", "</ABBR>", 5)
End If

End Sub


Private Sub ssmnuHeader_Click(Index As Integer)
    
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
End If
If Index = 1 Then
RichTextBox1.SelText = Chr(170) & "<H1>  " & RichTextBox1.SelText & " </H1> "
Call MakemeBlue("<H1>", "</H1>", 5)

ElseIf Index = 2 Then
 RichTextBox1.SelText = Chr(170) & "<H2>  " & RichTextBox1.SelText & " </H2> "
 Call MakemeBlue("<H2>", "</H2>", 5)
ElseIf Index = 3 Then
 RichTextBox1.SelText = Chr(170) & "<H3>  " & RichTextBox1.SelText & " </H3> "
 Call MakemeBlue("<H3>", "</H3>", 5)
ElseIf Index = 4 Then
 RichTextBox1.SelText = Chr(170) & "<H4>  " & RichTextBox1.SelText & " </H4> "
 Call MakemeBlue("<H4>", "</H4>", 5)
ElseIf Index = 5 Then
 RichTextBox1.SelText = Chr(170) & "<H5>  " & RichTextBox1.SelText & " </H5> "
 Call MakemeBlue("<H5>", "</H5>", 5)
ElseIf Index = 6 Then
  RichTextBox1.SelText = Chr(170) & "<H6>  " & RichTextBox1.SelText & " </H6> "
  Call MakemeBlue("<H6>", "</H6>", 5)
End If


End Sub

Private Sub ssmnuOthers_Click(Index As Integer)
RichTextBox1.SelColor = &H59C6
 If Index = 0 Then
  RichTextBox1.SelText = "&szlig;"
  ElseIf Index = 1 Then RichTextBox1.SelText = "&oslash;"
  ElseIf Index = 2 Then RichTextBox1.SelText = "&aelig;"
  ElseIf Index = 3 Then RichTextBox1.SelText = "&eth;"
  ElseIf Index = 4 Then RichTextBox1.SelText = "&thorn;"
  ElseIf Index = 5 Then RichTextBox1.SelText = "&oelig;"
  ElseIf Index = 6 Then RichTextBox1.SelText = "&Aelig;"
  ElseIf Index = 7 Then RichTextBox1.SelText = "&Eth;"
  ElseIf Index = 8 Then RichTextBox1.SelText = "&Oslash;"
  ElseIf Index = 9 Then RichTextBox1.SelText = "&THORN;"
End If
RichTextBox1.SelColor = vbBlack
  
End Sub

Private Sub ssmnuSymbols_Click(Index As Integer)
RichTextBox1.SelColor = &H59C6
 If Index = 0 Then
  RichTextBox1.SelText = "&nbsp;"
  ElseIf Index = 1 Then RichTextBox1.SelText = "&copy;"
 ElseIf Index = 2 Then RichTextBox1.SelText = "&reg;"
 ElseIf Index = 3 Then RichTextBox1.SelText = "&sup1;"
 ElseIf Index = 4 Then RichTextBox1.SelText = "&sup2;"
 ElseIf Index = 5 Then RichTextBox1.SelText = "&sup3;"
 ElseIf Index = 6 Then RichTextBox1.SelText = "&frac12;"
 ElseIf Index = 7 Then RichTextBox1.SelText = "&frac14;"
 ElseIf Index = 8 Then RichTextBox1.SelText = "&frac34;"
 ElseIf Index = 9 Then RichTextBox1.SelText = "&micro;"
 ElseIf Index = 10 Then RichTextBox1.SelText = "&para;"
 ElseIf Index = 11 Then RichTextBox1.SelText = "&middot;"
 ElseIf Index = 12 Then RichTextBox1.SelText = "&yen;"
 ElseIf Index = 13 Then RichTextBox1.SelText = "&pound;"
 ElseIf Index = 14 Then RichTextBox1.SelText = "&cent;"
 ElseIf Index = 15 Then RichTextBox1.SelText = "&ixecl;"
 ElseIf Index = 16 Then RichTextBox1.SelText = "&curren;"
 ElseIf Index = 17 Then RichTextBox1.SelText = "&brvbar;"
 ElseIf Index = 18 Then RichTextBox1.SelText = "&sect;"
 ElseIf Index = 19 Then RichTextBox1.SelText = "&uml;"
 ElseIf Index = 20 Then RichTextBox1.SelText = "&ordf;"
 ElseIf Index = 21 Then RichTextBox1.SelText = "&laquo;"
 ElseIf Index = 22 Then RichTextBox1.SelText = "&not;"
 ElseIf Index = 23 Then RichTextBox1.SelText = "&macr;"
 ElseIf Index = 24 Then RichTextBox1.SelText = "&deg;"
 ElseIf Index = 25 Then RichTextBox1.SelText = "&plusmn;"
 ElseIf Index = 26 Then RichTextBox1.SelText = "&accute;"
 ElseIf Index = 27 Then RichTextBox1.SelText = "&iquest;"
 ElseIf Index = 28 Then RichTextBox1.SelText = "&times;"
 ElseIf Index = 29 Then RichTextBox1.SelText = "&ordm;"
 ElseIf Index = 30 Then RichTextBox1.SelText = "&raquo;"
End If
RichTextBox1.SelColor = vbBlack
End Sub

Private Sub ssoun_Click()
RichTextBox1.SelText = Chr(170) & "<BGSOUND SRC=" & Chr(34) & "c:\kasmir.mid" & Chr(34) & " LOOP= >  " & RichTextBox1.SelText & " </BGSOUND> "
Call MakemeBlue("<BGSOUND SRC=" & Chr(34) & "c:\kasmir.mid" & Chr(34) & " LOOP= >", "</BGSOUND>", 9)
End Sub







Private Sub submnuBold_Click()
If Trim(RichTextBox1.SelText) = "" Then
 MsgBox "Please first select any sentance "
 Exit Sub
End If
'Dim bymenuhere As Long
'   submnuBold.Checked = Not submnuBold.Checked
    RichTextBox1.SelBold = 1 'submnuBold.Checked
   RichTextBox1.SelText = Chr(170) & "<B>  " & RichTextBox1.SelText & " </B> "
   Call MakemeBlue("<B>", "</B>", 4)
   'bymenuhere = RichTextBox1.Find(Chr(170), 0, , 0)
  'RichTextBox1.SelText = ""        'delete chr170
  'bymenuhere = RichTextBox1.Find("<B>", bymenuhere, , 0)
  'RichTextBox1.SelColor = vbBlue
  'bymenuhere = RichTextBox1.Find("</B>", bymenuhere + 5, , 0)
  'RichTextBox1.SelColor = vbBlue
  'RichTextBox1.SelStart = bymenuhere + 4
  'RichTextBox1.SelLength = 1
  RichTextBox1.SelBold = 0
End Sub



Private Sub submnuPara_Click()
If Trim(RichTextBox1.SelText) = "" Then
  MsgBox "Please first select any sentance "
  Exit Sub
 End If
RichTextBox1.SelText = Chr(170) & "<P>  " & RichTextBox1.SelText & " </P> "
Call MakemeBlue("<P>", "</P>", 4)
End Sub





















Private Sub Timer1_Timer()
'If prevfromhere = fromhere Then

'Else
 If tothere > fromhere Then

  RichTextBox1.SelStart = fromhere
  RichTextBox1.SelLength = tothere - fromhere + 1
  'RichTextBox1.SelText = "" 'RichTextBox1.SelText
  
 RichTextBox1.SelColor = vbBlue
 'MsgBox tothere & " DFDD" & fromhere
   'prevent after > to made blue
   RichTextBox1.SelStart = tothere
   RichTextBox1.SelLength = 1   'SELECT ANY BLANK SPACE AND SET TO BLACK
      RichTextBox1.SelText = RichTextBox1.SelText '& Chr(32)
      RichTextBox1.SelBold = 0
      RichTextBox1.SelColor = vbBlack
      RichTextBox1.SelText = "   "
 End If
 tothere = 0   'do not overblue again
 'prevfromhere = fromhere
'End If
'RenderDocument
End Sub







Private Sub update_Click()
RenderDocument
End Sub

Private Sub vanch_Click()
anchor_Click
End Sub



Private Sub vbgc_Click()
bgcolor_Click
End Sub

Private Sub vblock_Click()
Blockquote_Click
End Sub

Private Sub vbold_Click()
submnuBold_Click
End Sub

Private Sub vcopy_Click()
EditCopy_Click
End Sub

Private Sub vcut_Click()
EditCut_Click
End Sub

Private Sub vdo_Click()
RichTextBox1.SelText = Chr(170) & "<A HREF=" & Chr(34) & "c:\sonu.avi" & Chr(34) & ">" & RichTextBox1.SelText & " </A> "
Call MakemeBlue("<A HREF=" & Chr(34) & "c:\sonu.avi" & Chr(34) & ">", "</A>", 8)

End Sub

Private Sub vfont_Click()
font_Click
End Sub


Private Sub vHeader_Click(Index As Integer)
 ssmnuHeader_Click (Index)
End Sub


Private Sub vhrType_Click(Index As Integer)
hrType_Click (Index)
End Sub



Private Sub vimgType_Click(Index As Integer)
   imgType_Click (Index)
End Sub

Private Sub vin_Click()
Dim temp As Long
On Error GoTo Exeerror
temp = Shell("vin_utility.exe", vbNormalFocus)
Exit Sub
Exeerror:
 MsgBox "Application 'vinutility.EXE' is not found in its " _
  & "Default directory vinutility.exe "
End Sub



Private Sub vlink_Click()
link_Click
End Sub

Private Sub vmark_Click()
smarq_Click
End Sub

Private Sub vpara_Click()
submnuPara_Click
End Sub

Private Sub vpaste_Click()
EditPaste_Click
End Sub

Private Sub vpre_Click()
pre_Click
End Sub



Private Sub vreal_Click()
vreal.Checked = Not vreal.Checked
EditRender_Click
End Sub



Private Sub vscripts_Click(Index As Integer)
 scripts_Click (Index)
End Sub

Private Sub vsound_Click()
ssoun_Click
End Sub



Private Sub vtxtcol_Click()
wincolor_Click
End Sub

Private Sub vunder_Click()
smnuStyles_Click (1)
End Sub

Private Sub vvedio_Click()
vdo_Click
End Sub

Private Sub WebBrowser1_OnStatusBar(ByVal StatusBar As Boolean)
WebBrowser1.StatusBar = True
End Sub

Private Sub WebBrowser1_Validate(Cancel As Boolean)
Debug.Print "Error"

End Sub
Private Sub MakeBlack(lamba As Long)
 RichTextBox1.SelColor = vbBlue
 'LAMBA = 2 FOR 1CHAR  = 3 FOR 2CHAR  = 4 FOR 3CHAR
      RichTextBox1.SelStart = Position + lamba
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
  Position = 1
Dim nbsppos As Long
Do While Position > 0      'exit when < not found
 Position = RichTextBox1.find("<", Position + 1, , 0)
 
  If Position < 0 Or RichTextBox1.SelColor = vbBlue Then
   ' MsgBox "LESS"
    Exit Do          'exit when < not found
  End If
 L = RichTextBox1.find(">", Position + 1, , 0)
 RichTextBox1.SelStart = Position
 RichTextBox1.SelLength = L - Position + 1
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
    DoEvents
  Loop
  'change the color of &nbsp
  nbsppos = 20
  Do While nbsppos > 0
   nbsppos = RichTextBox1.find("&nbsp;", nbsppos + 1, , 0)
    RichTextBox1.SelFontSize = 10
    RichTextBox1.SelColor = vbRed
    DoEvents
  Loop
  'change color of <html>
    nbsppos = RichTextBox1.find("</html>", L - 40, , 0)
    RichTextBox1.SelFontSize = 14
    RichTextBox1.SelColor = vbRed
    nbsppos = RichTextBox1.find("<html>", 1, , 0)
    RichTextBox1.SelFontSize = 14
    RichTextBox1.SelColor = vbRed
    Dim tmp As Integer
    tmp = RichTextBox1.find("VIN WEB-COMPILER", 1, , 0)     'tmp used temporarily
    If tmp < 1 Then RichTextBox1.SelText = "<!--DocType HTML BY VIN WEB-COMPILER copywrite 2003 Vinod Kotiya All rights unreserved --> " & Chr(13) & RichTextBox1.SelText
    'MsgBox tmp
    nbsppos = RichTextBox1.find("Type keyword here", 1, , 0)
    RichTextBox1.SelColor = vbBlack
    nbsppos = RichTextBox1.find("Type Discription Here", 1, , 0)
    RichTextBox1.SelColor = vbBlack
    nbsppos = RichTextBox1.find("your name here", 1, , 0)
    RichTextBox1.SelColor = vbBlack
     nbsppos = RichTextBox1.find("<HEAD>", 1, , 0)
    RichTextBox1.SelColor = vbMagenta
   nbsppos = RichTextBox1.find("</HEAD>", 1, , 0)
    RichTextBox1.SelColor = vbMagenta
    nbsppos = RichTextBox1.find("start typing from here inside body and apply tags by selecting and right clicking", nbspos, , 0)
    RichTextBox1.SelColor = vbRed
     nbsppos = RichTextBox1.find("</BODY>", 1, , 0)
    Exit Sub
HI:
'MsgBox Position & "eRROR"

End Sub

Private Sub Rang_De_limited()
Dim starpos As Long
Position = 1
Do While Position > 0 And Position < starpos + 2    'exit when < not found
 Position = RichTextBox1.find("<", Position + 1, , 0)
 
 If Position < 0 Or RichTextBox1.SelColor = vbBlue Then
 'MsgBox "LESS"
   Exit Do          'exit when < not found
 End If
 L = RichTextBox1.find(">", Position + 1, , 0)
 RichTextBox1.SelStart = Position
 RichTextBox1.SelLength = L - Position + 1
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
Private Sub loadnewdocument()
HTMLEdit.Caption = "VIN WEB-COMPILER v3.0  : " & "New webpage"
RichTextBox1.Text = Chr(13) & "<HTML> " & Space(100) & Chr(13) & "<HEAD><TITLE>Enter your Title here *Compulsory </TITLE>" & vbCrLf & _
"<!--  Meta tags help  search engines to identify the content of your page.       -->" & vbCrLf & _
" <META http-equiv=Content-Type content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">" & vbCrLf & _
 "<META content=" & Chr(34) & "Type keyword here" & Chr(34) & " name=keywords>" & vbCrLf & _
 "<META content=" & Chr(34) & "Type Discription Here" & Chr(34) & " name=description>" & vbCrLf & _
"<META content=" & Chr(34) & "your name here" & Chr(34) & " name=author>" & vbCrLf & _
 "<META content=" & Chr(34) & "MSHTML 6.00.2600.0" & Chr(34) & " name=GENERATOR>" & vbCrLf & vbCrLf & _
 "<!-- Enter your script here (if any)---> " & vbCrLf & vbCrLf & _
"</HEAD>" & vbCrLf & _
"<BODY BGCOLOR = #d9e7ff TEXT = blue ALINK = GREEN VLINK = RED  topmargin=0 marginwidth=0 marginheight=0   >" & vbCrLf & _
 "<!-- start typing from here inside body and apply tags by selecting and right clicking -->" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
     "</BODY>" & vbCrLf & "</HTML>"
     
     'RenderDocument   '''giving an error
     
End Sub
Private Sub MakemeBlue(head As String, tail As String, lamba As Integer)
 Dim bymenuhere As Integer
 bymenuhere = RichTextBox1.find(Chr(170), 0, , 0)
  RichTextBox1.SelText = ""        'delete chr170
  bymenuhere = RichTextBox1.find(head, bymenuhere, , 0)
  RichTextBox1.SelColor = vbBlue
  bymenuhere = RichTextBox1.find(tail, bymenuhere + lamba, , 0)
  RichTextBox1.SelColor = vbBlue
  RichTextBox1.SelStart = bymenuhere + lamba    'make a blank space to default
  RichTextBox1.SelLength = 1
  
  
End Sub

Private Sub win_Click(Index As Integer)
win(Index).Checked = Not win(Index).Checked
 If Index = 3 Then   'both
  win(1).Checked = False
  win(2).Checked = False
   RichTextBox1.Width = HTMLEdit.Width - RichTextBox1.Left - 200
    WebBrowser1.Width = RichTextBox1.Width
    RichTextBox1.Height = 0.3 * HTMLEdit.Height
    WebBrowser1.Move WebBrowser1.Left, RichTextBox1.Top _
    + RichTextBox1.Height + 70, WebBrowser1.Width, HTMLEdit.Height _
    - RichTextBox1.Top - RichTextBox1.Height - 940
 ElseIf Index = 1 Then   'source
    win(2).Checked = False
    win(3).Checked = False
    RichTextBox1.Height = HTMLEdit.ScaleHeight
    WebBrowser1.Move WebBrowser1.Left, RichTextBox1.Top + RichTextBox1.Height, WebBrowser1.Width, 0
        HTMLEdit.Refresh
 ElseIf Index = 2 Then 'web
    win(1).Checked = False
    win(3).Checked = False
    RichTextBox1.Height = 0
        WebBrowser1.Move WebBrowser1.Left, RichTextBox1.Top, WebBrowser1.Width, HTMLEdit.ScaleHeight
        HTMLEdit.Refresh
 End If
 
End Sub

Private Sub wincolor_Click()
Dim CDFlags As Long
Dim Lal As Integer, Hara As Integer, Nila As Integer
Dim Rang As Long
Dim hcode As String
On Error GoTo ColorError

    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)

    CommonDialog1.Flags = CDFlags
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    
    Rang& = CommonDialog1.Color
    Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    hcode = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)
    Dim startpos As Integer
    startpos = RichTextBox1.find("bgcolor", , , 0)
    Rang = RichTextBox1.find("TEXT", startpos, , 0)
    CDFlags = RichTextBox1.find("ALINK", Rang, Rang + 30, 0)
    RichTextBox1.SelStart = Rang + 7
    RichTextBox1.SelLength = CDFlags - Rang - 8

    RichTextBox1.SelText = hcode
    Exit Sub
    
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
    
    Else
        MsgBox "An error occured"
    End If

End Sub

