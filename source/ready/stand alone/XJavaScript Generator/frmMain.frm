VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Script Generator v2.0"
   ClientHeight    =   7395
   ClientLeft      =   150
   ClientTop       =   900
   ClientWidth     =   7095
   HelpContextID   =   5
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Code and Options"
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   6855
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H0000FF00&
         Caption         =   "Embedded the code in any existing webpage"
         Height          =   495
         Index           =   3
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H0000FF00&
         Caption         =   "Show output as webpage"
         Height          =   495
         Index           =   2
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H00FF80FF&
         Caption         =   "Save"
         Height          =   495
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H0080FFFF&
         Caption         =   "Copy"
         Height          =   495
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2280
         Width           =   735
      End
      Begin RichTextLib.RichTextBox rtfCode 
         Height          =   1935
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3413
         _Version        =   393217
         BackColor       =   16777152
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":1CCA
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Choose Any Script"
      ForeColor       =   &H0000FF00&
      Height          =   3255
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox txtDetail 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Shows the discription of Script"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.ListBox listScripts 
         BackColor       =   &H00C0FFFF&
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
         Height          =   2205
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Please Enter the Values of variables"
      ForeColor       =   &H0000FF00&
      Height          =   3255
      Left            =   3120
      TabIndex        =   19
      Top             =   1080
      Width           =   3855
      Begin VB.CommandButton cmdCode 
         BackColor       =   &H0000FF00&
         Caption         =   ": : Generate or Update the Code : :"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Type the value of variable shown above."
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Label lblInput 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblInput 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblInput 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblInput 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.Frame frmCat 
      BackColor       =   &H00000000&
      Caption         =   "Category"
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Select any one category then choose any script from the list"
      Top             =   0
      Width           =   6855
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Various"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Image and Sound"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Status Bar"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Text"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Links and Menus"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Date and Time"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Backgrounds"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Mouse Cursor"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Menu adder 
      Caption         =   "&Script Adder"
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu helper 
         Caption         =   "Help"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu helper 
         Caption         =   "About"
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu helper 
         Caption         =   "About Me"
         Index           =   2
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
'///////////////// VIN Script Generator v1.0 ///////////////////////////////////
'//////////        Created By : - VINOD KOTIYA             ///////////////////////////
'/////////          Created On:- 09-06-2003 to 10-06-2003 ////////////////////
'/////////          Time :- 23:00 PM to 2:00 AM    ///////////////////////////////
'/////////          Total Hours :- 3hr.         ///////////////////////////////////////
'/////////          Time for arranging data files :- 14:30 PM to 16:00 PM ///////
'/////////          Proudly releasing version 1.0 ////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////
'////           FIRST MODIFICATION :- Making user freindly so that  / /////
'////           user can update the data files //////////////////////////////
'////           Proudly releasing version 2.0 ///////////////////////////////
'////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////


Option Explicit


Private Sub adder_Click()
Load frmAdder
frmAdder.Visible = True
End Sub

Private Sub cmdCode_Click()
'original script is in temple.vin.inside it the variables assigned the value input1 2 3 4 etc
'the values given by user will replace the input1 2 3 4

rtfCode.LoadFile App.Path & "\temple .vin", rtfText

'so find the input1234 and replace them with value given by user inside txtInput1 2 3 4

Dim pos As Long
Dim i As Integer

For i = 0 To 3
     pos = rtfCode.Find("input" & i + 1, pos + 1, , 0)
  Do While pos > 0
              rtfCode.SelText = txtInput(i).Text
              pos = rtfCode.Find("input" & i + 1, pos + 1, , 0)
              
   Loop
 Next
End Sub

Private Sub cmdOption_Click(Index As Integer)
  Dim txtTemp As String
  Dim pos As Long
 If Index = 0 Then                   'copy button
  Clipboard.SetText rtfCode.Text
 ElseIf Index = 1 Then             'SAVE Button
    CommonDialog1.DefaultExt = "vin"
    CommonDialog1.Filter = "VIN Files|*.vin|All Files|*.*"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    rtfCode.SaveFile CommonDialog1.FileName, 1
 ElseIf Index = 2 Then       'Show output as webpage
  CommonDialog1.Filter = "HTML Files|*.html"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    Dim Fnum As Integer
     txtTemp = "<!---DocType HTML VINOD KOTIYA s-2 Shrimaya Apartment Sector-B/363 Sarvdharm Colony Bhopal-42.Fone +91-0755-2794428 --->" & _
      vbCrLf & "<HTML>" & vbCrLf & "      <HEAD><TITLE> VIN Script Generator </TITLE>" & vbCrLf
    Fnum = FreeFile
    Open CommonDialog1.FileName For Output As Fnum
     Print #Fnum, txtTemp
     Print #Fnum, rtfCode.Text
     txtTemp = "</HEAD> " & vbCrLf & "     <BODY>               </BODY> " & vbCrLf & " </HTML>"
     Print #Fnum, txtTemp
     Close Fnum        ''=====================closing the output file
  ElseIf Index = 3 Then                 ''Embedded code in existing file
    CommonDialog1.Filter = "HTML Files|*.html;*.htm | All Files | *.*"
    CommonDialog1.ShowSave
     If CommonDialog1.FileName = "" Then Exit Sub
     txtTemp = rtfCode.Text
     rtfCode.LoadFile CommonDialog1.FileName, rtfText
     pos = 0
     pos = rtfCode.Find("/head", pos + 1, , 0)
     rtfCode.SelStart = pos - 2
     rtfCode.SelText = txtTemp
     rtfCode.SaveFile CommonDialog1.FileName, 1
 End If
End Sub

Private Sub helper_Click(Index As Integer)
'CommonDialog1.HelpFile = "C:\Documents and Settings\raj1\Desktop\all2html.chm"
'CommonDialog1.ShowHelp
End Sub

Private Sub listScripts_Click()

Dim Fnum As Integer   ' Descriptor for file.
Dim currentLine As String       'in loop store currentline of selected file
Dim txtloadrtfCode As String
Dim i As Integer
For i = 0 To 3                   ''Refresh the input label and text box
  lblInput(i).Enabled = True
  txtInput(i).Enabled = True
  lblInput(i).Caption = "Variable " & i + 1
  txtInput(i).Text = " "
 Next
rtfCode.Text = " "
cmdCode.Enabled = True
'txtInput(0).Text = listScripts.List(listScripts.ListIndex)

Fnum = FreeFile
Open App.Path & "\vinscript\" & listScripts.List(listScripts.ListIndex) & ".vin" For Input As Fnum
       Line Input #Fnum, currentLine         'getting 1st line
       
       If Trim(currentLine) = "<inputYes>" Then   '!!!!!!!!!!!!!!!!! Script has user variables
                                                                                    '!!!!!!!!!!!!!!!! hence it need to be updated
         For i = 0 To 3
            Line Input #Fnum, currentLine      'nil
            Line Input #Fnum, currentLine     'input label
            If Trim(currentLine) <> "no" Then     'i.e. label is not no
              lblInput(i).Caption = currentLine
              Line Input #Fnum, currentLine     'input text box
              txtInput(i).Text = currentLine
            Else
             Line Input #Fnum, currentLine      'nil i.e. label was no so text is also no so disable them
              lblInput(i).Enabled = False
              txtInput(i).Enabled = False
            End If
         Next
           
           Line Input #Fnum, currentLine    '<discription>
           txtDetail.Text = currentLine
           Line Input #Fnum, currentLine   'input discription
           txtDetail.Text = txtDetail.Text & "    " & currentLine           'vbcrlf
           Line Input #Fnum, currentLine 'nil <end By vinod kotiya >
           
          While Not EOF(Fnum)              'now original script part is storing in to txtloadrtfCode
              Line Input #Fnum, currentLine
              txtloadrtfCode = txtloadrtfCode & vbCrLf & currentLine
           Wend
        'Input #fnum, txtloadrtfCode
            Close Fnum          '=====================closing the input file when if
           Fnum = FreeFile
           Open App.Path & "\temple .vin" For Output As Fnum
           Print #Fnum, txtloadrtfCode     'for frequently used original script is saved in temple.vin
            Close Fnum        ''=====================closing the output file
            txtloadrtfCode = " "                  'freeing memory
          '  rtfCode.Text = txtloadrtfCode
    Else        '!!!!!!!!!!!!!!!!!!!       'i.e. the script has no variable setting so get 1st line discription and directly show
                                                    'the code .disable input variables and generate code
                                           '!!!!!!!!!!!!!!!! hence it need not to be updated
           txtDetail.Text = currentLine
           Close Fnum           '=====================closing the input file when else
        For i = 0 To 3
         lblInput(i).Enabled = False
         txtInput(i).Enabled = False
        Next
        cmdCode.Enabled = False
               
          rtfCode.LoadFile App.Path & "\vinscript\" & listScripts.List(listScripts.ListIndex) & ".vin", rtfText
   End If       'end of <inputYes>
End Sub

Private Sub optCategory_Click(Index As Integer)
listScripts.Clear
If Index = 0 Then
 
 listScripts.AddItem "Elastic Trail"
 listScripts.AddItem "Image Mouse Stars"
 listScripts.AddItem "Logo Orbit"
 listScripts.AddItem "Mouse Comet"
 listScripts.AddItem "Mouse Star"
 listScripts.AddItem "Sparkle Trail"
 listScripts.AddItem "Text Trail"
 listScripts.AddItem "Trio2"
ElseIf Index = 1 Then
 
 listScripts.AddItem "Fireworks"
 listScripts.AddItem "Snow"
 ElseIf Index = 2 Then
  listScripts.AddItem "Alarm Clock"
  listScripts.AddItem "Count Down"
  listScripts.AddItem "Current date1"
  listScripts.AddItem "Current date2"
  listScripts.AddItem "Current date3"
  listScripts.AddItem "Current date4"
  listScripts.AddItem "Current date5"
  listScripts.AddItem "DateTime"
  listScripts.AddItem "DateTimeAlert"
  listScripts.AddItem "GMT Clock"
  listScripts.AddItem "Simple Clock"
  listScripts.AddItem "Title Date and Time"
  listScripts.AddItem "World Time"
ElseIf Index = 3 Then
 listScripts.AddItem "Email Stalker"
 listScripts.AddItem "Link Stalker"
 listScripts.AddItem "No Click Link"
 listScripts.AddItem "Print Button"
 listScripts.AddItem "Reload"
 listScripts.AddItem "Button Popup Menu"
 listScripts.AddItem "List Box Menu"
 listScripts.AddItem "List Box Menu2"
 listScripts.AddItem "List Box Menu3"
 listScripts.AddItem "Simple NavBar"
ElseIf Index = 4 Then
 listScripts.AddItem "Superscroll"
 listScripts.AddItem "Scrolling Text"
 listScripts.AddItem "Zooming Text"
 listScripts.AddItem "Telex Type"
 listScripts.AddItem "Popup Message"
ElseIf Index = 5 Then
 listScripts.AddItem "Text Scroll"
 listScripts.AddItem "Typing Message"
 listScripts.AddItem "Clock"
 listScripts.AddItem "Flying Text"
 listScripts.AddItem "Bubble Text"
ElseIf Index = 6 Then
listScripts.AddItem "Bouncing Image"
listScripts.AddItem "image animator"
listScripts.AddItem "Mouse Over Button"
ElseIf Index = 7 Then
 listScripts.AddItem "Drive Viewer"
 listScripts.AddItem "Horoscope Sign"
 listScripts.AddItem "Length Converter"
 listScripts.AddItem "Temperature Converter"
 listScripts.AddItem "Web Search"
 
End If
End Sub
