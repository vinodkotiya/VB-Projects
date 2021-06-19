VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Script Generator v2.0"
   ClientHeight    =   6915
   ClientLeft      =   2565
   ClientTop       =   1215
   ClientWidth     =   6315
   HelpContextID   =   5
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6315
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Code and Options"
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   90
      TabIndex        =   26
      Top             =   4200
      Width           =   6135
      Begin VB.CommandButton cmdPreview 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preview"
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H0000FF00&
         Caption         =   "Embedded the code in any existing webpage"
         Height          =   495
         Index           =   3
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H0000FF00&
         Caption         =   "Save as webpage"
         Height          =   495
         Index           =   2
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Save"
         Height          =   495
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdOption 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Copy"
         Height          =   495
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Width           =   495
      End
      Begin RichTextLib.RichTextBox rtfCode 
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2990
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
      Left            =   90
      TabIndex        =   24
      Top             =   960
      Width           =   2295
      Begin VB.TextBox txtDetail 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Shows the discription of Script"
         Top             =   2520
         Width           =   2055
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
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Please Enter the Values of variables"
      ForeColor       =   &H0000FF00&
      Height          =   3255
      Left            =   2400
      TabIndex        =   19
      Top             =   960
      Width           =   3855
      Begin VB.CommandButton cmdCode 
         BackColor       =   &H0000FF00&
         Caption         =   ": : Generate or Update the Code : :"
         Enabled         =   0   'False
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
      Height          =   855
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Select any one category then choose any script from the list"
      Top             =   0
      Width           =   6135
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Various"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Image and Sound"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Status Bar"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Text"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00000000&
         Caption         =   "Links and Menus"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   4440
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
         Left            =   2760
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
         Left            =   1440
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
         Left            =   120
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
Dim writable As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


Private Sub adder_Click()
Load frmAdder
frmAdder.Visible = True
End Sub

Private Sub cmdCode_Click()
'original script is in temple.vin.inside it the variables assigned the value input1 2 3 4 etc
'the values given by user will replace the input1 2 3 4
If writable Then
  rtfCode.LoadFile App.Path & "\temple .vin", rtfText
Else
  rtfCode.LoadFile "c:\temple .vin", rtfText
End If

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
If Trim(rtfCode.Text) = "" Then
  MsgBox "Please First Click on Generate Code Button"
  Exit Sub
End If
  
  On Error GoTo vinerror
  Dim txttemp As String
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
     txttemp = "<!---DocType HTML VINOD KOTIYA s-2 Shrimaya Apartment Sector-B/363 Sarvdharm Colony Bhopal-42.Fone +91-0755-2794428 --->" & _
      vbCrLf & "<HTML>" & vbCrLf & "      <HEAD><TITLE> VIN Script Generator </TITLE>" & vbCrLf
    Fnum = FreeFile
    Open CommonDialog1.FileName For Output As Fnum
     Print #Fnum, txttemp
     Print #Fnum, rtfCode.Text
     txttemp = "</HEAD> " & vbCrLf & "     <BODY bgcolor = #0000ff text = #ffff00>               </BODY> " & vbCrLf & " </HTML>"
     Print #Fnum, txttemp
     Close Fnum        ''=====================closing the output file
  ElseIf Index = 3 Then                 ''Embedded code in existing file
    CommonDialog1.Filter = "HTML Files|*.html;*.htm | All Files | *.*"
    CommonDialog1.ShowOpen
     If CommonDialog1.FileName = "" Then Exit Sub
     txttemp = rtfCode.Text
     rtfCode.LoadFile CommonDialog1.FileName, rtfText
     pos = 0
     pos = rtfCode.Find("/head", pos + 1, , 0)
     rtfCode.SelStart = pos - 2
     rtfCode.SelText = txttemp
     '' used temporarly for my site delete this/////
   '  pos = rtfCode.Find("/BODY", pos + 1, , 0)
   '  rtfCode.SelStart = pos - 2
   '  rtfCode.SelText = "<H4>" & listScripts.List(listScripts.ListIndex) & "<H4>" & vbCrLf & _
      "<h6>Third party scripts</h6>" & vbCrLf & _
      "<center><h3>1 : To get the script click on view---->> Source </h3></center>" & vbCrLf & _
      "<center><h3>The html coding will appear  </h3></center>" & vbCrLf & _
      "<center><h3>Copy the script Source placed between HEAD tag and paste on your page </h3></center>" & vbCrLf & _
      "<center><h3>2 :  Or save this page on your harddisk by File->> SaveAs  </h3></center>" & vbCrLf & vbCrLf & vbCrLf & _
      "<font color = red size = 4> To get  more scripts,source code,viruses,softwares contact at vinodkotiya24@rediffmail.com </font>" & vbCrLf & vbCrLf & _
      "<hr><font color = white > <center><a href =" & Chr(34) & "..\index.html" & Chr(34) & "> HOME  </a>| <a href =" & Chr(34) & "..\source.html" & Chr(34) & "> BACK </a>" & _
      " | <a href =" & Chr(34) & "..\stand.html" & Chr(34) & "> STAND ALONES </a>| <a href =" & Chr(34) & "..\aboutme.html" & Chr(34) & "> About </a> </center></font><hr>"
     
   '  If Len(listScripts.List(listScripts.ListIndex)) > 8 Then     'If "F:\" = 3
   '     CommonDialog1.FileName = LCase(Left(listScripts.List(listScripts.ListIndex), 8))   'if "F:\" it return "F:"
   ' Else
   '  CommonDialog1.FileName = LCase(listScripts.List(listScripts.ListIndex))
   '  End If
     '//////////////////////////
     
     CommonDialog1.Filter = "HTML Files|*.html;*.htm | All Files | *.*"
    CommonDialog1.ShowSave
     If CommonDialog1.FileName = "" Then
        rtfCode.Text = ""
        Exit Sub
    End If
   
     rtfCode.SaveFile CommonDialog1.FileName, 1
 End If
 Exit Sub
vinerror:
 MsgBox "unable to write on disk "
End Sub

Private Sub cmdPreview_Click()
If Trim(rtfCode.Text) = "" Then
  MsgBox "Please First Click on Generate Code Button"
  Exit Sub
End If
Dim Fnum As Integer
 Dim txttemp As String
Dim strDir As String
     txttemp = "<!---DocType HTML VINOD KOTIYA s-2 Shrimaya Apartment Sector-B/363 Sarvdharm Colony Bhopal-42.Fone +91-0755-2794428 --->" & _
      vbCrLf & "<HTML>" & vbCrLf & "      <HEAD><TITLE> VIN Script Generator </TITLE>" & vbCrLf
    Fnum = FreeFile
    If writable Then
        Open App.Path + "\vinscript\preview.html" For Output As Fnum
        strDir = App.Path + "\vinscript\"  'here you get surely an error if app.path is a drive letter so error is e:\\vinscript\
    Else
     Open "c:\preview.html" For Output As Fnum
     strDir = "c:\"
    End If
     Print #Fnum, txttemp
     Print #Fnum, rtfCode.Text
     txttemp = "</HEAD> " & vbCrLf & "     <BODY bgcolor = #ffffff text = #ffff00>               </BODY> " & vbCrLf & " </HTML>"
     Print #Fnum, txttemp
     Close Fnum        ''=====================closing the output file
     




' Launch topic
Dim hinst As Long
hinst = ShellExecute(Me.hwnd, vbNullString, "preview.html", vbNullString, strDir, SW_SHOWNORMAL)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmAdder
End Sub

Private Sub helper_Click(Index As Integer)
'CommonDialog1.HelpFile = "C:\Documents and Settings\raj1\Desktop\all2html.chm"
'CommonDialog1.ShowHelp
If Index = 0 Then    'help
 MsgBox "                                HELP " & vbCrLf & _
 "**************************************" & vbCrLf & _
 "Step1:- Select any category" & vbCrLf & _
"Step2:- Choose any script from list box." & vbCrLf & _
"              The discription of selected script will shown below list box." & vbCrLf & _
"      - ->>If script do not contain any user defined variables then script code will directly displayed." & vbCrLf & _
"      - ->>But if script contain any user defined variable then you have to enter their values in right pane." & vbCrLf & _
"Step3:- Now press the button to Generate Code. " & vbCrLf & _
"            You must click this button whenever values of variable are modified." & vbCrLf & _
"            The resultant code will be shown in text box." & vbCrLf & _
"Step4:- Now you can Copy this code to paste in any webpage" & vbCrLf & _
"             Or you can save this code for further use." & vbCrLf & _
"             Or you can see the preview of webpage with resultant code." & vbCrLf & _
"             Or you can embedded this code in to any existing webpage." & vbCrLf & _
"*****" & vbCrLf & _
" Click on script adder menu to add more scripts and make a huge script collection."
ElseIf Index = 1 Then    'about
 MsgBox "                                About " & vbCrLf & _
 "**************************************" & vbCrLf & _
 "                               VIN Script Generator " & vbCrLf & _
"                               Programmed By : - VINOD KOTIYA    " & vbCrLf & _
"                              Created On:- 09-06-2003 to 10-06-2003 " & vbCrLf & _
"                              Time :- 23:00 PM to 2:00 AM   " & vbCrLf & _
"                              Total Hours :- 3hr.       " & vbCrLf & _
"                              Time for arranging data files :- 14:30 PM to 16:00 PM " & vbCrLf & _
"                              Proudly releasing version 1.0 " & vbCrLf & _
"///////////////////////////////////////////////////////////////////////////////" & vbCrLf & _
"  FIRST MODIFICATION :- Making user freindly so that" & vbCrLf & _
"  user can upgrade the data files by adding more scripts " & vbCrLf & _
"                           Proudly releasing version 2.0" & vbCrLf & _
"//////////////////////////////////////////////////////////////////////////////"
ElseIf Index = 2 Then    'about me
 MsgBox "                              About Me" & vbCrLf & _
 "**************************************" & vbCrLf & _
"  Programmer: - VINOD KOTIYA    " & vbCrLf & _
"                             s/o Shri Ramesh Kotiya " & vbCrLf & _
"                             B.E. 2nd Year (Information Technology) " & vbCrLf & _
"                             Add:- S-2 Shrimaya Apart Sector - B/363 " & vbCrLf & _
"                                        Sarvdharm Colony, Bhopal (India)" & vbCrLf & _
"                             Fone:- +91-0755-2794428" & vbCrLf & _
"                             Web:- http:\\vinodkotiya.tripod.com " & vbCrLf & _
"                             Email:- vinodkotiya24@rediffmail.com" & vbCrLf & _
"**********" & vbCrLf & _
" Please send your complain's and suggestions." & vbCrLf & _
"//////////////////////////////////////////////////////////////////////////////"

End If
End Sub

Private Sub listScripts_Click()
checkMedia
On Error GoTo vinerror
Dim Fnum As Integer   ' Descriptor for file.
Dim currentline As String       'in loop store currentline of selected file
Dim txtloadrtfCode As String
Dim i As Integer
For i = 0 To 3                   ''Refresh the input label and text box
  lblInput(i).Enabled = True
  txtInput(i).Enabled = True
   txtInput(i).BackColor = vbWhite
  lblInput(i).Caption = "Variable " & i + 1
  txtInput(i).Text = " "
 Next
rtfCode.Text = " "
cmdCode.Enabled = True
'txtInput(0).Text = listScripts.List(listScripts.ListIndex)

Fnum = FreeFile
Open App.Path & "\vinscript\" & listScripts.List(listScripts.ListIndex) & ".vin" For Input As Fnum
       Line Input #Fnum, currentline         'getting 1st line
       
       If Trim(currentline) = "<inputYes>" Then   '!!!!!!!!!!!!!!!!! Script has user variables
                                                                                    '!!!!!!!!!!!!!!!! hence it need to be updated
         For i = 0 To 3
            Line Input #Fnum, currentline      'nil
            Line Input #Fnum, currentline     'input label
            If Trim(currentline) <> "no" Then     'i.e. label is not no
              lblInput(i).Caption = currentline
              Line Input #Fnum, currentline     'input text box
              txtInput(i).Text = currentline
              txtInput(i).BackColor = &HFFFF80
            Else
             Line Input #Fnum, currentline      'nil i.e. label was no so text is also no so disable them
              lblInput(i).Enabled = False
              txtInput(i).Enabled = False
            End If
         Next
           
           Line Input #Fnum, currentline    '<discription>
           txtDetail.Text = currentline
           Line Input #Fnum, currentline   'input discription
           txtDetail.Text = txtDetail.Text & "    " & currentline           'vbcrlf
           Line Input #Fnum, currentline 'nil <end By vinod kotiya >
           
          While Not EOF(Fnum)              'now original script part is storing in to txtloadrtfCode
              Line Input #Fnum, currentline
              txtloadrtfCode = txtloadrtfCode & vbCrLf & currentline
           Wend
        'Input #fnum, txtloadrtfCode
            Close Fnum          '=====================closing the input file when if
           Fnum = FreeFile
           If writable Then
            Open App.Path & "\temple .vin" For Output As Fnum
           Else
             Open "c:\temple .vin" For Output As Fnum
           End If
           Print #Fnum, txtloadrtfCode     'for frequently used original script is saved in temple.vin
            Close Fnum        ''=====================closing the output file
            txtloadrtfCode = " "                  'freeing memory
          '  rtfCode.Text = txtloadrtfCode
    Else        '!!!!!!!!!!!!!!!!!!!       'i.e. the script has no variable setting so get 1st line discription and directly show
                                                    'the code .disable input variables and generate code
                                           '!!!!!!!!!!!!!!!! hence it need not to be updated
           txtDetail.Text = currentline
           Close Fnum           '=====================closing the input file when else
        For i = 0 To 3
         lblInput(i).Enabled = False
         txtInput(i).Enabled = False
        Next
        cmdCode.Enabled = False
               
          rtfCode.LoadFile App.Path & "\vinscript\" & listScripts.List(listScripts.ListIndex) & ".vin", rtfText
   End If       'end of <inputYes>
   Exit Sub
vinerror:
 MsgBox "unable to write on disk.Probably you are running this application directly from CD Drive." & vbCrLf & _
 "To Eliminate error please copy it to your hard drive"
End Sub
Private Sub checkMedia()
On Error GoTo notWrite
 rtfCode.SaveFile App.Path & "\temple.vin"
 writable = True
 Exit Sub
 
notWrite:
 writable = False
End Sub
Private Sub optCategory_Click(Index As Integer)
listScripts.Clear
On Error GoTo vinerror
If Index = 0 Then
 listScripts.AddItem "Magic_Wand"
 listScripts.AddItem "Magic_Wand2"
 listScripts.AddItem "Dancing Stars"
 listScripts.AddItem "Elastic Trail"
 listScripts.AddItem "Image Mouse Stars"
 listScripts.AddItem "Logo Orbit"
 listScripts.AddItem "Mouse Comet"
 listScripts.AddItem "Mouse Star"
 listScripts.AddItem "Sparkle Trail"
 listScripts.AddItem "Text Trail"
 listScripts.AddItem "Trio2"
ElseIf Index = 1 Then
 listScripts.AddItem "Background Image"
 listScripts.AddItem "Fireworks"
 listScripts.AddItem "Snow"
 ElseIf Index = 2 Then
  listScripts.AddItem "Calander"
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
 listScripts.AddItem "Translucent Messages"
 listScripts.AddItem "Text Trail"
 listScripts.AddItem "Circular Text"
ElseIf Index = 5 Then
 listScripts.AddItem "Text Scroll"
 listScripts.AddItem "Typing Message"
 listScripts.AddItem "Clock"
 listScripts.AddItem "Flying Text"
 listScripts.AddItem "Bubble Text"
 
ElseIf Index = 6 Then
listScripts.AddItem "Kiss"
listScripts.AddItem "Elastic Image"
listScripts.AddItem "Bouncing Image"
listScripts.AddItem "image animator"
listScripts.AddItem "Mouse Over Button"
ElseIf Index = 7 Then
listScripts.AddItem "TicTacToe"
listScripts.AddItem "Crosshair"
 listScripts.AddItem "Drive Viewer"
 listScripts.AddItem "Horoscope Sign"
 listScripts.AddItem "Length Converter"
 listScripts.AddItem "Temperature Converter"
 listScripts.AddItem "Web Search"
 listScripts.AddItem "Punchline"
 
End If
Dim Fnum As Integer
Dim currentline As String
Fnum = FreeFile
Open App.Path & "\vinscript\" & Index & ".vin" For Input As Fnum
  While Not EOF(Fnum)              'now original script part is storing in to txtloadrtfCode
              Line Input #Fnum, currentline
              If Trim(currentline) <> " " Then
                  listScripts.AddItem currentline
               End If
   Wend
   Close Fnum
 Exit Sub
vinerror:
  MsgBox "file handling error"
End Sub
