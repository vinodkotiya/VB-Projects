VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   0  'None
   Caption         =   "VIN Setup Wizard"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":1CCA
   ScaleHeight     =   7185
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEnd 
      BackColor       =   &H0017D43D&
      Caption         =   "Reboot System"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   23
      Top             =   3480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkEnd 
      BackColor       =   &H0017D43D&
      Caption         =   "View ReadMe"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkEnd 
      BackColor       =   &H0017D43D&
      Caption         =   "Launch Application"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   6720
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   3
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   18
      Text            =   "Next"
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   2
      Left            =   5160
      MaxLength       =   4
      TabIndex        =   17
      Text            =   "Clik"
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   16
      Text            =   "Need"
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtReg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   15
      Text            =   "No"
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSys 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Text            =   "VINSOFT"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtSys 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   11
      Text            =   "VINOD KOTIYA"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox chkAgree 
      BackColor       =   &H0000C000&
      Caption         =   "No"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   5520
      Width           =   615
   End
   Begin VB.CheckBox chkAgree 
      BackColor       =   &H0000C000&
      Caption         =   "Yes"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   5520
      Width           =   615
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   2640
      X2              =   6120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   6
      Index           =   0
      Visible         =   0   'False
      X1              =   2640
      X2              =   6120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label txtDir 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\PROGRAM FILES\VINSOFT"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   1920
      Y1              =   6530
      Y2              =   6530
   End
   Begin VB.Image imgSpark 
      Height          =   75
      Left            =   1800
      Picture         =   "Form1.frx":988B
      Top             =   6510
      Width           =   300
   End
   Begin VB.Label lblDir 
      BackStyle       =   0  'Transparent
      Caption         =   "Installation Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblSys 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Code :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblSys 
      BackStyle       =   0  'Transparent
      Caption         =   "Company :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSys 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Extracting Files"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   6840
      Width           =   4815
   End
   Begin VB.Image imgWel 
      Height          =   375
      Index           =   1
      Left            =   480
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image imgWel 
      Height          =   2895
      Index           =   0
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label lblWel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To VIN Split And Merge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblWel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To VIN Split And Merge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Image imgBrowse 
      Height          =   300
      Index           =   0
      Left            =   6720
      Picture         =   "Form1.frx":99F9
      Top             =   4320
      Width           =   900
   End
   Begin VB.Image imgFinish 
      Height          =   300
      Index           =   0
      Left            =   5760
      Picture         =   "Form1.frx":9E22
      Top             =   6000
      Width           =   900
   End
   Begin VB.Image imgPrint 
      Height          =   300
      Index           =   0
      Left            =   7080
      Picture         =   "Form1.frx":A275
      Top             =   500
      Width           =   750
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "I Am Agree With Terms And Conditions"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Image imgCancel 
      Height          =   300
      Index           =   0
      Left            =   2640
      Picture         =   "Form1.frx":A63C
      Top             =   6000
      Width           =   750
   End
   Begin VB.Image imgBack 
      Height          =   300
      Index           =   0
      Left            =   4920
      Picture         =   "Form1.frx":AA32
      Top             =   6000
      Width           =   750
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   0
      Left            =   6720
      Picture         =   "Form1.frx":B656
      Top             =   6000
      Width           =   750
   End
   Begin VB.Image imgDown 
      Height          =   180
      Index           =   1
      Left            =   7920
      Picture         =   "Form1.frx":C27A
      Top             =   5400
      Width           =   195
   End
   Begin VB.Image imgDown 
      Height          =   180
      Index           =   0
      Left            =   7920
      Picture         =   "Form1.frx":C3DE
      Top             =   5400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgUp 
      Height          =   180
      Index           =   0
      Left            =   7920
      Picture         =   "Form1.frx":C554
      Top             =   5160
      Width           =   195
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO VIN SETUP WIZARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Image imgUp 
      Height          =   180
      Index           =   1
      Left            =   7920
      Picture         =   "Form1.frx":C924
      Top             =   5160
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   1
      Left            =   6720
      Picture         =   "Form1.frx":CA8C
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgBack 
      Height          =   300
      Index           =   1
      Left            =   4920
      Picture         =   "Form1.frx":CE28
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgCancel 
      Height          =   300
      Index           =   1
      Left            =   2640
      Picture         =   "Form1.frx":D1C3
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgPrint 
      Height          =   300
      Index           =   1
      Left            =   7080
      Picture         =   "Form1.frx":D5B0
      Top             =   500
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgFinish 
      Height          =   300
      Index           =   1
      Left            =   5760
      Picture         =   "Form1.frx":D947
      Top             =   6000
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgBrowse 
      Height          =   300
      Index           =   1
      Left            =   6720
      Picture         =   "Form1.frx":DD8D
      Top             =   4320
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblTxt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VINOD KOTIYA"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO VIN SETUP WIZARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   1
      Left            =   2600
      TabIndex        =   5
      Top             =   550
      Width           =   5295
   End
   Begin VB.Label lblWel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To VIN Split And Merge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1455
      Index           =   2
      Left            =   2565
      TabIndex        =   9
      Top             =   1965
      Width           =   5415
   End
   Begin VB.Line lnBack 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      Visible         =   0   'False
      X1              =   2640
      X2              =   7680
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Product As String   'applications name and ver
Dim temp As String 'temp dir
Dim InstallDir As String
Dim isUser As Boolean
Dim isCompany As Boolean
Dim isCode As Boolean
Dim RegCode As String
Dim isPromptReadMe As Boolean
Dim isChkReadMe As Boolean
Dim isPromptLaunch As Boolean
Dim ischkLaunch As Boolean
Dim appName As String 'name of the launching app
Dim isbkLaunch As Boolean 'launch app in background if true
Dim BackappName As String 'name of the launching app in background
Dim isReboot As Boolean
Dim isEndImage As Boolean 'true if image message
Dim endImage As String 'filename
Dim endMsg As String
Dim endMsgTrans As Boolean 'tranceparency
Dim endMsgFore As Long 'forecol
Dim endMsgBack As Long 'BackCol
Dim colDll As New Collection 'contain where dll files to be installed
Dim progressSpeed As Integer  'speed of progressbar
Dim welDelay As Integer ' used for display dealys
Dim endDelay As Integer ' used for display dealys
Dim colInstalledFile As New Collection ''contains path of installed file
''''''''''''''''''''''''
Dim curPage As Integer 'determines the no of currently displayed page
Dim Pages As Integer 'total no of pages
Dim colInfo(1 To 16) As String ' New Collection 'contain info per page
Dim NextWindow As Integer '1 = soft info , 2 = licence agree , 3 = appl
''set on top
'Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const HWND_TOPMOST = -1
'Private Const SWP_SHOWWINDOW = &H40
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


Private Sub chkAgree_Click(Index As Integer)
If Index = 0 Then
 chkAgree(1).Value = Unchecked
' chkAgree(0).Value = Checked
Else
 chkAgree(0).Value = Unchecked
 'chkAgree(1).Value = Checked
End If
If chkAgree(0).Value Then
imgNext(0).Enabled = True
Else
 imgNext(0).Enabled = False
End If
End Sub










Private Sub Form_Load()
' If Screen.Width > 15000 Then        'for 1024X 768
 '      SetWindowPos Me.hwnd, HWND_TOPMOST, 350, 270, 574, 480, SWP_SHOWWINDOW
 ' Else                                                    '800 X 600
  '     SetWindowPos Me.hwnd, HWND_TOPMOST, 260, 210, 574, 480, SWP_SHOWWINDOW
  'End If



imgWel(0).Height = imgBack(0).Top - imgWel(0).Top - 120
imgWel(1).Height = imgBack(0).Top - imgWel(1).Top
lblWel(0).Height = imgBack(0).Top - lblWel(0).Top
lblWel(2).Height = imgBack(0).Top - lblWel(2).Top
lblTxt.Height = imgDown(0).Top - lblTxt.Top
lblSys(0).Top = lblTxt.Top + 1000
 txtSys(0).Top = lblSys(0).Top - 30
 lblSys(1).Top = lblSys(0).Top + lblSys(0).Height + 500
 txtSys(1).Top = lblSys(1).Top - 30
 lblSys(2).Top = lblSys(1).Top + lblSys(1).Height + 500
 txtReg(0).Top = lblSys(2).Top + lblSys(2).Height + 300
 txtReg(1).Top = txtReg(0).Top
 txtReg(2).Top = txtReg(0).Top
 txtReg(3).Top = txtReg(0).Top
 
 
lblTxt.Visible = False
imgWel(0).Visible = False
imgWel(1).Visible = False
lblWel(0).Visible = False
lblWel(2).Visible = False
lblWel(1).Visible = False
imgNext(0).Visible = False
imgBack(0).Visible = False
imgCancel(0).Visible = False
imgPrint(0).Visible = False
imgFinish(0).Visible = False
imgBrowse(0).Visible = False
Label.Visible = False
chkAgree(0).Visible = False
chkAgree(1).Visible = False
lblTxt.Visible = False
 imgDown(0).Visible = False
 imgDown(1).Visible = False
 imgUp(0).Visible = False
 imgUp(1).Visible = False
 
Me.Visible = True
loadSettings
delay (welDelay)
  lblStatus.Caption = "Click on Next to proceed..."
      imgWel(0).Visible = False
      lblWel(2).Visible = False
      lblWel(0).Visible = False
      lblTxt.Visible = True
      imgUp(0).Visible = True
      imgDown(1).Visible = True
      imgNext(0).Visible = True
      imgCancel(0).Visible = True
      imgPrint(0).Visible = True
      lblCap(0).Caption = "Software Information"
      lblCap(1).Caption = "Software Information"
  NextWindow = 1
  LoadFile ("software.vin")
'frmStart.Picture = LoadPicture(App.Path & "\pic.jpg")

End Sub
Private Sub LoadFile(fileName As String)
Dim item As Integer
'MsgBox colInfo.item(6)
'If colInfo.Count < Pages Then
'For item = 1 To Pages - 1
'  colInfo.Remove item
'Next
'End If
'MsgBox colInfo.Count
Dim Line As Integer
Pages = 0
Line = 0
curPage = 1
Dim Fnum As Integer
Dim currentline As String
Dim msg As String
msg = ""
On Error GoTo fileerror
   
    Fnum = FreeFile    'getting file no for futures referance
    Open temp & "\" & fileName For Input As Fnum    'dont use #1 for multiple file openings
    While Not EOF(Fnum)
     Line Input #Fnum, currentline  '<color>
     msg = msg & currentline & vbCrLf
      Line = Line + 1
      If Line = 22 Then
      Pages = Pages + 1
      colInfo(Pages) = msg 'colInfo.Add msg
        
        msg = ""
        Line = 0
      End If
    Wend
      Close #Fnum
      If Line > 0 Then
       Pages = Pages + 1
       colInfo(Pages) = msg 'colInfo.Add msg
        
      End If
        
    curPage = 0

    Call imgDown_MouseUp(0, 0, 0, 0, 0)
    'MsgBox line
    Exit Sub
fileerror:

End Sub

Private Sub delay(t As Double)
Dim i As Double
i = Timer()
While Timer() - i < t
DoEvents
Wend
End Sub
Private Sub loadSettings()
     
     temp = Environ("tmp")
     'temp\software.vin   lisence.vin  readme.txt  link.vbs
   On Error Resume Next
Dim txtInfo As String
Dim Onum As Integer 'for o/p
Dim Fnum As Integer
Dim currLine As String
Dim fsys As New FileSystemObject
Fnum = FreeFile
Open App.path & "\vinsetup1.vin" For Input As Fnum
     ''''''''' remove redundance ''''''''''''
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     Line Input #Fnum, currLine
     '''''''''''''''''''''''''removed'''''''''''''
     
     Line Input #Fnum, currLine  '<WELLCOME>
     Line Input #Fnum, currLine 'appname
     Product = Trim(currLine)
     
     Line Input #Fnum, currLine 'ver
     Product = Product & " " & Trim(currLine)
     Me.Caption = "Installing " & Product & "By VIN Setup Wizard"
     Line Input #Fnum, currLine  'company
     Line Input #Fnum, currLine
     If Trim(currLine) = "True" Then 'welcome mssg is image
      Line Input #Fnum, currLine 'txtmssg
      Line Input #Fnum, currLine '1 for tranceparent
      Line Input #Fnum, currLine 'backcol
      Line Input #Fnum, currLine 'forecol
      Line Input #Fnum, currLine 'name of image
      imgWel(0).Visible = True
      lblWel(0).Visible = False
      If fsys.FileExists(App.path & "\images\" & Trim(currLine)) Then
        imgWel(0).Picture = LoadPicture(App.path & "\images\" & Trim(currLine))
      Else
       MsgBox "Welcome Image File " & App.path & "\images\" & Trim(currLine) & " Does not exist"
      End If
     Else 'welcome mssg is text
      Line Input #Fnum, currLine 'txtmssg
      imgWel(0).Visible = False
      lblWel(0).Visible = True
      lblWel(2).Visible = True
      lblWel(0).Caption = currLine
      lblWel(2).Caption = currLine
      Line Input #Fnum, currLine '1 for tranceparent
      If Trim(currLine) = "1" Then
       lblWel(0).BackStyle = 0
       lblWel(2).BackStyle = 0
      Else
       lblWel(0).BackStyle = 1
       lblWel(2).BackStyle = 1
      End If
      Line Input #Fnum, currLine 'backcol
      lblWel(0).BackColor = Trim(currLine)
      lblWel(2).BackColor = Trim(currLine)
      Line Input #Fnum, currLine 'forecol
      lblWel(0).ForeColor = Trim(currLine)
      Line Input #Fnum, currLine 'imagename
    End If
    Line Input #Fnum, currLine 'time to display
    welDelay = Trim(currLine)
     'leftdisplay
    Line Input #Fnum, currLine
     If Trim(currLine) = "True" Then 'welcome mssg is image
      Line Input #Fnum, currLine 'txtmssg
      Line Input #Fnum, currLine '1 for tranceparent
      Line Input #Fnum, currLine 'backcol
      Line Input #Fnum, currLine 'forecol
      Line Input #Fnum, currLine 'name of image
      imgWel(1).Visible = True
      lblWel(1).Visible = False
      If fsys.FileExists(App.path & "\images\" & Trim(currLine)) Then
        imgWel(1).Picture = LoadPicture(App.path & "\images\" & Trim(currLine))
      Else
       MsgBox "Welcome Image File " & App.path & "\images\" & Trim(currLine) & " Does not exist"
      End If
      
     Else 'welcome mssg is text
      Line Input #Fnum, currLine 'txtmssg
      imgWel(1).Visible = False
      lblWel(1).Visible = True
      lblWel(1).Caption = currLine
      Line Input #Fnum, currLine '1 for tranceparent
      If Trim(currLine) = "1" Then
       lblWel(1).BackStyle = 0
      Else
       lblWel(1).BackStyle = 1
      End If
      Line Input #Fnum, currLine 'backcol
      lblWel(1).BackColor = Trim(currLine)
      Line Input #Fnum, currLine 'forecol
      lblWel(1).ForeColor = Trim(currLine)
      Line Input #Fnum, currLine 'imagename
     End If
     Line Input #Fnum, currLine  'end of </welcome>
     
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     
     Line Input #Fnum, currLine ' <AGREEMENT>
     DoEvents
     Line Input #Fnum, currLine   '  <Software>
     Do While (Not "</Software>" = Trim(currLine)) 'Or (Not EOF(FNum))
      Line Input #Fnum, currLine
      If "</Software>" = Trim(currLine) Then Exit Do
      txtInfo = txtInfo & currLine & vbCrLf
     Loop
     Onum = FreeFile
     Open temp & "\software.vin" For Output As Onum
     Print #Onum, txtInfo
     Close Onum
     DoEvents
     txtInfo = ""
     Line Input #Fnum, currLine   ' <Lisence>
     Do While (Not "</Lisence>" = Trim(currLine)) 'Or (Not EOF(FNum))
      Line Input #Fnum, currLine
      If "</Lisence>" = Trim(currLine) Then Exit Do
      txtInfo = txtInfo & currLine & vbCrLf
     Loop
     Onum = FreeFile
     Open temp & "\lisence.vin" For Output As Onum
     Print #Onum, txtInfo
     Close Onum
     DoEvents
     txtInfo = ""
     Line Input #Fnum, currLine   ' <Read ME>
     Do While (Not "</ReadMe>" = Trim(currLine)) 'Or (Not EOF(FNum))
      Line Input #Fnum, currLine
      If "</ReadMe>" = Trim(currLine) Then Exit Do
      txtInfo = txtInfo & currLine & vbCrLf
     Loop
     Onum = FreeFile
     Open temp & "\readme.txt" For Output As Onum
     Print #Onum, txtInfo
     Close Onum
     Line Input #Fnum, currLine ' </EOF AGREEMENT>
     
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     Line Input #Fnum, currLine ' <APPLICATION>
     Line Input #Fnum, currLine 'installation dir
     txtDir.Caption = Trim(currLine)
      Line Input #Fnum, currLine '</SystemFiles>
      Line Input #Fnum, currLine 'blankline
     Do While (Not "</SystemFiles>" = Trim(currLine)) 'Or (Not EOF(FNum))
      Line Input #Fnum, currLine
      If "</SystemFiles>" = Trim(currLine) Then Exit Do
      colDll.Add Trim(currLine)
     Loop
     txtInfo = ""
     Line Input #Fnum, currLine '<Shortcut vbscript>
     Do While (Not "</Shortcutvbscript>" = Trim(currLine)) 'Or (Not EOF(FNum))
      Line Input #Fnum, currLine
      If "</Shortcutvbscript>" = Trim(currLine) Then Exit Do
      txtInfo = txtInfo & currLine & vbCrLf
     Loop
     Onum = FreeFile
     Open temp & "\link.vbs" For Output As Onum
     Print #Onum, txtInfo
     Close Onum
     
     Line Input #Fnum, currLine ' eof </APPLICATION>
     
     ''''''''''''''''''''''''''''''''''''''''''''''''
     
     
     Line Input #Fnum, currLine ' <SYSTEM>
     Line Input #Fnum, currLine ' True for username
     If currLine = "True" Then
      isUser = True
     Else
      isUser = False
     End If
     Line Input #Fnum, currLine ' True for company
     If Trim(currLine) = "True" Then
      isCompany = True
     Else
      isCompany = False
     End If
     Line Input #Fnum, currLine ' True for regcode
     If currLine = "True" Then
      isCode = True
     Else
      isCode = False
     End If
     Line Input #Fnum, currLine ' regcode
     RegCode = Trim(currLine)
     Line Input #Fnum, currLine 'eof </SYSTEM>
     
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     
     Line Input #Fnum, currLine ' <END>
     Line Input #Fnum, currLine 'prompt readme  for 1
     If Trim(currLine) = "1" Then
      isPromptReadMe = True
     Else
      isPromptReadMe = False
     End If
     Line Input #Fnum, currLine 'prompt readme  checked if true
     If Trim(currLine) = "False" Then
      isChkReadMe = False
     Else
      isChkReadMe = True
     End If
     Line Input #Fnum, currLine 'prompt launch  for 1
     If Trim(currLine) = "1" Then
      isPromptLaunch = True
     Else
      isPromptLaunch = False
     End If
     Line Input #Fnum, currLine 'prompt launch  checked if true
     If Trim(currLine) = "False" Then
      ischkLaunch = False
     Else
      ischkLaunch = True
     End If
     Line Input #Fnum, currLine ' launch appname
     appName = currLine
     Line Input #Fnum, currLine
     If Trim(currLine) = 1 Then  'reboot system
      isReboot = True
     Else
      isReboot = False
     End If
      Line Input #Fnum, currLine 'launch in background  for 1
     If Trim(currLine) = "1" Then
       isbkLaunch = True
     Else
       isbkLaunch = False
     End If
      Line Input #Fnum, currLine 'launch background appname
      BackappName = currLine
     DoEvents
     'finishing message
      Line Input #Fnum, currLine
     If Trim(currLine) = "True" Then 'finish mssg is image
      Line Input #Fnum, currLine 'txtmssg
      Line Input #Fnum, currLine '1 for tranceparent
      Line Input #Fnum, currLine 'backcol
      Line Input #Fnum, currLine 'forecol
      Line Input #Fnum, currLine 'name of image
      isEndImage = True
      If fsys.FileExists(App.path & "\images\" & Trim(currLine)) Then
        endImage = App.path & "\images\" & Trim(currLine)
      Else
       MsgBox "Welcome Image File " & App.path & "\images\" & Trim(currLine) & " Does not exist"
      End If
      
     Else 'welcome mssg is text
      isEndImage = False
      Line Input #Fnum, currLine 'txtmssg
      endMsg = currLine
      Line Input #Fnum, currLine '1 for tranceparent
      If Trim(currLine) = "1" Then
       endMsgTrans = True
      Else
       endMsgTrans = False
      End If
      Line Input #Fnum, currLine 'backcol
      endMsgBack = Trim(currLine)
      Line Input #Fnum, currLine 'forecol
      endMsgFore = Trim(currLine)
      Line Input #Fnum, currLine 'imagename
     End If
     Line Input #Fnum, currLine 'time to display
    endDelay = Trim(currLine)
     Line Input #Fnum, currLine 'eof</END>
     Close Fnum
End Sub




Private Sub imgBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgBack(1).Visible = True
 imgBack(0).Visible = False
End If


End Sub

Private Sub imgBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgBack(0).Visible = True
 imgBack(1).Visible = False
End If
If NextWindow = 2 Then
  NextWindow = 1
  LoadFile ("software.vin")
  Label.Visible = False
 chkAgree(0).Visible = False
 chkAgree(0).Value = Unchecked
 chkAgree(1).Visible = False
 lblCap(0).Caption = "Software Information"
 lblCap(1).Caption = "Software Information"
 imgBack(0).Visible = False
  imgBack(1).Visible = False
   imgNext(0).Enabled = True
End If

If NextWindow = 3 Then
lblSys(0).Visible = False
 lblSys(1).Visible = False
 lblSys(2).Visible = False
 txtSys(0).Visible = False
 txtSys(1).Visible = False
 txtReg(0).Visible = False
 txtReg(1).Visible = False
 txtReg(2).Visible = False
 txtReg(3).Visible = False
lblTxt.Visible = True
imgUp(0).Visible = True
imgDown(1).Visible = True
NextWindow = 2
 imgBack(0).Visible = True
 LoadFile ("lisence.vin")
 Label.Visible = True
 chkAgree(0).Visible = True
 chkAgree(1).Visible = True
 chkAgree_Click (1)
 lblCap(0).Caption = "Lisence And Agreement"
 lblCap(1).Caption = "Lisence And Agreement"
End If
End Sub

Private Sub imgBrowse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgBrowse(1).Visible = True
 imgBrowse(0).Visible = False
End If
End Sub

Private Sub imgBrowse_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgBrowse(0).Visible = True
 imgBrowse(1).Visible = False
End If
Load frmBrowse
frmBrowse.Show
End Sub

Private Sub imgCancel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgCancel(1).Visible = True
 imgCancel(0).Visible = False
End If

End Sub

Private Sub imgCancel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 Dim temp As String
 imgCancel(0).Visible = True
 imgCancel(1).Visible = False
 
 temp = MsgBox("Are you sure that you want to stop VIN Setup Wizard.", vbYesNo)
 If temp = vbYes Then
  MsgBox "You can install this application any time." & vbCrLf & _
  "VIN Setup Wizard Will Now Removing Backup Files." & vbCrLf & _
  "Please Wait...."
  RemoveBackups
  delay (2)
  End
 End If
End If

End Sub

Private Sub RemoveBackups()
'temp\software.vin   lisence.vin  readme.txt  link.vbs
 On Error Resume Next
Dim fsys As New FileSystemObject
fsys.DeleteFile temp & "\software.vin", True
fsys.DeleteFile temp & "\lisence.vin", True
'Fsys.DeleteFile temp & "\readme.txt", True
fsys.DeleteFile temp & "\link.vbs", True
End Sub
Private Sub imgDown_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then
 imgDown(0).Visible = True
 imgDown(1).Visible = False
End If

End Sub

Private Sub imgDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then
 imgDown(1).Visible = True
 imgDown(0).Visible = False
End If
If (curPage < Pages) Then
   curPage = curPage + 1
   lblTxt.Caption = colInfo(curPage) 'colInfo.item(curPage)
  
End If
If curPage = Pages Then
imgDown(1).Visible = False
imgDown(0).Visible = True
End If
If curPage > 1 Then
imgUp(0).Visible = False
imgUp(1).Visible = True
End If
End Sub

Private Sub imgFinish_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgFinish(1).Visible = True
 imgFinish(0).Visible = False
 
End If

End Sub

Private Sub imgFinish_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Index = 0 Then
 imgFinish(0).Visible = True
 imgFinish(1).Visible = False
 chkEnd(0).Visible = False
 chkEnd(1).Visible = False
 chkEnd(2).Visible = False
  Dim hinst As Long
  ischkLaunch = True
  lblStatus.Caption = "Removing Backups and closing setup..."
 
 If isbkLaunch = True Then hinst = ShellExecute(Me.hwnd, vbNullString, BackappName, vbNullString, txtDir.Caption & "\", SW_SHOWNORMAL)

 RemoveBackups
 If isEndImage Then
  Dim fsys As New FileSystemObject
  imgWel(0).Visible = True
  lblWel(0).Visible = False
      If fsys.FileExists(endImage) Then
        imgWel(0).Picture = LoadPicture(endImage)
      Else
       MsgBox "Finishing Image File " & endImage & " Does not exist"
      End If
  Else  'text message
      imgWel(0).Visible = False
      lblWel(0).Visible = True
      lblWel(2).Visible = True
      lblWel(0).Caption = endMsg
      lblWel(2).Caption = endMsg
        '1 for tranceparent
      If endMsgTrans = "1" Then
       lblWel(0).BackStyle = 0
       lblWel(2).BackStyle = 0
      Else
       lblWel(0).BackStyle = 1
       lblWel(2).BackStyle = 1
      End If
             'backcol
      lblWel(0).BackColor = endMsgBack
      lblWel(2).BackColor = endMsgBack
           'forecol
      lblWel(0).ForeColor = endMsgFore
  
 End If  'end of isendimage
 delay (endDelay)
 If chkEnd(0).Value Then hinst = ShellExecute(Me.hwnd, vbNullString, appName, vbNullString, txtDir.Caption & "\", SW_SHOWNORMAL)
 If chkEnd(1).Value Then hinst = ShellExecute(Me.hwnd, vbNullString, "readme.txt", vbNullString, temp, SW_SHOWNORMAL)
 End
End If

End Sub

Private Sub imgNext_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgNext(1).Visible = True
 imgNext(0).Visible = False
End If
End Sub
Private Sub InstallFiles()
 On Error Resume Next
 progressSpeed = 80
 delay (4)
Dim fsys As New FileSystemObject
Dim ThisFolder As Folder
Dim AllFolders As Folders
Dim AllFiles As Files
Dim i As File
Dim j As Folder
Set ThisFolder = fsys.GetFolder(App.path & "\data")
Set AllFolders = ThisFolder.SubFolders
Set AllFiles = ThisFolder.Files
'If fsys.FolderExists(txtDir.Caption) = False Then
 ' fsys.CreateFolder (txtDir.Caption)
'End If
CreatePath (txtDir.Caption & "\")
'''''''' copy files inside data
For Each i In AllFiles
 fsys.CopyFile i, txtDir.Caption & "\", True
 colInstalledFile.Add txtDir.Caption & "\" & i.Name
 lblStatus.Caption = "Copying File " & i.Name & " ..."
Next
'''''' copy all files and folders inside first subfolders
For Each j In AllFolders
 If fsys.FolderExists(txtDir.Caption & "\" & j.Name) = False Then
  fsys.CreateFolder (txtDir.Caption & "\" & j.Name)
 End If
 Set AllFiles = j.Files
  For Each i In AllFiles
   'delay (0.5)
    progressSpeed = Rnd(23) * 150
    If progressSpeed < 50 Then progressSpeed = 64
   lblStatus.Caption = "Copying File " & i.Name
   fsys.CopyFile i, txtDir.Caption & "\" & j.Name & "\", True
   colInstalledFile.Add txtDir.Caption & "\" & j.Name & "\" & i.Name
  Next
  Dim AllSubFolders As Folders
  Set AllSubFolders = j.SubFolders
  Dim k As Folder
  For Each k In AllSubFolders
  ' delay (0.3)
   progressSpeed = Rnd(23) * 150
    If progressSpeed < 50 Then progressSpeed = 70
   lblStatus.Caption = "Copying Folder " & k.path
   fsys.CopyFolder k, txtDir.Caption & "\" & j.Name & "\", True
  Next
Next 'nandu mama 5091662    '2460081
''' COPYING SYSTEM FILES
Dim FileDll As String
Dim v As Integer
For v = 1 To colDll.Count

 FileDll = fsys.GetFileName(colDll.item(v))
If fsys.FolderExists(fsys.GetParentFolderName(colDll.item(v))) = False Then
  fsys.CreateFolder (fsys.GetParentFolderName(colDll.item(v)))
End If

 fsys.CopyFile App.path & "\system\" & FileDll, fsys.GetParentFolderName(colDll.item(v)) & "\"
 lblStatus.Caption = "Copying System File " & colDll.item(v)
 'delay (0.9)
 progressSpeed = Rnd(23) * 150
    If progressSpeed < 50 Then progressSpeed = 80

Next
'''''''' create uninstall.vin
Dim Fnum As Integer
Fnum = FreeFile
Open txtDir.Caption & "\uninstall.vin" For Output As Fnum
  '''''adding redundancy
  Print #Fnum, "Donot alter anything in this file."
  Print #Fnum, "This File contains information about VIN Setup Wizard."
  Print #Fnum, "Author: VINOD KOTIYA."
  Print #Fnum, "Address: S-2 shrimaya apartment sector-b/363"
  Print #Fnum, "        Sarvdharm colony bhopal-42 india ."
  Print #Fnum, "Fone : +91-0755-2794428  ."
  Print #Fnum, "Email : vinodkotiya24@rediffmail.com"
  Print #Fnum, "Web : http:\\vinodkotiya.tripod.com"
  Print #Fnum, "Student: B.E. 3rd year "
  Print #Fnum, "        Information Technology"
  Print #Fnum, "   University Institute of Technology"
  Print #Fnum, "   Rajeev Gandhi Produogiki Vishwavidyalaya, BHOPAL"
  Print #Fnum, "<JAIMATADI>"
  Print #Fnum, Product
  '''''''
For v = 1 To colInstalledFile.Count
 Print #Fnum, colInstalledFile.item(v)
Next
 Print #Fnum, "</JAIMATADI>"
Close #Fnum, Fnum
lblStatus.Caption = "Copying Successfully Complete.. "

End Sub
Private Sub createShortcuts()
 On Error Resume Next
lblStatus.Caption = "Creating Icons.........."
MsgBox "Now Creating icons......" & vbCrLf & _
 "The icons will be created by executing a script." & vbCrLf & _
 "It may be possible in Network System your antivirus will alert you " & vbCrLf & _
 "Please ignore this alert and allow script to execute otherwise icons will not created." & vbCrLf & _
 "The script is not harmfull to your system"
Dim fsys As New FileSystemObject
If fsys.FileExists(temp & "\link.vbs") Then fsys.CopyFile temp & "\link.vbs", txtDir.Caption & "\"
'If Fsys.FileExists(App.Path & "\.vbs") Then Fsys.CopyFile temp & "\link.vbs", txtDir.Caption & "\"

Dim hinst As Long
If fsys.FileExists(txtDir.Caption & "\vinscript.exe") = True Then
hinst = ShellExecute(Me.hwnd, vbNullString, "zz.bat", vbNullString, txtDir.Caption & "\", SW_SHOWNORMAL)
Else
'MsgBox "vinscript not exist"
hinst = ShellExecute(Me.hwnd, vbNullString, "link.vbs", vbNullString, txtDir.Caption & "\", SW_SHOWNORMAL)
End If
lblStatus.Caption = "Icons Created Press Next to Proceed ..."
End Sub

Private Sub imgNext_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Index = 0 Then
 imgNext(0).Visible = True
 imgNext(1).Visible = False
End If
If NextWindow = 5 Then
  NextWindow = 6
  lnBack.Visible = False
 lnTop(0).Visible = False
 lnTop(1).Visible = False
 lblCap(0).Caption = "Installation Complete"
  lblCap(1).Caption = "Installation Complete"
  chkEnd(0).Visible = True
  chkEnd(1).Visible = True
  chkEnd(2).Visible = True
  lblStatus.Caption = "Click on FINISH to exit Setup.."
  If isPromptReadMe = True Then chkEnd(1).Enabled = True
  If isChkReadMe = True Then chkEnd(1).Value = Checked
  If isPromptLaunch = True Then chkEnd(0).Enabled = True
  If ischkLaunch = True Then chkEnd(0).Value = Checked
  If isReboot = True Then chkEnd(2).Enabled = True
  imgNext(0).Visible = False
 imgNext(1).Visible = False
 imgCancel(0).Visible = False
 imgCancel(1).Visible = False
 imgFinish(0).Visible = True
 imgFinish(1).Visible = True
End If
If NextWindow = 4 Then
  NextWindow = 5
  lblCap(0).Caption = "Copying Files"
  lblCap(1).Caption = "Copying Files"
  lblDir.Visible = False
 txtDir.Visible = False
 imgBrowse(0).Visible = False
 imgBrowse(1).Visible = False
 imgNext(0).Visible = False
 imgNext(1).Visible = False
 lnBack.Visible = True
 lnTop(0).Visible = True
 lnTop(1).Visible = True
 InstallFiles
 createShortcuts
  progressSpeed = 0 'stop progressbar
 lnTop(0).X2 = lnBack.X2
 lnTop(1).X2 = lnBack.X2
 delay (0.5)
 imgNext(0).Visible = True
 imgNext(1).Visible = True

 End If

If NextWindow = 3 Then
 Dim isCodeMatch As Boolean
 If isCode = True Then
   isCodeMatch = checkRegCode
  If isCodeMatch = False Then Exit Sub
 End If
 NextWindow = 4
 lblCap(0).Caption = "Installation Directory"
 lblCap(1).Caption = "Installation Directory"
 lblSys(0).Visible = False
 lblSys(1).Visible = False
 lblSys(2).Visible = False
 txtSys(0).Visible = False
 txtSys(1).Visible = False
 txtReg(0).Visible = False
 txtReg(1).Visible = False
 txtReg(2).Visible = False
 txtReg(3).Visible = False
 imgBack(0).Visible = False
 imgBack(1).Visible = False
 lblDir.Visible = True
 txtDir.Visible = True
 imgBrowse(0).Visible = True
 lblStatus.Caption = "Specify Directory Where you want to install this application.."
End If
If NextWindow = 2 Then
 NextWindow = 3
 Label.Visible = False
 chkAgree(0).Visible = False
 chkAgree(1).Visible = False
 lblCap(0).Caption = "Registration"
 lblCap(1).Caption = "Registration"
 lblTxt.Visible = False
 imgDown(0).Visible = False
 imgDown(1).Visible = False
 imgUp(0).Visible = False
 imgUp(1).Visible = False
 imgPrint(0).Visible = False
 imgPrint(1).Visible = False
 
 If isCode = True Then
 
  txtReg(0).Text = ""
  txtReg(1).Text = ""
  txtReg(2).Text = ""
  txtReg(3).Text = ""
 End If
 If isUser = True Then
 txtSys(0).Text = Environ("username")
 End If
 If isCompany = True Then
 txtSys(1).Text = Environ("USERDOMAIN")
 End If
 lblSys(0).Visible = True
 lblSys(1).Visible = True
 lblSys(2).Visible = True
 txtSys(0).Visible = True
 txtSys(1).Visible = True
 txtReg(0).Visible = True
 txtReg(1).Visible = True
 txtReg(2).Visible = True
 txtReg(3).Visible = True
 lblStatus.Caption = "Enter Registration Code if needed."
 'If isCode = True Then imgNext(0).Enabled = False
 End If

If NextWindow = 1 Then
 NextWindow = 2
 imgBack(0).Visible = True
 LoadFile ("lisence.vin")
 Label.Visible = True
 chkAgree(0).Visible = True
 chkAgree(1).Visible = True
 imgNext(0).Enabled = False
 lblCap(0).Caption = "Lisence And Agreement"
 lblCap(1).Caption = "Lisence And Agreement"
 lblStatus.Caption = "Click on Yes to proceed next...."
End If

End Sub

Private Sub imgPrint_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgPrint(1).Visible = True
 imgPrint(0).Visible = False
End If
End Sub

Private Sub imgPrint_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgPrint(0).Visible = True
 imgPrint(1).Visible = False
 Clipboard.SetText lblTxt.Caption
 MsgBox "Unable to access the Printer.Sorry For inconvenience." & vbCrLf & _
 "The data is copied to clipboard you can paste it on any " & vbCrLf & _
 "text editor and then print it."
End If
End Sub

Private Sub imgUp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then
 imgUp(0).Visible = True
 imgUp(1).Visible = False
End If
End Sub

Private Sub imgUp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then
 imgUp(1).Visible = True
 imgUp(0).Visible = False
End If
If (curPage > 1) Then
  curPage = curPage - 1
   lblTxt.Caption = colInfo(curPage) 'colInfo.item(curPage)
   
End If
If curPage = 1 Then
imgUp(1).Visible = False
imgUp(0).Visible = True
End If
If curPage < Pages Then
imgDown(0).Visible = False
imgDown(1).Visible = True
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If progressSpeed > 0 Then      ' progressSpeed = 0 to skip this area
 If lnTop(0).X2 < lnBack.X2 Then
  lnTop(0).X2 = lnTop(0).X2 + progressSpeed
  lnTop(1).X2 = lnTop(1).X2 + progressSpeed
 Else
 lnTop(0).X2 = lnTop(0).X1
 lnTop(1).X2 = lnTop(1).X1
 End If
End If
If (imgSpark.Left + imgSpark.Width) < (frmStart.ScaleWidth - 300) Then
 imgSpark.Left = imgSpark.Left + 20
 Line1.X2 = Line1.X2 + 20
Else
 imgSpark.Left = 0
 Line1.X2 = 0
End If
DoEvents
End Sub

Private Sub txtReg_Change(Index As Integer)
If Len(txtReg(Index).Text) = 4 Then
 If Index < 3 Then
  txtReg(Index + 1).SetFocus
 ElseIf Index = 3 Then
  Dim isCodeMatch As Boolean
   If isCode = True Then isCodeMatch = checkRegCode
 End If
End If
End Sub
Private Sub CreatePath(path As String)
''''''''''''''''''''''''''''''''''''''''
'''' This Procedure is written by vinod kotiya
''''' 24-08-2003 - 9:30 pm
''''''''''''''''''''''''''''''''''''''''''''
 On Error Resume Next
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

Private Function checkRegCode() As Boolean
If Trim(txtReg(0).Text & txtReg(1).Text & txtReg(2).Text & txtReg(3).Text) = RegCode Then
  checkRegCode = True
Else
 MsgBox "Registration Code Is Wrong. You Can't Proceed......"
 checkRegCode = False
End If
End Function
