VERSION 5.00
Begin VB.Form frmPrev 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPrev.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
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
      Picture         =   "frmPrev.frx":7BC1
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
      Picture         =   "frmPrev.frx":7D2F
      Top             =   4320
      Width           =   900
   End
   Begin VB.Image imgFinish 
      Height          =   300
      Index           =   0
      Left            =   5760
      Picture         =   "frmPrev.frx":8158
      Top             =   6000
      Width           =   900
   End
   Begin VB.Image imgPrint 
      Height          =   300
      Index           =   0
      Left            =   7080
      Picture         =   "frmPrev.frx":85AB
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
      Picture         =   "frmPrev.frx":8972
      Top             =   6000
      Width           =   750
   End
   Begin VB.Image imgBack 
      Height          =   300
      Index           =   0
      Left            =   4920
      Picture         =   "frmPrev.frx":8D68
      Top             =   6000
      Width           =   750
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   0
      Left            =   6720
      Picture         =   "frmPrev.frx":998C
      Top             =   6000
      Width           =   750
   End
   Begin VB.Image imgDown 
      Height          =   180
      Index           =   1
      Left            =   7920
      Picture         =   "frmPrev.frx":A5B0
      Top             =   5400
      Width           =   195
   End
   Begin VB.Image imgDown 
      Height          =   180
      Index           =   0
      Left            =   7920
      Picture         =   "frmPrev.frx":A714
      Top             =   5400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgUp 
      Height          =   180
      Index           =   0
      Left            =   7920
      Picture         =   "frmPrev.frx":A88A
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
      Picture         =   "frmPrev.frx":AC5A
      Top             =   5160
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   1
      Left            =   6720
      Picture         =   "frmPrev.frx":ADC2
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgBack 
      Height          =   300
      Index           =   1
      Left            =   4920
      Picture         =   "frmPrev.frx":B15E
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgCancel 
      Height          =   300
      Index           =   1
      Left            =   2640
      Picture         =   "frmPrev.frx":B4F9
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgPrint 
      Height          =   300
      Index           =   1
      Left            =   7080
      Picture         =   "frmPrev.frx":B8E6
      Top             =   500
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgFinish 
      Height          =   300
      Index           =   1
      Left            =   5760
      Picture         =   "frmPrev.frx":BC7D
      Top             =   6000
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgBrowse 
      Height          =   300
      Index           =   1
      Left            =   6720
      Picture         =   "frmPrev.frx":C0C3
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
Attribute VB_Name = "frmPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temp As String 'temp dir
Dim progressSpeed As Integer

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

  temp = Environ("tmp")
On Error Resume Next
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
 
 hideCtrl
 
Me.Visible = True
End Sub
Private Sub hideCtrl()
 progressSpeed = 0
  lnTop(0).Visible = False
 lnTop(1).Visible = False
  lnBack.Visible = False
 
   
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
txtSys(0).Visible = False
txtSys(1).Visible = False
lblSys(0).Visible = False
lblSys(1).Visible = False
lblSys(2).Visible = False
txtReg(0).Visible = False
txtReg(1).Visible = False
txtReg(2).Visible = False
txtReg(3).Visible = False


End Sub
Public Sub step1()
'loadSettings
hideCtrl
On Error Resume Next
Dim fsys As New FileSystemObject
If frmStart.optMsg(1).Value Then  'welcome mssg is image
      imgWel(0).Visible = True
      lblWel(0).Visible = False
      If fsys.FileExists(frmStart.txtMsg(1).Text) Then
        imgWel(0).Picture = LoadPicture(frmStart.txtMsg(1).Text)
      Else
       MsgBox "Welcome Image File " & frmStart.txtMsg(1).Text & " Does not exist"
      End If
ElseIf frmStart.optMsg(0).Value Then 'welcome mssg is text

    imgWel(0).Visible = False
    lblWel(0).Visible = True
    lblWel(2).Visible = True
    lblWel(0).Caption = frmStart.txtMsg(0).Text
    lblWel(2).Caption = frmStart.txtMsg(0).Text
     '1 for tranceparent
      If frmStart.chkBack(0).Value Then
       lblWel(0).BackStyle = 0
       lblWel(2).BackStyle = 0
      Else
       lblWel(0).BackStyle = 1
       lblWel(2).BackStyle = 1
      End If
      lblWel(0).BackColor = frmStart.lblCol(0).BackColor
      lblWel(2).BackColor = frmStart.lblCol(0).BackColor
      lblWel(0).ForeColor = frmStart.lblCol(1).BackColor
 End If
If frmStart.optMsg(3).Value Then 'welcome mssg is image
      imgWel(1).Visible = True
      lblWel(1).Visible = False
      If fsys.FileExists(Trim(frmStart.txtMsg(3).Text)) Then
        imgWel(1).Picture = LoadPicture(Trim(frmStart.txtMsg(3).Text))
      Else
       MsgBox "Welcome Image File " & frmStart.txtMsg(3).Text & " Does not exist"
      End If
      
     Else 'welcome mssg is text
      
      imgWel(1).Visible = False
      lblWel(1).Visible = True
      lblWel(1).Caption = frmStart.txtMsg(2).Text
      
      If frmStart.chkBack(0).Value Then
       lblWel(1).BackStyle = 0
      Else
       lblWel(1).BackStyle = 1
      End If
      lblWel(1).BackColor = frmStart.lblCol(2).BackColor
      lblWel(1).ForeColor = frmStart.lblCol(3).BackColor
 End If
     
delay (frmStart.txtTime.Text)
Me.Visible = False
End Sub
Public Sub step2()
hideCtrl
On Error Resume Next
      imgWel(0).Visible = False
      lblWel(2).Visible = False
      lblWel(0).Visible = False
      lblTxt.Visible = True
      imgUp(0).Visible = True
      imgDown(1).Visible = True
      imgNext(0).Visible = True
      imgNext(0).Enabled = True
      imgCancel(0).Visible = True
      imgPrint(0).Visible = True
      lblCap(0).Caption = "Software Information"
      lblCap(1).Caption = "Software Information"
      NextWindow = 1
      Dim fnum As Integer
      fnum = FreeFile    'getting file no for futures referance
    
      Open temp & "\software.vin" For Output As fnum    'dont use #1 for multiple file openings
      If Trim(frmAgree.txtAgree(0).Text) = "" Then Print #fnum, "There is no Software Information For" & vbCrLf & frmStart.txtInfo(0).Text & frmStart.txtInfo(1).Text
        Print #fnum, frmAgree.txtAgree(0).Text
     Close fnum
     fnum = FreeFile
      Open temp & "\lisence.vin" For Output As fnum    'dont use #1 for multiple file openings
      If Trim(frmAgree.txtAgree(0).Text) = "" Then Print #fnum, "There is no Lisence Agreement For" & vbCrLf & frmStart.txtInfo(0).Text & frmStart.txtInfo(1).Text
     Print #fnum, frmAgree.txtAgree(1).Text
     Close fnum
      LoadFile ("software.vin")
'frmStart.Picture = LoadPicture(App.Path & "\pic.jpg")

End Sub
Public Sub step3()
NextWindow = 4
hideCtrl
On Error Resume Next
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
  imgNext(0).Visible = True
  imgNext(0).Enabled = True
 lblDir.Visible = True
 txtDir.Visible = True
 txtDir.Caption = frmAppl.txtTarget.Text
 imgBrowse(0).Visible = True
 lblStatus.Caption = "Specify Directory Where you want to install this application.."
 
 'InstallFiles
 'createShortcuts
'  progressSpeed = 0 'stop progressbar
' lnTop(0).X2 = lnBack.X2
' lnTop(1).X2 = lnBack.X2
' delay (0.5)
 'imgNext(0).Visible = True
 'imgNext(1).Visible = True
 
End Sub

Public Sub step4()
hideCtrl
NextWindow = 5
On Error Resume Next
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
 imgNext(0).Visible = True
 imgNext(0).Enabled = True
 If frmSys.chkSys(2).Value Then
 
  txtReg(0).Text = ""
  txtReg(1).Text = ""
  txtReg(2).Text = ""
  txtReg(3).Text = ""
 End If
 If frmSys.chkSys(0).Value = True Then
 txtSys(0).Text = Environ("username")
 End If
 If frmSys.chkSys(1).Value = True Then
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

End Sub
Public Sub step5()
hideCtrl
On Error Resume Next
Dim fsys As New FileSystemObject
  lnBack.Visible = False
 lnTop(0).Visible = False
 lnTop(1).Visible = False
 lblCap(0).Caption = "Installation Complete"
  lblCap(1).Caption = "Installation Complete"
  chkEnd(0).Visible = True
  chkEnd(1).Visible = True
  chkEnd(2).Visible = True
  lblStatus.Caption = "Click on FINISH to exit Setup.."
  If frmEnd.chkSys(0).Value Then chkEnd(1).Enabled = True
  If frmEnd.optChk(0).Value Then chkEnd(1).Value = Checked
  If frmEnd.chkSys(1).Value Then chkEnd(0).Enabled = True
  If frmEnd.optChk(2).Value Then chkEnd(0).Value = Checked
  If frmEnd.chkSys(2).Value Then chkEnd(2).Enabled = True
  imgNext(0).Visible = False
 imgNext(1).Visible = False
 imgCancel(0).Visible = False
 imgCancel(1).Visible = False
 imgFinish(0).Visible = True
 imgFinish(1).Visible = True
 
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
Dim fnum As Integer
Dim currentline As String
Dim msg As String
msg = ""
On Error GoTo fileerror
   
    fnum = FreeFile    'getting file no for futures referance
    Open temp & "\" & fileName For Input As fnum    'dont use #1 for multiple file openings
    While Not EOF(fnum)
     Line Input #fnum, currentline  '<color>
     msg = msg & currentline & vbCrLf
      Line = Line + 1
      If Line = 22 Then
      Pages = Pages + 1
      colInfo(Pages) = msg 'colInfo.Add msg
        
        msg = ""
        Line = 0
      End If
    Wend
      Close #fnum
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
If NextWindow = 5 Then
  step3
  NextWindow = 4
End If

If NextWindow = 2 Then
  NextWindow = 1
  LoadFile ("software.vin")
  Label.Visible = False
 chkAgree(0).Visible = False
 chkAgree(1).Visible = False
 lblCap(0).Caption = "Software Information"
 lblCap(1).Caption = "Software Information"
 imgBack(0).Visible = False
  imgBack(1).Visible = False
   imgNext(0).Enabled = True
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
  'RemoveBackups
  delay (2)
 frmPrev.Visible = False
 End If
 
End If

End Sub

Private Sub RemoveBackups()
'temp\software.vin   lisence.vin  readme.txt  link.vbs
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
If Index = 0 Then
lblStatus.Caption = "Removing backup files......"
 imgFinish(0).Visible = True
 imgFinish(1).Visible = False
 chkEnd(0).Visible = False
 chkEnd(1).Visible = False
 chkEnd(2).Visible = False

 If frmEnd.optMsg(1).Value Then
  Dim fsys As New FileSystemObject
  imgWel(0).Visible = True
  lblWel(0).Visible = False
      If fsys.FileExists(EndImage) Then
        imgWel(0).Picture = LoadPicture(Trim(frmEnd.txtMsg(1).Text))
      Else
       MsgBox "Finishing Image File " & Trim(frmEnd.txtMsg(1).Text) & " Does not exist"
      End If
  Else  'text message
      imgWel(0).Visible = False
      lblWel(0).Visible = True
      lblWel(2).Visible = True
      lblWel(0).Caption = Trim(frmEnd.txtMsg(0).Text)
      lblWel(2).Caption = Trim(frmEnd.txtMsg(0).Text)
        '1 for tranceparent
      If frmEnd.chkBack.Value Then
       lblWel(0).BackStyle = 0
       lblWel(2).BackStyle = 0
      Else
       lblWel(0).BackStyle = 1
       lblWel(2).BackStyle = 1
      End If
             'backcol
      lblWel(0).BackColor = frmEnd.lblCol(0).BackColor
      lblWel(2).BackColor = frmEnd.lblCol(0).BackColor
           'forecol
      lblWel(0).ForeColor = frmEnd.lblCol(1).BackColor
  
 End If  'end of isendimage
 delay (frmEnd.txtTime.Text)
 frmPrev.Visible = False
End If

End Sub

Private Sub imgNext_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgNext(1).Visible = True
 imgNext(0).Visible = False
End If
End Sub

Private Sub createShortcuts()
End Sub

Private Sub imgNext_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 imgNext(0).Visible = True
 imgNext(1).Visible = False
End If
If NextWindow = 5 Then
  Dim isCodeMatch As Boolean
   If frmSys.chkSys(2).Value Then isCodeMatch = checkRegCode
   If isCodeMatch = True Then MsgBox "Registration Code is Matched Successfully."
End If
If NextWindow = 4 Then
 NextWindow = 5
 progressSpeed = 68
 lblCap(0).Caption = "Copying Files"
  lblCap(1).Caption = "Copying Files"
  lblDir.Visible = False
 txtDir.Visible = False
 imgBrowse(0).Visible = False
 imgBrowse(1).Visible = False
 imgNext(0).Visible = False
 imgCancel(0).Visible = True
 imgNext(1).Visible = False
 lnBack.Visible = True
 lnTop(0).Visible = True
 lnTop(1).Visible = True
  
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
If progressSpeed > 0 Then      ' progressSpeed = 0 to skip this area
 If lnTop(0).X2 < lnBack.X2 Then
  lnTop(0).X2 = lnTop(0).X2 + progressSpeed
  lnTop(1).X2 = lnTop(1).X2 + progressSpeed
  progressSpeed = Rnd(23) * 150
    If progressSpeed < 50 Then progressSpeed = 64
 Else
 lnTop(0).X2 = lnTop(0).X1
 lnTop(1).X2 = lnTop(1).X1
 lblStatus.Caption = "Copying File.... " & Chr(Int(Rnd(95) * 95)) & Chr(Int(Rnd(95) * 105)) & Chr(Int(Rnd(25) * 125)) & Chr(Int(Rnd(25) * 85)) & Chr(Int(Rnd(35) * 100)) & Chr(Int(Rnd(75) * 108)) & Chr(Int(Rnd(85) * 122)) & Chr(Int(Rnd(120) * 98)) & "." & Chr(Int(Rnd(65) * 100)) & Chr(Int(Rnd(75) * 107)) & Chr(Int(Rnd(85) * 75))
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
   If frmSys.chkSys(2).Value Then isCodeMatch = checkRegCode
   If isCodeMatch = True Then MsgBox "Registration Code is Matched Successfully."
 End If
End If
End Sub
Private Function checkRegCode() As Boolean
Dim RegCode As String
RegCode = Trim(frmSys.txtReg(0).Text & frmSys.txtReg(1).Text & frmSys.txtReg(2).Text & frmSys.txtReg(3).Text)
If Trim(txtReg(0).Text & txtReg(1).Text & txtReg(2).Text & txtReg(3).Text) = RegCode Then
  checkRegCode = True
Else
 MsgBox "Registration Code Is Wrong. You Can't Proceed......"
 checkRegCode = False
End If
End Function
