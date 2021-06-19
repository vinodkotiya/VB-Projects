VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   3420
   ClientTop       =   3390
   ClientWidth     =   6885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmStart"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2730
      Left            =   40
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   40
      Width           =   6810
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   540
         Left            =   2280
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   337.12
         ScaleMode       =   0  'User
         ScaleWidth      =   337.12
         TabIndex        =   7
         Top             =   360
         Width           =   540
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   6360
         Top             =   600
      End
      Begin VB.Label lblLoad 
         BackStyle       =   0  'Transparent
         Caption         =   "Initiating BGM ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Shape shpSeek 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF80&
         FillColor       =   &H0000FFFF&
         Height          =   105
         Left            =   2640
         Shape           =   3  'Circle
         Top             =   1515
         Width           =   105
      End
      Begin VB.Line gulabi 
         BorderColor     =   &H00FF00FF&
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   5
         X1              =   2400
         X2              =   2640
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line red 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   2400
         X2              =   6600
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   120
         Picture         =   "frmSplash.frx":0316
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright: Vinod Kotiya"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   1860
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Company : VinSoft ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Add : S-2 Shrimaya Apartment Sector-B/363 Sarvdharm Col.   Bhopal. +91-0755-2794428  E-mail :- vinodkotiya24@rediffmail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Left            =   2280
         TabIndex        =   2
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIN UTILITY KIT "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   660
         Left            =   2280
         TabIndex        =   6
         Top             =   900
         Width           =   4335
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Licensed To : Unknown User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6615
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vinod Kotiya's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2880
         TabIndex        =   5
         Top             =   480
         Width           =   2505
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   7
         DrawMode        =   14  'Copy Pen
         X1              =   2400
         X2              =   6600
         Y1              =   1560
         Y2              =   1560
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Count_ten_sec As Integer   'after ten second unload me
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_SHOWWINDOW = &H40


Private Sub Form_KeyPress(KeyAscii As Integer)
   If Timer1.Interval > 3000 Then
    younload
    End If
End Sub

Private Sub Form_Load()
   ' lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    frmSplash.Show
  
  songfilename = App.Path & "\data\fall.mid"
  closesong
  loaduserName
  
  'init global
  Count_ten_sec = 0
  End Sub

Private Sub Frame1_Click()
If Timer1.Interval > 3000 Then
    younload
End If
End Sub

Private Sub Timer1_Timer()
gulabi.X2 = gulabi.X1 + Round(Count_ten_sec * (4200 / 10))
shpSeek.Left = gulabi.X1 + Round(Count_ten_sec * (4200 / 10))  '6120 is lines length
'gulabi.X2 = shpSeek.Left
lblLoad.Caption = lblLoad.Caption & ".."
Count_ten_sec = Count_ten_sec + 1
If Count_ten_sec = 10 Then    '10 seccompletes

younload
ElseIf Count_ten_sec = 2 Then   'play the song
lblLoad.Caption = "Now Loading.."
playsong
End If

End Sub

Private Sub imgLogo_Click()
If Timer1.Interval > 3000 Then
younload
End If
End Sub
Private Sub younload()
Beep
lblLoad.Caption = "Now Initializing"
  shpSeek.Left = red.X2 - shpSeek.Width
  gulabi.X2 = red.X2
  Load frmmain
    frmmain.Show
  'playsong
    Unload Me

End Sub



Private Sub lblCopyright_Click()
If Timer1.Interval > 3000 Then
younload
End If
End Sub




Private Sub loaduserName()
Dim fnum As Integer
Dim txt As String
Timer1.Interval = 0
On Error GoTo FileError
   
    fnum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\user.vin" For Input As fnum    'dont use #1 for multiple file openings
   Line Input #fnum, txt
    Close #fnum

    If Trim(txt) = "unknown user" Then
   '  MsgBox "icome"
       txt = InputBox("Congratulations !!! You are first time using VIN UTILITY KIT v1.0 " & Chr(13) & _
       "This software is FREE4YOU " & Chr(13) & "Please Enter your name", "User Name", txt)
       writeuserName (txt)
       MsgBox "Now creating some backups for Troubleshooter" & Chr(13) & _
       "Please wait for a little While......" & Chr(13) & Chr(13) & "Press OK"
       BackupData        'firsttime running so backup the data directory
       If Screen.Height < 9005 Then
        MsgBox "It is advisable to you that increase your monitors resolution " & Chr(13) & _
        " to 1024 X 768 for better picture quality"
       End If
    End If
    Dim retValue As Long     'this is not first time so top of lall windows
    retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, 250, 300, _
               466, 190, SWP_SHOWWINDOW)
    lblLicenseTo.Caption = "This software is Licensed To : " & txt
    Timer1.Interval = 300
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "user.vin" _
     & "file is affected by any fool "
'End
End Sub
Private Sub writeuserName(username As String)
Dim fnum As Integer
On Error GoTo FileError
 
    fnum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\user.vin" For Output As fnum    'dont use #1 for multiple file openings
    Print #fnum, username
    Close #fnum
   
   Exit Sub

FileError:
    MsgBox "Unkown error while closing file " & "user.vin" _
     & "file is effected by any fool "

End Sub

Private Sub BackupData()
Dim Fsys As New FileSystemObject
On Error GoTo vinerror
If Fsys.FolderExists("c:\windows\vinbakup") = False Then 'if folder not exist create it
  Fsys.CreateFolder "c:\windows\vinbakup"
End If
Fsys.CopyFolder App.Path & "\data", "c:\windows\vinbakup", True
'Fsys.CopyFile "c:\cdata\*.txt", "c:\windows\vinbakup", True

Exit Sub
vinerror:
 MsgBox "file handling error occured,trouble shoot will probably not work for restore"

End Sub
