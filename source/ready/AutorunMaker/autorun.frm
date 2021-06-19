VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Easy Autorun Maker"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   Icon            =   "autorun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H00FFFFFF&
      Caption         =   "# About Me #"
      Height          =   255
      Index           =   3
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H00FFFFFF&
      Caption         =   "# About #"
      Height          =   255
      Index           =   2
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H00FFFFFF&
      Caption         =   "? HELP ?"
      Height          =   375
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create Autorun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   6255
      Begin VB.OptionButton optIco 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Main File's(*.exe)  Icon"
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
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   0
         Width           =   2895
      End
      Begin VB.OptionButton optIco 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Specify The Icon File"
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open *.ico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Open the *.*.vin file for merging."
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "No icon file is selected for autorun"
         Top             =   480
         Width           =   4695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open *.*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Open the *.*.vin file for merging."
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "No file is selected for autorun"
      Top             =   600
      Width           =   4815
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1920
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3360
      TabIndex        =   0
      Top             =   2280
      Width           =   3360
   End
   Begin VB.Label lblTip 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   6375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Specify the Main File Which will be run by AUTORUN when inserting CD"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   6375
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1320
      X2              =   1320
      Y1              =   2400
      Y2              =   2760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Destination Folder"
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1200
      X2              =   1440
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   1320
      X2              =   3360
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mainFile As String
Dim iconFile As String
Dim destFolder As String
Dim txtAutorun As String
Dim isWhite As Boolean 'false when not white

Private Sub cmdMake_Click(Index As Integer)

If Index = 0 Then
 If Trim(mainFile) = "" Then
   MsgBox "Please Specify the Main File"
   Exit Sub
 End If
  If Trim(iconFile) = "" And optIco(0).Value Then
   MsgBox "Please Specify the Icon File"
   Exit Sub
 End If
 
   destFolder = Dir1.Path 'folder where splitted files be stored
    If Len(destFolder) < 4 Then    'If "F:\" = 3
      destFolder = Left(destFolder, 2)  'if "F:\" it return "F:"
    End If
  Dim pos As Long
    pos = InStrRev(mainFile, "\", -1, vbBinaryCompare)
  txtAutorun = "[autorun]" & vbCrLf
  txtAutorun = txtAutorun & "open=" & "start " & Right(mainFile, Len(mainFile) - pos) & vbCrLf
  If optIco(0).Value Then
  pos = InStrRev(iconFile, "\", -1, vbBinaryCompare)
   txtAutorun = txtAutorun & "icon=" & Right(iconFile, Len(iconFile) - pos) & vbCrLf
  ElseIf optIco(1).Value Then
   txtAutorun = txtAutorun & "icon=" & Right(mainFile, Len(mainFile) - pos) & ",0" & vbCrLf
  End If
GenerateData
MsgBox "         Congratulations..." & vbCrLf & _
 "The AUTORUN Is successfully Created and saved in " & destFolder

ElseIf Index = 1 Then
 MsgBox "                       HELP" & vbCrLf & _
                "               *************" & vbCrLf & _
"         Create AUTORUN in 4 easy Steps" & vbCrLf & vbCrLf & _
" STEP1:  Specify your main file which will be run by AUTORUN when CD inserted." & vbCrLf & vbCrLf & _
" STEP2:  (a)Specify your icon file which will be shown by AUTORUN as CD's icon." & vbCrLf & _
"         (b) Or if main file is an exe then you can use its icon as CD's icon." & vbCrLf & vbCrLf & _
" STEP3:  Specify your Destination Directory where Autorun and supporting files to be saved." & vbCrLf & vbCrLf & _
" STEP4:  Click on Create Autorun Button." & vbCrLf & vbCrLf & _
"          Congratulations"
ElseIf Index = 2 Then
MsgBox "                     VINSOFT" & vbCrLf & _
                "               *************" & vbCrLf & _
 "         VIN EASY AUTORUN MAKER" & vbCrLf & _
                "               *************" & vbCrLf & _
                "Programmer : VINOD KOTIYA " & vbCrLf & _
" date created: 05/08/2003 " & vbCrLf & _
 " time: 9.30 AM to 10:30 AM " & vbCrLf & _
" web : http:\\vinodkotiya.tripod.com " & vbCrLf & _
"email : vinodkotiya24@rediffmail.com "
ElseIf Index = 3 Then
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
Private Sub GenerateData()
On Error GoTo vinerror
Dim v As Integer
v = FreeFile
   Open destFolder & "\" & "autorun.inf" For Output As v
    Print #v, txtAutorun
   Close #v
Dim Fsys As New FileSystemObject

Fsys.CopyFile mainFile, destFolder & "\", True        'copy main file
If optIco(0).Value Then
 Fsys.CopyFile iconFile, destFolder & "\", True        'copy icon file
End If
Exit Sub
vinerror:
 MsgBox "file handling error occured."
End Sub
Private Sub cmdMake_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If isWhite Then
cmdMake(Index).BackColor = &HC0C0C0
isWhite = False
End If
If Index = 0 Then
  lblTip.Caption = "Click to create autorun"
ElseIf Index = 1 Then
  lblTip.Caption = "Click to get HELP ?"
ElseIf Index = 2 Then
  lblTip.Caption = "Click to see About this utility."
ElseIf Index = 3 Then
  lblTip.Caption = "Click to know about programmer."
End If
  
End Sub

Private Sub Command1_Click(Index As Integer)
CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   CommonDialog1.FileName = ""
   If Index = 0 Then
     CommonDialog1.Filter = "All Files|*.*"
   Else
       CommonDialog1.Filter = "Icon Files|*.ico"
   End If
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then
      Text2(Index).Text = "No file is selected"
      Exit Sub
   End If
   If Index = 0 Then
   mainFile = CommonDialog1.FileName '*.*.vin
    Text2(0).Text = mainFile
   Else
    iconFile = CommonDialog1.FileName '*.*.vin
    Text2(1).Text = iconFile
   End If
   
    

End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If isWhite Then
Command1(Index).BackColor = &HC0C0C0
isWhite = False
End If
If Index = 0 Then
 lblTip.Caption = "Click to Open the main file which will be run by AUTORUN when inserting CD."
Else
  lblTip.Caption = "Click to Open the icon file which will be shown by AUTORUN as your CD's icon."
End If
End Sub

Private Sub Dir1_Change()
ChDir Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo vinerror
ChDrive Dir1.Path
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    Exit Sub
vinerror:
  MsgBox "There is no disk in drive"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isWhite = False Then
cmdMake(0).BackColor = vbWhite
cmdMake(1).BackColor = vbWhite
cmdMake(2).BackColor = vbWhite
cmdMake(3).BackColor = vbWhite
Command1(0).BackColor = vbWhite
isWhite = True
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isWhite = False Then
 Command1(1).BackColor = vbWhite
  isWhite = True
End If
End Sub

Private Sub optIco_Click(Index As Integer)
If optIco(0).Value Then
 Command1(1).Enabled = True
 Text2(1).Enabled = True
Else
 If "EXE" <> UCase(Right(mainFile, 3)) Then
   MsgBox "Your main file " & mainFile & " is not an exe. So its icon can't be used as your cd icon"
   Exit Sub
 End If
 Command1(1).Enabled = False
 Text2(1).Enabled = False
 
End If
End Sub

Private Sub optIco_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
  lblTip.Caption = "Click to Open the icon file which will be shown by AUTORUN as your CD's icon."
Else
  lblTip.Caption = "The icon of your main file (it must be an exe)will be shown by AUTORUN as your CD's icon."
End If
End Sub
