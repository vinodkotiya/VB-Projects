VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCrypto 
   BackColor       =   &H00C0C0C0&
   Caption         =   "VIN Cryptograph v2.0"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   Icon            =   "frmCrypto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkKey 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   15
      ToolTipText     =   "Enter your Key Here before Encryption/Decryption. When Decrypting the key Must Be Similar to the Key at Time of Encryption."
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   255
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton cmdAbort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Abort"
      Enabled         =   0   'False
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Decrypt"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Open the *.*.vin file for merging."
      Top             =   5280
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSize 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtInfile 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4320
      Width           =   4695
   End
   Begin VB.TextBox txtoutfile 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   525
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   4695
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Encrypt"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Open the *.*.vin file for merging."
      Top             =   5280
      Width           =   1815
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
      Height          =   3030
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Click to Select input file for Encryption/Decryption"
      Top             =   1200
      Width           =   2355
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
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click to change drive"
      Top             =   1200
      Width           =   3360
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
      Height          =   2730
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click to change folders."
      Top             =   1560
      Width           =   3360
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "0  KB"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Show no of bytes processed."
      Top             =   6240
      Width           =   960
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1440
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      ToolTipText     =   "Show processing time"
      Top             =   5760
      Width           =   600
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00.00 Sec"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   10
      Top             =   6000
      Width           =   960
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   120
      X2              =   5880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblScan 
      BackStyle       =   0  'Transparent
      Caption         =   "No Files Selected .................."
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   6480
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "VIN Cryptograph v2.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "By: VINOD KOTIYA"
      Top             =   0
      Width           =   6015
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   6
      Index           =   0
      X1              =   120
      X2              =   5880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   8
      Index           =   0
      X1              =   120
      X2              =   5880
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "frmCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim outputfile As String
Dim inputfile As String




Private Sub Check1_Click()

End Sub

Private Sub chkKey_Click()
'chkKey.Value = Not chkKey.Value
If chkKey.Value Then
 txtKey.PasswordChar = ""
Else
 txtKey.PasswordChar = "*"
End If

End Sub

Private Sub cmdAbort_Click()
Dim reply As Integer
reply = MsgBox("This will abort the process.Are you sure ?", vbYesNo, "Warning")
If reply = vbYes Then         'yes
 isAbort = True
 lblScan.Caption = "Process Aborted......"
 lnTop(1).X2 = lnTop(1).X1
 lnTop(0).X2 = lnTop(0).X1
End If

End Sub

Private Sub cmdAbort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblScan.Caption = "Click to Abort the Process."
cmdAbort.BackColor = &HFFC0C0
End Sub

Private Sub cmdAbout_Click(Index As Integer)
If Index = 0 Then
 MsgBox "                       HELP" & vbCrLf & _
                "               *************" & vbCrLf & _
 vbCrLf & _
" STEP1: Enter Your Key for Encryption/Decryption.At Decryption This key must be similar to the key as time of encryption." & vbCrLf & vbCrLf & _
" STEP2:  Select the file you want to encrypt/decrypt from mini browser." & vbCrLf & vbCrLf & _
" STEP3:  (a)To Encrypt: Click on button 'Encrypt' then you will be prompted for output filename. " & vbCrLf & _
"         (b)To Decrypt: Click on button 'Decrypt' then you will be prompted for output filename." & vbCrLf & vbCrLf & _
" Abort:  Use this if you want to abort the process." & vbCrLf & vbCrLf & _
" Logout: You can encrypt/decrypt files only ones ." & vbCrLf & _
"          To encrypt/decrypt any other file restart application again."
ElseIf Index = 1 Then
MsgBox "         VIN Cryptograph v2.0" & vbCrLf & _
                "               *************" & vbCrLf & _
                " date created: 11/10/2003 " & vbCrLf & _
 " time: 5 hrs. " & vbCrLf & _
"  Programmer: - VINOD KOTIYA    " & vbCrLf & _
"                             s/o Shri Ramesh Kotiya " & vbCrLf & _
"                             B.E. 2nd Year (Information Technology) " & vbCrLf & _
"                             Add:- S-2 Shrimaya Apart Sector - B/363 " & vbCrLf & _
"                                        Sarvdharm Colony, Bhopal (India)" & vbCrLf & _
"                             Fone:- +91-0755-2794428" & vbCrLf & _
"                             Web:- http:\\vinodkotiya.tripod.com " & vbCrLf & _
"                                   http:\\vinsoftindia.tripod.com " & vbCrLf & _
"                             Email:- vinodkotiya24@rediffmail.com" & vbCrLf & _
"**********" & vbCrLf & _
" Please send your complain's and suggestions." & vbCrLf & _
"//////////////////////////////////////////////////////////////////////////////"

End If
End Sub

Private Sub cmdAbout_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
 lblScan.Caption = "Click to get Help."
Else
 lblScan.Caption = "About VIN Cryptograph."
End If
cmdAbout(Index).BackColor = &HFFC0C0
End Sub

Private Sub cmdGo_Click(Index As Integer)
If Trim(txtKey.Text) = "" Then
 If Index = 0 Then
 MsgBox "Please Enter Your Key for Encryption."
 Else
 MsgBox "Please Enter Your Key for Decryption.This key must be similar to the key as time of encryption."
 End If
Exit Sub
End If
Dim isSourcefileSelected As Boolean
Dim X As Integer
isSourcefileSelected = False
Dim i As Integer
For i = 0 To File1.ListCount - 1
   If File1.Selected(i) = True Then
      isSourcefileSelected = True
       Exit For
    End If
Next
If isSourcefileSelected = False Then
 'Command2.SetFocus
   MsgBox "First Select the file to be Encrypted/Decrypted." & Chr(13) & _
   "To select any file click it in source file box."
   Exit Sub 'no file selected
End If
   '///////////////
   
   CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   CommonDialog1.FileName = ""
   If Index = 0 Then
    CommonDialog1.Filter = "vin Files(*.vin)|*.vin"
   Else
    CommonDialog1.Filter = "All Files|*.*"
   End If
   CommonDialog1.ShowSave
   If CommonDialog1.FileName = "" Then
      txtoutfile.Text = "No file is selected."
      Exit Sub
   End If
   outputfile = CommonDialog1.FileName '*.*.vin
   txtoutfile.Text = outputfile
cmdAbort.Enabled = True
cmdGo(0).Enabled = False
cmdGo(1).Enabled = False
starttime = Now
X = XORED(inputfile, outputfile, txtKey.Text)
Dim endtime As Integer
endtime = DateDiff("s", starttime, Now)
lblTime(0).Caption = endtime & " Sec"
cmdGo(0).Caption = "Session"
cmdGo(1).Caption = "Log Out"
If Index = 0 Then
MsgBox "File " & inputfile & " is Encrypted in to File " & outputfile & _
 " in " & endtime & " Seconds."

Else
 MsgBox "File " & inputfile & " is Decrypted in to File " & outputfile & _
 " in " & endtime & " Seconds."
 
End If
'Load frmReset
'frmReset.Show
'Unload Me

End Sub

Private Sub cmdGo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdGo(Index).BackColor = &HFFC0C0
If Index = 0 Then
lblScan.Caption = "Click to Encrypt the File " & inputfile
Else
lblScan.Caption = "Click to Decrypt the File " & outputfile
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblScan.Caption = "Click to select the folder.."
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    
End Sub

Private Sub File1_Click()
Dim flen As Long
If Right(Dir1.Path, 1) = "\" Then
    inputfile = Dir1.Path & File1.FileName
Else
    inputfile = Dir1.Path + "\" + File1.FileName
End If
txtInfile.Text = inputfile
flen = FileLen(inputfile)
If flen < 1024 Then
txtSize.Text = flen & " Bytes"
ElseIf flen >= 1024 And flen < 1048576 Then
txtSize.Text = Round(flen / 1024) & " KB"
Else
txtSize.Text = Round(flen / 1048576) & " MB"
End If
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblScan.Caption = "Click to Select the Input File"
End Sub

Private Sub Form_Load()
 
 Dir1.Path = App.Path
 isAbort = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAbort.BackColor = vbWhite
cmdGo(0).BackColor = vbWhite
cmdGo(1).BackColor = vbWhite
cmdAbout(0).BackColor = vbWhite
cmdAbout(1).BackColor = vbWhite
End Sub

Private Sub optKey_Click()
End Sub

Private Sub Label_Click(Index As Integer)
cmdAbout_Click (1)
End Sub

Private Sub txtInfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblScan.Caption = "Show the path of input File " & File1.List(File1.ListIndex)
End Sub

Private Sub txtKey_Change()
If Len(txtKey.Text) > 0 Then
 txtKey.BackColor = &HFFFFC0
Else
txtKey.BackColor = vbWhite
End If
End Sub

Private Sub txtoutfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblScan.Caption = "Shows the path of output File " & outputfile
End Sub

Private Sub txtSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblScan.Caption = "Shows the size of Input File " & File1.List(File1.ListIndex)
End Sub
