VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmsend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyChat-SendFile"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "Frmsend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Send a File"
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdcancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Timer TmrSend 
         Interval        =   500
         Left            =   2400
         Top             =   3720
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "Frmsend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SOFrame As Integer = 8000 'Size of data to be read at a time from file

Private Sub cmdCancel_Click()
Me.Hide 'Cancel the process
End Sub

Private Sub cmdSend_Click()

If File1.ListCount <> 0 Then
    Dim fnum As Integer, parts As Long, location As Long ' File variables
    Dim Path As String, remain As String  'path of file 2b sent
    Dim fname As String, File As String 'name and contents of file
    
    
    Call frmChat.wskClient.SendData("Faccept " & frmChat.txtUser.Text _
                    & " " & File1.FileName & " ASK")       'Ask to accept file
    
    While (accept <> 1)
        DoEvents
        Call frmChat.wskClient.GetData(remain)
        Call frmChat.DataParsing(remain)
        
    Wend
    accept = 0
    
    If Right(File1.Path, 1) = "\" Then
        Path = File1.Path & File1.FileName
    Else
        Path = File1.Path & "\" & File1.FileName
    End If
    
    fnum = FreeFile
    Open Path For Binary As #fnum
    
    ProgressBar1.Max = LOF(fnum)
    ProgressBar1.Value = 0
    parts = CInt(LOF(fnum) / SOFrame)
    If parts > (LOF(fnum) / SOFrame) Then parts = parts - 1
    
    Do While (LOF(fnum) - location) > SOFrame ' Loop until end of file.
        File = Input(SOFrame, #fnum) ' Read character into variable.
        frmChat.wskClient.SendData ("File " & Trim(frmChat.txtUser.Text) _
            & "#@%$" & File1.FileName & "#@%$" & File)
        ProgressBar1.Value = ProgressBar1.Value + SOFrame
        
        While (strack <> "ACKFILE")
            Call frmChat.wskClient.GetData(strack)
            Call frmChat.DataParsing(strack)
            'frmChat.wskClient.SendData ("File " & Trim(frmChat.txtUser.Text) _
                & "#@%$" & File1.FileName & "#@%$END")
            DoEvents
        Wend
        strack = ""
        DoEvents
        location = Loc(fnum) ' Get current position within file.
        Debug.Print File ' Print to the Immediate window.
    Loop

    File = Input(LOF(fnum) - ((parts) * CLng(SOFrame)), #fnum)
    frmChat.wskClient.SendData ("File " & Trim(frmChat.txtUser.Text) _
        & "#@%$" & File1.FileName & "#@%$" & File)
    DoEvents
    While (strack <> "ACKFILE")
        Call frmChat.wskClient.GetData(strack)
        Call frmChat.DataParsing(strack)
        'frmChat.wskClient.SendData ("File " & Trim(frmChat.txtUser.Text) _
        & "#@%$" & File1.FileName & "#@%$END")
        DoEvents
    Wend
    strack = ""
    Debug.Print File
    Close #fnum
            
    frmChat.wskClient.SendData ("File " & Trim(frmChat.txtUser.Text) _
        & "#@%$" & File1.FileName & "#@%$END")
    DoEvents
    While (strack <> "ACKENDFILE")
        Call frmChat.wskClient.GetData(strack)
        Call frmChat.DataParsing(strack)
        
        DoEvents
    Wend
    strack = ""
    Debug.Print File
    
End If
cmdSend.Enabled = False
Me.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
    File1.Selected(0) = True
End If

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
cmdSend.Enabled = True
End Sub

Private Sub Form_Load()
Drive1.Drive = "c:"
End Sub

Private Sub TmrSend_Timer()
strack = "ACKFILE"
End Sub
