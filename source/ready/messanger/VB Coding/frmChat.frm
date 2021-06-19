VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "My Chat"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7935
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   0
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "MyChat-Saveas"
   End
   Begin VB.ListBox lstUsers 
      Height          =   3435
      IntegralHeight  =   0   'False
      ItemData        =   "frmChat.frx":1CFA
      Left            =   5760
      List            =   "frmChat.frx":1CFC
      Style           =   1  'Checkbox
      TabIndex        =   6
      ToolTipText     =   "Select the users to talk with"
      Top             =   0
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8943
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
            Object.ToolTipText     =   "Time (in sec) elapsed in this session"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   "12:28 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrConnection 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   3360
   End
   Begin VB.TextBox txtText 
      Height          =   510
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3720
      Width           =   6735
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   480
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmChat.frx":1CFE
   End
   Begin RichTextLib.RichTextBox rt2 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmChat.frx":1DC7
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtDisplay 
      Height          =   3435
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6059
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":1E90
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnusave 
         Caption         =   "Save &Chat"
      End
      Begin VB.Menu sndfile 
         Caption         =   "&Send File"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuPrompt 
         Caption         =   "&Prompt"
         Visible         =   0   'False
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu con 
      Caption         =   "&Connection"
      Begin VB.Menu mnuDiscon 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnurecon 
         Caption         =   "&Reconnect"
         Visible         =   0   'False
      End
      Begin VB.Menu spacerr 
         Caption         =   "-"
      End
      Begin VB.Menu mnumcast 
         Caption         =   "&Multicast"
      End
      Begin VB.Menu mnuPM 
         Caption         =   "&Private Message"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileName(10) As String, filedata(10) As String, nofile As Integer
Dim fileno As Integer
Dim minutes As Integer, seconds As Integer, nTime As Long
Private SvrIPAdd As String
Dim strmulticast As String

Function Connect(svrIP As String)
wskClient.Close
wskClient.Connect svrIP, "5001"
SvrIPAdd = svrIP
StBar.Panels(1).Text = "Connecting..."
End Function

Private Sub cmdSend_Click()
If wskClient.State = sckConnected Then
    If txtText.Text <> "" Then
        If mnumcast.Checked = False Then
            Dim allText As String
            allText = txtUser.Text & ":" & txtText.Text
            DoEvents
            Call wskClient.SendData("Message " & allText)
        Else
            wskClient.SendData ("Multicast " & txtUser & " " & strmulticast & " " & txtText)
        End If
            If PlaySnd = True Then
                Call PlayWav("Send.wav")
            End If
    Else
        Call AddText("  ~// You must enter text to send it")
    End If
Else
    Call AddText("  ~// You must be connected to some to send text")
End If

txtText.Text = ""

End Sub


Private Sub Form_Load()
nTime = 0
strmulticast = ""
If Prompt = True Then
    mnuPrompt.Visible = True
End If
cmdSend.Enabled = False
End Sub

Private Sub Form_Resize()
If (frmChat.Height > 2000) And (frmChat.Width > lstUsers.Width + 100) Then
    txtDisplay.Width = frmChat.Width - lstUsers.Width - 100
    txtDisplay.Height = frmChat.Height - 2000
    lstUsers.ToolTipText = txtDisplay.Top
    lstUsers.Height = txtDisplay.Height
    lstUsers.Left = txtDisplay.Left + txtDisplay.Width
    txtText.Top = (txtDisplay.Height + frmChat.Height - StBar.Height - txtText.Height - 600) / 2
    txtText.Width = txtDisplay.Width
    cmdSend.Top = txtText.Top
    cmdSend.Left = txtText.Width + txtText.Left
    txtDisplay.Height = lstUsers.Height
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If wskClient.State = sckConnected Then
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
End If
End Sub

Private Sub lstUsers_ItemCheck(Item As Integer)
If Item > 0 Then
strmulticast = Left(strmulticast, Item) & CStr(-1 * CInt(Not CBool(Mid(strmulticast, Item + 1, 1)))) & Right(strmulticast, Len(strmulticast) - Item - 1)
Else
strmulticast = CStr(-1 * CInt(Not CBool(Left$(strmulticast, 1)))) & Right$(strmulticast, Len(strmulticast) - 1)
End If
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If lstUsers.SelCount = "0" Then
Exit Sub
Else
If Button = 1 Then
    Exit Sub
Else
    Call Me.PopupMenu(mnu, , lstUsers.Left + 50)
End If
End If
End Sub


Public Sub wskClient_Close()
mnuDiscon.Caption = "&Connect"
StBar.Panels(1).Text = "Closed"
Call AddText("  ~// The connection was unexpectedly dropped.")
lstUsers.Clear
txtDisplay.Text = ""
End Sub

Private Sub wskClient_Connect()
nTime = CLng(Timer)
mnuDiscon.Caption = "&Disconnect"
mnurecon.Visible = True
tmrConnection.Enabled = True
StBar.Panels(1).Text = "Connected"
Call wskClient.SendData("Join " & txtUser.Text)
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
wskClient.GetData strData, vbString
Call DataParsing(strData)

End Sub

Function DataParsing(strData As String)
On Error Resume Next
strack = strData
Dim Command As String, Info As String, User As String, Text As String

Command$ = Left$(strData$, InStr(strData$, " ") - 1)
Info = Right(strData$, Len(strData$) - InStr(strData$, " "))

Select Case Command

Case "Message":
    User$ = Left$(Info$, InStr(Info$, ":") - 1)
    Text$ = Right$(Info$, Len(Info$) - InStr(Info$, ":"))
    DoEvents
    
    Call AddChat(User, Text)
    DoEvents
    If Not User = txtUser.Text Then
        If PlaySnd = True Then
            Call PlayWav("Recv.wav")
        End If
    End If
    
Case "PMMessage":
    User$ = Left$(Info$, InStr(Info$, "|") - 1)
    Text$ = Right$(Info, Len(Info) - InStr(Info, "|"))
    DoEvents
    
    Call AddChatPM(User, Text)
    DoEvents
    
Case "ErrUsername":
    Dim Answer As VbMsgBoxResult
    Answer = MsgBox("The username " & Info & " is in use." & vbCrLf & "Would you like to enter a new one.", vbYesNo, "Username Error")
        If Answer = vbYes Then
            Dim strUser As String
            strUser = InputBox("Please enter a new username.", "New Username")
            If strUser <> "" Then
            Call wskClient.SendData("Join " & strUser)
            DoEvents
            txtUser.Text = strUser
            Me.Caption = "MyChat - (" & txtUser.Text & ")"
            End If
        ElseIf Answer = vbNo Then
            wskClient.Close
            frmSignon.Show
            DoEvents
            Unload Me
        End If

Case "PMError":
    Call AddText("  ~// User: " & Info & " doesn't exist, or has left")
    DoEvents

Case "UserList":
    wskClient.SendData ("ACKLIST")
    DoEvents
    Call lstUsers.AddItem(Info)
    strmulticast = strmulticast & "0"
    
    
Case "Joined":
    Call lstUsers.AddItem(Info)
    strmulticast = strmulticast & "0"
    DoEvents
    Call AddText("  ~// User: " & Info & " has joined the chat")
    
Case "Left":
    For I = 0 To lstUsers.ListCount
        If lstUsers.List(I) = Info Then
            Call lstUsers.RemoveItem(I)
            strmulticast = Left(strmulticast, I) & Right$(strmulticast, Len(strmulticast) - I - 1)
            DoEvents
        End If
        DoEvents
    Next I
    Call AddText("  ~// User: " & Info & " left the chat room")
    
Case "Kicked":
    For I = 0 To lstUsers.ListCount
        If lstUsers.List(I) = Info Then
            Call lstUsers.RemoveItem(I)
            DoEvents
        End If
        DoEvents
    Next I
    
    Call AddText("  ~// User: " & Info & " was kicked")
    
Case "UKicked":
    Call AddText("  ~// You have been kicked by: " & Info)

Case "File":
    Dim recindex As Integer, progress As String
    
    User$ = Left$(Info$, InStr(Info$, "#@%$") - 1)
    Text = Right(Info$, Len(Info$) - InStr(Info$, "#@%$") - 3)
    progress = Right(Text$, Len(Text$) - InStr(Text$, "#@%$") - 3)
    
    For recindex = 0 To nofile
        If (FileName(recindex) = Left$(Text$, InStr(Text$, "#@%$") - 1)) Then
            If ((progress <> "START") And (progress <> "END")) Then
                filedata(recindex) = filedata(recindex) & progress
                Debug.Print filedata(recindex)
                wskClient.SendData ("ACKFILE")
                DoEvents
            End If
        End If
        Exit For
    Next recindex
    
    DoEvents
    
    
    If progress = "END" Then
        Print #fileno, filedata(recindex)
        Close #fileno
        filedata(recindex) = ""
        FileName(recindex) = ""
        progress = ""
        
        wskClient.SendData ("ACKENDFILE")
        DoEvents
        nofile = nofile - 1
    End If
    
Case "Faccept":
    
    Dim NameFile As String
    User$ = Left$(Info$, InStr(Info$, " ") - 1)
    Text$ = Right(Info$, Len(Info$) - InStr(Info$, " "))
    NameFile = Left$(Text$, InStr(Text$, " ") - 1)
    accept = CInt(Left$(Text$, InStr(Text$, " ") - 1))
    Text$ = Right(Text$, Len(Text$) - InStr(Text$, " "))
    
    If Trim(Text$) = "ASK" Then
        msg$ = "User: " & User$ & " has sent you a file '" & NameFile & "'." & Chr(13) & "Do you want to recieve it?"
        If MsgBox(msg$, vbYesNo, "MyChat-File Recieve") = vbYes Then
            
            FileName(nofile) = NameFile
            ComDlg.Flags = cdlOFNOverwritePrompt
            ComDlg.Filter = "All Files(*.*)|*.*"
            ComDlg.FileName = FileName(nofile)
            ComDlg.ShowSave
            fileno = FreeFile
            If ComDlg.CancelError = False Then
                Open ComDlg.FileName For Output As #fileno
                Call wskClient.SendData("Faccept " & txtUser.Text & " 1 TELL")
                nofile = (nofile Mod 32000) + 1
            Else
                Call wskClient.SendData("Faccept " & txtUser.Text & " 0 TELL")
            End If
            
        Else
            Call wskClient.SendData("Faccept " & txtUser.Text & " 0 TELL")
        End If
        
    ElseIf Trim(Text$) = "TELL" Then
        
        If accept = 0 Then
            MsgBox "User: " & User & " does not want to accept file"
        End If
                        
    End If
    
End Select

End Function

Function AddText(Text As String)

rt2.SelStart = 0
rt2.SelLength = 0
rt2.TextRTF = Text
rt2.SelStart = 2
rt2.SelLength = Len(Text)
rt2.SelColor = &H8000&
rt2.SelStart = 0
rt2.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt2.TextRTF & vbCrLf
If DisplayCorr = True Then
    txtDisplay.SelStart = Len(txtDisplay)
    txtDisplay.SelRTF = vbCrLf
End If
DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt2.TextRTF = ""
End Function

Function AddChat(User As String, Text As String)
rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = User & ": " & Text
rt.SelStart = 0
rt.SelLength = Len(User) + 1
If User = txtUser.Text Then
rt.SelColor = vbRed
Else
rt.SelColor = vbBlue
End If
rt.SelBold = True
rt.SelStart = 0
rt.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt.TextRTF & vbCrLf

If DisplayCorr = True Then
    txtDisplay.SelStart = Len(txtDisplay)
    txtDisplay.SelRTF = vbCrLf
End If

DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt.TextRTF = ""
End Function

Function AddChatPM(User As String, Text As String)
rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = User & ": " & Text
rt.SelStart = 0
rt.SelLength = Len(User) + 1
rt.SelColor = &H8000&
rt.SelBold = True
rt.SelStart = 0
rt.SelLength = 0
DoEvents
txtDisplay.SelStart = Len(txtDisplay)
DoEvents
txtDisplay.SelRTF = rt.TextRTF & vbCrLf
If DisplayCorr = True Then
    txtDisplay.SelStart = Len(txtDisplay)
    txtDisplay.SelRTF = vbCrLf
End If
DoEvents
txtDisplay.SelStart = Len(txtDisplay)

rt.TextRTF = ""
End Function


Private Sub mnuDiscon_Click()
If mnuDiscon.Caption = "&Disconnect" Then
    mnuDiscon.Caption = "&Connect"
    mnurecon.Visible = False
    If wskClient.State = sckConnected Then
        Call wskClient.SendData("Leave " & txtUser.Text)
        DoEvents
        wskClient.Close
    End If
    DoEvents
    frmSignon.txtUser.Text = txtUser.Text
    frmSignon.Show
    DoEvents
    
    Unload Me
Else
    wskClient.Close
    wskClient.Connect SvrIPAdd, "1290"
    StBar.Panels(1).Text = "Connecting..."
End If

End Sub

Private Sub mnuExit_Click()
If wskClient.State = sckConnected Then
Call wskClient.SendData("Leave " & txtUser.Text)
DoEvents
End If
End
End Sub

Private Sub mnumcast_Click()
If mnumcast.Checked = True Then
    mnumcast.Checked = False
Else
    mnumcast.Checked = True
End If

End Sub

Private Sub mnuOpt_Click()
frmOptions.Show
End Sub

Private Sub mnuPM_Click()
If lstUsers.Text = txtUser.Text Then
    Call AddText("  ~// You can't Private message yourself")
Else
frmPM.lblUserPM.Caption = lstUsers.Text
frmPM.Show
End If
End Sub

Private Sub mnuPrompt_Click()
frmPrompt.Show
End Sub

Private Sub mnurecon_Click()
If wskClient.State = sckConnected Then
    Call wskClient.SendData("Leave " & txtUser.Text)
    DoEvents
End If
wskClient.Close
wskClient.Connect SvrIPAdd, "1290"
StBar.Panels(1).Text = "Connecting..."
End Sub

Private Sub mnusave_Click()
    
    Dim filno As Integer
    Dim txt As String
    txt = txtDisplay.Text
    ComDlg.Flags = cdlOFNOverwritePrompt
    ComDlg.Filter = "Text Documents|*.txt|All Files(*.*)|*.*"
    ComDlg.FileName = "Chat_Safe.txt"
    ComDlg.ShowSave
    If ComDlg.CancelError = False Then
        filno = FreeFile
        Open ComDlg.FileName For Output As #filno
        Print #filno, txt
        Close #filno
    End If
End Sub

Private Sub sndfile_Click()
    Frmsend.Show (vbModal)
    
End Sub

Private Sub tmrConnection_Timer()
If seconds = 59 Then
    minutes = minutes + 1
End If
seconds = (CLng(Timer) - nTime) Mod 60

StBar.Panels(2).Text = minutes & "min " & seconds & "sec"
End Sub

Private Sub txtText_Change()
If txtText.Text > "" Then
    cmdSend.Enabled = True
Else
    cmdSend.Enabled = False
End If
End Sub


Private Sub txtText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub

