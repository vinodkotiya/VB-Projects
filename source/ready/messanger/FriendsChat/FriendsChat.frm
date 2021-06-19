VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Friends Chat !"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "FriendsChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000018&
      Caption         =   "(!)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Switch this button to reload the system, Power reboot1"
      Top             =   1080
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483644
      ForeColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Directions"
      TabPicture(0)   =   "FriendsChat.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label17"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "TCP/IP Connection Setup"
      TabPicture(1)   =   "FriendsChat.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Option1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Option2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblstatus"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Image4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Friends Chat !"
      TabPicture(2)   =   "FriendsChat.frx":12AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtmsg"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtstatus"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtCHAT"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label10"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Image1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Options"
      TabPicture(3)   =   "FriendsChat.frx":12C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label11"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label12"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Inet1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtcomp"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtversion"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Command2"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Command1"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtlocal"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Command6"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).ControlCount=   11
      Begin VB.CommandButton Command6 
         Caption         =   "< Back to chat"
         Height          =   375
         Left            =   -74910
         TabIndex        =   30
         Top             =   4950
         Width           =   2415
      End
      Begin VB.TextBox txtlocal 
         BackColor       =   &H8000000C&
         ForeColor       =   &H0000FFFF&
         Height          =   3255
         Left            =   -74860
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "FriendsChat.frx":12E2
         Top             =   1680
         Width           =   7335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close Current Connection"
         Height          =   375
         Left            =   -69870
         TabIndex        =   28
         Top             =   4950
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save Chat"
         Height          =   375
         Left            =   -72390
         TabIndex        =   27
         Top             =   4950
         Width           =   2415
      End
      Begin VB.TextBox txtversion 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   -71670
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "1.0.0"
         Top             =   1170
         Width           =   2775
      End
      Begin VB.TextBox txtcomp 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   525
         Left            =   -71670
         MaxLength       =   20
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Text            =   "FriendsChat.frx":166B
         Top             =   570
         Width           =   4095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Server:"
         ForeColor       =   &H000000FF&
         Height          =   3015
         Left            =   -74790
         TabIndex        =   14
         Top             =   1470
         Width           =   3615
         Begin VB.TextBox txtserverport 
            BackColor       =   &H8000000C&
            ForeColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            Text            =   "5001"
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdlisten 
            Caption         =   "&Listen"
            Height          =   375
            Left            =   840
            TabIndex        =   15
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label7 
            Caption         =   "Port to listen on:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Client:"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   3015
         Left            =   -71070
         TabIndex        =   8
         Top             =   1460
         Width           =   3615
         Begin VB.TextBox txtserverip 
            BackColor       =   &H8000000C&
            ForeColor       =   &H00C0C0FF&
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtclientport 
            BackColor       =   &H8000000C&
            ForeColor       =   &H00C0C0FF&
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Text            =   "5001"
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton cmdconnect 
            Caption         =   "&Connect"
            Height          =   375
            Left            =   840
            TabIndex        =   9
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Server IP address:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Port to connect to:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Server"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -73470
         TabIndex        =   7
         Top             =   1110
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Client"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -69510
         TabIndex        =   6
         Top             =   1110
         Width           =   735
      End
      Begin VB.TextBox txtmsg 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -73470
         TabIndex        =   5
         Top             =   690
         Width           =   4935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   405
         Left            =   -68430
         TabIndex        =   4
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox txtstatus 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   590
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   4950
         Width           =   7335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00008000&
         Caption         =   "Lets BEGIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3150
         Width           =   2175
      End
      Begin RichTextLib.RichTextBox txtCHAT 
         Height          =   3495
         Left            =   -74790
         TabIndex        =   3
         Top             =   1350
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         _Version        =   393217
         BackColor       =   12632319
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"FriendsChat.frx":1672
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   -68430
         Top             =   750
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label Label17 
         Caption         =   "Then click either ""listen"" (if you are the server), or click ""connect"" (if you are a client). "
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   210
         TabIndex        =   38
         Top             =   2400
         Width           =   7335
      End
      Begin VB.Label Label16 
         Caption         =   "Then type in your friend's computer IP address, the port you are using, and the name you will chat with."
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   210
         TabIndex        =   37
         Top             =   2160
         Width           =   7335
      End
      Begin VB.Label Label15 
         Caption         =   "Third, click on the button ""Click to start"" below in Friends Chat !"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   210
         TabIndex        =   36
         Top             =   1920
         Width           =   7335
      End
      Begin VB.Label Label14 
         Caption         =   "Second, start this program."
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   210
         TabIndex        =   35
         Top             =   1680
         Width           =   7335
      End
      Begin VB.Label Label3 
         Caption         =   "First, connect to the internet."
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   210
         TabIndex        =   34
         Top             =   1440
         Width           =   7335
      End
      Begin VB.Label Label13 
         Caption         =   "What is my..."
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   -74910
         TabIndex        =   33
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Friends Chat ! current version:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -74190
         TabIndex        =   32
         Top             =   1170
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "The other computer is known as:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -74190
         TabIndex        =   31
         Top             =   570
         Width           =   2415
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   -74790
         Picture         =   "FriendsChat.frx":175C
         Top             =   570
         Width           =   540
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   210
         Picture         =   "FriendsChat.frx":1B9E
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   $"FriendsChat.frx":1FE0
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   810
         TabIndex        =   24
         Top             =   570
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   " If you are the client, your friend will have to choose server, and vice versa.  Have fun!"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   7215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         X1              =   90
         X2              =   7410
         Y1              =   4230
         Y2              =   4230
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Copyright © 2001.  All rights reserved.  Any questions, comments, ideas, etc. should be sent to sonal3k@yahoo.com."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   975
         Left            =   90
         TabIndex        =   22
         Top             =   4350
         Width           =   7335
      End
      Begin VB.Label Label5 
         Caption         =   "Please enter the appropriate information below in order to establish a connection."
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74190
         TabIndex        =   21
         Top             =   510
         Width           =   6615
      End
      Begin VB.Label Label6 
         Caption         =   "What are you going to work as :"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74190
         TabIndex        =   20
         Top             =   870
         Width           =   3135
      End
      Begin VB.Label lblstatus 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FriendsChat.frx":20D2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   -74790
         TabIndex        =   19
         Top             =   4590
         Width           =   7335
      End
      Begin VB.Label Label10 
         Caption         =   "Message:"
         Height          =   255
         Left            =   -74190
         TabIndex        =   18
         Top             =   690
         Width           =   735
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   -74790
         Picture         =   "FriendsChat.frx":2198
         Top             =   570
         Width           =   540
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   -74790
         Picture         =   "FriendsChat.frx":25DA
         Top             =   570
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.txt"
      DialogTitle     =   "Save your Chat !"
      FileName        =   "Friends Chat"
      InitDir         =   "c:\Friends Chat"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err:
Winsock1.SendData "Hay Buddy, I am leaving .... "
Winsock1.Close
Frame1.Enabled = True
Frame2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
SSTab1.Tab = 0
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub cmdconnect_Click()
On Error GoTo err:
Winsock1.RemoteHost = txtserverip.Text
Winsock1.RemotePort = txtclientport.Text
Winsock1.Connect
SSTab1.Tab = 2
txtmsg.SetFocus
Frame1.Enabled = False
Frame2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub cmdlisten_Click()
On Error GoTo err:
Winsock1.LocalPort = txtserverport.Text
Winsock1.Listen
SSTab1.Tab = 2
txtmsg.SetFocus
Frame1.Enabled = False
Frame2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Command2_Click()
If txtCHAT.Text <> "" Then
    CommonDialog1.Filter = "Text files (*.txt)|*.txt"
    CommonDialog1.ShowSave
        If CommonDialog1.FileName <> "" Then
            Open CommonDialog1.FileName For Output As #1
            Print #1, txtCHAT.Text
            Close #1
        End If
End If
End Sub

Private Sub Command3_Click()
If txtmsg.Text <> "" Then
    On Error GoTo err:
    Winsock1.SendData txtmsg.Text
    txtCHAT.Text = txtCHAT.Text & "Me: " & txtmsg.Text & vbCrLf
    txtmsg.Text = ""
    txtCHAT.SelStart = Len(txtCHAT.Text)
End If
Exit Sub
err:
    txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
    txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Command4_Click()
frmSplash.Show
Unload Me
End Sub

Private Sub Command5_Click()
SSTab1.Tab = 1
Option1.Value = True
Option2.Value = False
On Error GoTo err:
txtserverport.SetFocus
txtserverport.SelStart = 0
txtserverport.SelLength = Len(txtserverport.Text)
err:

End Sub

Private Sub Command6_Click()
SSTab1.Tab = 2
txtmsg.SetFocus
End Sub

Private Sub Form_Load()
txtversion.Text = App.Major & "." & App.Minor & "." & App.Revision
txtlocal.Text = "Local Host Name (networking name): " & Winsock1.LocalHostName _
& vbCrLf & "Local IP Address: " & Winsock1.LocalIP & vbCrLf _
& "Local Port: " & Winsock1.LocalPort & vbCrLf _
& vbCrLf & "---------What is---------" & vbCrLf & vbCrLf _
& "An IP address?" & vbCrLf & "An IP address is the address you use when you are online; your computer's online address.  Other computers " _
& "can talk to you using this IP address.  Your LOCAL IP address is listed above, but you need to check your internet connection settings in your ISP's dialer for your IP address." _
& vbCrLf & vbCrLf & "A Port?" & vbCrLf & "A port is space in your computer reserved for connecting to other computers.  Most computers have 5000+ ports on their computer.  For example, Friends Chat ! asks the server for the port to 'listen' on, and you have to connect to that port so you can chat using Friends Chat !" & vbCrLf & vbCrLf & "Friends Chat ! Copyright © 2003.  All rights reserved.  Any questions or comments should be sent to sonal3k@yahoo.com and will get a response within 48 hours."
CommonDialog1.FileName = App.Path & "\knownas.txt"
On Error GoTo toobig:
        Open CommonDialog1.FileName For Input As #1
        On Error GoTo toobig:    'set error handler
        Do Until EOF(1)          'then read lines from file
            Line Input #1, LineOfText$
            alltext$ = alltext$ & LineOfText$
        Loop
        txtcomp.Text = alltext$  'display file
        Close #1                 'close file
If Winsock1.State <> sckClosed Then
On Error GoTo err:
Winsock1.SendData "Other user connected!"
Winsock1.Close
End If
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
toobig:

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Winsock1.State <> sckClosed Then
On Error GoTo err:
Winsock1.SendData "Hay man ! I am going offline!  Catch you later..."
Winsock1.Close
End If
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Option1_Click()
Frame1.Enabled = True
Frame2.Enabled = False
lblstatus.Caption = "After you type in the port you will listen on, click listen to listen for users to connect.  Then click on the tab Friends Chat to chat when the other user connects."
txtserverport.SetFocus
'txtserverport.SelStart = 0
'txtserverport.SelLength = Len(txtserverport.Text)
End Sub

Private Sub Option2_Click()
Frame2.Enabled = True
Frame1.Enabled = False
lblstatus.Caption = "After you type in the server IP address and the port you will connect to, click connect to connect to the server that is listening for you to connect.  NOTE:  There must be a server listening for clients to connect first or TCP connection will not work!"
txtserverip.SetFocus
End Sub

Private Sub Winsock1_Close()
MsgBox "User Logged Out.", vbExclamation
DoEvents
Frame1.Enabled = True
Frame2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
SSTab1.Tab = 0
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Winsock1_Connect()
On Error GoTo err:
MsgBox "User Logged In!", vbExclamation
Winsock1.SendData "Communicaion channel LogOn process completed!"
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error GoTo err:
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
If Me.WindowState = 1 Then Me.WindowState = 0
Dim strdata As String
On Error GoTo err:
Winsock1.GetData strdata
txtCHAT.Text = txtCHAT.Text & txtcomp.Text & ": " & strdata & vbCrLf
txtCHAT.SelStart = Len(txtCHAT.Text)
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub
