VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Instant Messanger"
   ClientHeight    =   7530
   ClientLeft      =   375
   ClientTop       =   495
   ClientWidth     =   6645
   Icon            =   "vinChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "How To Start"
      TabPicture(0)   =   "vinChat.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image2(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label15"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label16"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label17"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "TCP/IP Connection Setting"
      TabPicture(1)   =   "vinChat.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblstatus"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Image2(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label11(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label11(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtversion"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Option2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Option1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtcomp"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtMe"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "VIN Chat"
      TabPicture(2)   =   "vinChat.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image2(3)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtCHAT"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtstatus"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtmsg"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Command1"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Command2"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "vinChat.frx":2D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image2(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Inet1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtlocal"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.TextBox txtMe 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   -72000
         TabIndex        =   36
         Text            =   "Me  :"
         Top             =   3840
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save Chat"
         Height          =   375
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6120
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close Current Connection"
         Height          =   375
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox txtcomp 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   -72000
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "vinChat.frx":2D6A
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox txtlocal 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   5055
         Left            =   -74860
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "vinChat.frx":2D74
         Top             =   1680
         Width           =   6480
      End
      Begin VB.Frame Frame1 
         Caption         =   "Server:"
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   -74790
         TabIndex        =   14
         Top             =   1470
         Width           =   3015
         Begin VB.TextBox txtserverport 
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            Text            =   "5001"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdlisten 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Listen"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
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
         Height          =   2040
         Left            =   -71760
         TabIndex        =   8
         Top             =   1440
         Width           =   3255
         Begin VB.TextBox txtserverip 
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Text            =   "127.0.0.1"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtclientport 
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Text            =   "5001"
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton cmdconnect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Connect"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73470
         TabIndex        =   7
         Top             =   1110
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Client"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -69510
         TabIndex        =   6
         Top             =   1110
         Width           =   735
      End
      Begin VB.TextBox txtmsg 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   645
         Left            =   -74760
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   4320
         Width           =   6375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   405
         Left            =   -69960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
         Width           =   1455
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
         Top             =   6720
         Width           =   6375
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click to Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5520
         Width           =   2895
      End
      Begin RichTextLib.RichTextBox txtCHAT 
         Height          =   3495
         Left            =   -74790
         TabIndex        =   3
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6165
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"vinChat.frx":3039
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
         Left            =   -69000
         Top             =   750
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.TextBox txtversion 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "1.0.0"
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Your computer is known as:"
         ForeColor       =   &H000000C0&
         Height          =   495
         Index           =   1
         Left            =   -74640
         TabIndex        =   37
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   $"vinChat.frx":30BF
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   4560
         Width           =   6255
      End
      Begin VB.Label Label11 
         Caption         =   "The other computer is known as:"
         ForeColor       =   &H000000C0&
         Height          =   495
         Index           =   0
         Left            =   -74640
         TabIndex        =   32
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   7440
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Index           =   3
         Left            =   -71400
         Picture         =   "vinChat.frx":31AA
         Top             =   5040
         Width           =   540
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Index           =   2
         Left            =   -74880
         Picture         =   "vinChat.frx":34B4
         Top             =   480
         Width           =   540
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Index           =   1
         Left            =   -74790
         Picture         =   "vinChat.frx":37BE
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label17 
         Caption         =   "Then click either ""listen"" (if you are the server), or click ""connect"" (if you are a client). "
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3600
         Width           =   7335
      End
      Begin VB.Label Label16 
         Caption         =   "Then type in your friend's computer IP address, the port you are using, and the name you will chat with."
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3120
         Width           =   7335
      End
      Begin VB.Label Label15 
         Caption         =   "Third, click on the button ""Click to start"" below in Friends Chat !"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   7335
      End
      Begin VB.Label Label14 
         Caption         =   "Second, start this program."
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   7335
      End
      Begin VB.Label Label3 
         Caption         =   "First, connect to the internet ."
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   7335
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Index           =   0
         Left            =   210
         Picture         =   "vinChat.frx":3AC8
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "You can chat with ur friend over net or LAN if both of you know each others IP Address"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   " If you are the client, your friend will have to choose server, and vice versa.  Have fun!"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4080
         Width           =   7215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Index           =   0
         X1              =   90
         X2              =   7410
         Y1              =   7350
         Y2              =   7350
      End
      Begin VB.Label Label5 
         Caption         =   "Please enter the appropriate information below in order to establish a connection."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74190
         TabIndex        =   21
         Top             =   510
         Width           =   6615
      End
      Begin VB.Label Label6 
         Caption         =   "What are you going to work as :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74190
         TabIndex        =   20
         Top             =   870
         Width           =   3135
      End
      Begin VB.Label lblstatus 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"vinChat.frx":3DD2
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
         Height          =   1455
         Left            =   -74880
         TabIndex        =   19
         Top             =   5640
         Width           =   6345
      End
      Begin VB.Label Label10 
         Caption         =   "Type Message Here: "
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   4080
         Width           =   1935
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
' ***      *** ***   *****  ***   *******    *******
'  ***    ***  ***   *****  ***  ***   ***   ***  ****
'   ***  ***   ***   *** ** ***  ***   ***   ***   ****
'    ******    ***   ***  *****  ***   ***   ***  ****
'     ****     ***   ***   ****   *******    *******
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Programmer : VINOD KOTIYA
'  B.E. (Information Technology)
'  Semester IV
'  University Institute of Technology
'  Rajeev Gandhi Prodyogiki Vishwavidyalaya Bhopal.
'  Address: S-2 ShreeMaya Apartment Sector-B/363
'           Sarvdharm Colony Bhopal-42 (India)
'  Email: vinodkotiya@yahoo.co.in
'  Web : http://vinodkotiya.tripod.com
'  cell: +91-9827394994
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Created on march 2003

Option Explicit
Dim ref As Boolean

Private Sub cmdconnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdconnect.BackColor = &HC0C0C0
ref = True
End Sub

Private Sub cmdlisten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdlisten.BackColor = &HC0C0C0
ref = True
End Sub

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

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HC0C0C0
ref = True
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

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HC0C0C0
ref = True
End Sub

Private Sub Command3_Click()
If txtmsg.Text <> "" Then
    On Error GoTo err:
    Winsock1.SendData txtmsg.Text
    txtCHAT.Text = txtCHAT.Text & txtMe.Text & txtmsg.Text & vbCrLf
    txtmsg.Text = ""
    txtCHAT.SelStart = Len(txtCHAT.Text)
End If
Exit Sub
err:
    txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
    txtstatus.SelStart = Len(txtstatus.Text)
End Sub



Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = &HE0E0E0
ref = True
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

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BackColor = &HC0C0C0
ref = True
End Sub



Private Sub Form_Load()
Dim LineOfText As String
Dim alltext As String
txtversion.Text = App.Major & "." & App.Minor & "." & App.Revision
txtlocal.Text = "Local Host Name (networking name): " & Winsock1.LocalHostName _
& vbCrLf & "Local IP Address: " & Winsock1.LocalIP & vbCrLf _
& "Local Port: " & Winsock1.LocalPort & vbCrLf _
& vbCrLf & "---------What is---------" & vbCrLf & vbCrLf _
& "An IP address?" & vbCrLf & "An IP address is the address you use when you are online; your computer's online address.  Other computers " _
& "can talk to you using this IP address.  Your LOCAL IP address is listed above, but you need to check your internet connection settings in your ISP's dialer for your IP address." _
& vbCrLf & vbCrLf & "A Port?" & vbCrLf & "A port is space in your computer reserved for connecting to other computers.  Most computers have 5000+ ports on their computer.  For example, VIN Instant Messanger asks the server for the port to 'listen' on, and you have to connect to that port so you can chat using VIN Instant Messanger" & vbCrLf & vbCrLf & "VIN Instant Messanger"
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
Winsock1.SendData "Bye ! I am shutting down..."
Winsock1.Close
End If
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
End Sub

Private Sub Image4_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Option1_Click()
Frame1.Enabled = True
Frame2.Enabled = False
txtcomp.Text = "Dost :"
lblstatus.Caption = "After you type in the port you will listen on, click listen to listen for users to connect.  Then click on the tab VIN Chat to chat when the other user connected."
txtserverport.SetFocus

'txtserverport.SelStart = 0
'txtserverport.SelLength = Len(txtserverport.Text)
End Sub

Private Sub Option2_Click()
Frame2.Enabled = True
Frame1.Enabled = False
txtcomp.Text = "Server :"
lblstatus.Caption = "After you type in the server IP address and the port you will connect to, click connect to connect to the server that is listening for you to connect.  NOTE:  There must be a server listening for clients to connect first or TCP connection will not work!"
txtserverip.SetFocus

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ref = True Then
Command5.BackColor = vbWhite
Command1.BackColor = vbWhite
Command2.BackColor = vbWhite
Command3.BackColor = vbWhite
cmdlisten.BackColor = vbWhite
cmdconnect.BackColor = vbWhite
ref = False
End If
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
