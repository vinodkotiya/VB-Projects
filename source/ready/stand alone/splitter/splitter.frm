VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplit 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN Split & Merge v1.0"
   ClientHeight    =   6750
   ClientLeft      =   240
   ClientTop       =   570
   ClientWidth     =   6870
   Icon            =   "splitter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   26
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Split The Files"
      TabPicture(0)   =   "splitter.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line5(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line6(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line7(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command3(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command3(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "listFile"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Drive1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Dir1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtRemarks"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Timer1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdSplit"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chkMedia"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkDel"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkBatch"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "&Merge The Files"
      TabPicture(1)   =   "splitter.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdStopmerge"
      Tab(1).Control(2)=   "chkDelmerge"
      Tab(1).Control(3)=   "Text1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtRem"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Dir2"
      Tab(1).Control(6)=   "Drive2"
      Tab(1).Control(7)=   "cmdMerge"
      Tab(1).Control(8)=   "Command1"
      Tab(1).Control(9)=   "Line10"
      Tab(1).Control(10)=   "Line7(1)"
      Tab(1).Control(11)=   "Line6(1)"
      Tab(1).Control(12)=   "Line5(1)"
      Tab(1).Control(13)=   "Line2"
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(15)=   "Label6"
      Tab(1).Control(16)=   "Label5"
      Tab(1).Control(17)=   "Label4"
      Tab(1).Control(18)=   "Line4"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "&About.."
      TabPicture(2)   =   "splitter.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command6(2)"
      Tab(2).Control(1)=   "Command6(1)"
      Tab(2).Control(2)=   "frmBball"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command6(0)"
      Tab(2).Control(4)=   "Line1(1)"
      Tab(2).Control(5)=   "Line3(1)"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "About"
         Height          =   375
         Index           =   2
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Help"
         Height          =   855
         Index           =   1
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   600
         Width           =   1455
      End
      Begin VB.PictureBox frmBball 
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   -74760
         ScaleHeight     =   4095
         ScaleWidth      =   6375
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1440
         Width           =   6375
         Begin VB.PictureBox shpBall 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   2160
            ScaleHeight     =   495
            ScaleWidth      =   255
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1080
            Width           =   255
            Begin VB.Shape shpBall1 
               FillColor       =   &H00FF0000&
               FillStyle       =   0  'Solid
               Height          =   495
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Timer timBall 
            Left            =   1320
            Top             =   1320
         End
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   -71520
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "No file is selected for merging"
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox chkBatch 
         Caption         =   "&Create batch file for merging on other system"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Credit"
         Height          =   375
         Index           =   0
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdStopmerge 
         BackColor       =   &H0080FF80&
         Caption         =   "STOP"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Stop the merging any time."
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CheckBox chkDelmerge 
         Caption         =   "Delete Splitted File After Merging"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   -70320
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CheckBox chkDel 
         Caption         =   "&Delete Source File After Splitting"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox chkMedia 
         Caption         =   "&Prompt for Removble Media when full"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CommandButton cmdSplit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-Split-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   1320
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   735
         Left            =   -73440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "splitter.frx":035E
         Top             =   3240
         Width           =   5175
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00FFFFC0&
         Height          =   615
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "splitter.frx":03AA
         ToolTipText     =   "Write the remember points about source file here."
         Top             =   3000
         Width           =   5415
      End
      Begin VB.DirListBox Dir2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   -71760
         TabIndex        =   19
         Top             =   1440
         Width           =   3480
      End
      Begin VB.DriveListBox Drive2 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -73440
         TabIndex        =   18
         Top             =   1440
         Width           =   1680
      End
      Begin VB.CommandButton cmdMerge 
         BackColor       =   &H00FFFF00&
         Caption         =   "-Merge-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Click to merge the file."
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open "" *.*.vin "" file"
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
         Left            =   -73440
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Open the *.*.vin file for merging."
         Top             =   600
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Splitting Options : Selected file size =  0 bytes"
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   6615
         Begin VB.CommandButton cmdStopsplit 
            BackColor       =   &H0080FF80&
            Caption         =   "STOP"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   3960
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Displays the total processing time in splitting"
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optKb 
            Caption         =   "KB"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   840
            TabIndex        =   12
            ToolTipText     =   "1 KB = 1024 Bytes"
            Top             =   960
            Width           =   615
         End
         Begin VB.OptionButton optMb 
            Caption         =   "MB"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            ToolTipText     =   "1 MB = 1048576 Bytes"
            Top             =   960
            Width           =   615
         End
         Begin VB.OptionButton optByte 
            Caption         =   "&Byte"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.TextBox txtParts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            Height          =   285
            Left            =   3000
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   960
            Width           =   855
         End
         Begin VB.ComboBox cmbSize 
            BackColor       =   &H00FFFF80&
            Height          =   315
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Enter the size of 1 splitted part or select predefined sizes from dropdown."
            Top             =   600
            Width           =   3015
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FF0000&
            Index           =   2
            X1              =   3480
            X2              =   3480
            Y1              =   720
            Y2              =   960
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00FF0000&
            X1              =   3000
            X2              =   3480
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label9 
            Caption         =   "Processing Time"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3960
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "No. Of Parts"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2060
            TabIndex        =   36
            Top             =   980
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Enter/Select Size of Each Splitted File"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Files"
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Open the files to be splitted"
         Top             =   720
         Width           =   1815
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   5
         Top             =   2160
         Width           =   3360
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   4
         Top             =   2160
         Width           =   1920
      End
      Begin VB.ListBox listFile 
         BackColor       =   &H00FFFFC0&
         Height          =   1230
         Left            =   1440
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         ToolTipText     =   "Display the file to be splitted .Select the file here."
         Top             =   720
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "Remove Selected"
         Height          =   255
         Index           =   0
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Remove the selected file"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "Remove All"
         Height          =   255
         Index           =   1
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Remove all files from list"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FF0000&
         X1              =   -73800
         X2              =   -73440
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   -73680
         X2              =   -71760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   -73680
         X2              =   -73680
         Y1              =   1560
         Y2              =   1920
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   -73800
         X2              =   -73440
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         X1              =   1200
         X2              =   1440
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   1320
         X2              =   3360
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   1320
         X2              =   1320
         Y1              =   2280
         Y2              =   2640
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   1200
         X2              =   1440
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -74880
         X2              =   -68280
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   2
         X1              =   -74880
         X2              =   -68280
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   6720
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label10 
         Caption         =   "Processing Time"
         Height          =   255
         Left            =   -70320
         TabIndex        =   38
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Source File To be Merged"
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
         Height          =   495
         Left            =   -74880
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Remarks About File"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -74880
         TabIndex        =   25
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Left            =   -74880
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Remarks About File"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Source File"
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
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   6720
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         X1              =   -74880
         X2              =   -68280
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0080FF80&
         BorderWidth     =   3
         Index           =   1
         X1              =   -74880
         X2              =   -68280
         Y1              =   480
         Y2              =   480
      End
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   375
      Left            =   0
      TabIndex        =   43
      ToolTipText     =   "Show the progress of Processing."
      Top             =   6360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   1.00000e5
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   44
      Top             =   400
      Width           =   6855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "VIN Split and Merge V1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   0
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   39
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmSplit"
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
'  Semester VI
'  U.I.T.
'  Rajeev Gandhi Prodyogiki Vishwavidyalaya Bhopal.
'  Address: S-2 ShreeMaya Apartment Sector-B/363
'           Sarvdharm Colony Bhopal-42 (India)
'  Email: vinodkotiya@yahoo.co.in
'  Web : http://vinodkotiya.tripod.com
'  cell: +91-9827394994
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Date of Starting:Sunday ,23 March 2003, 3.00 AM Morning
'  Completion Date :Sunday,23 March 2003,7:00 AM
'  Associated Projects: Nil
'
'  First Modification : 24-Mar-2003
'                       Multiple Splitting Option Enabled
'  Second Modification :25-Mar-2003
'                       Settings option added.
'  Third Modification : 26-Mar-2003
'                       Batch file generation bug was fixed
'  Fifth Modification : Date not known
'                       Help support added.
'  Forth Modification : Date not known
'                       After May 2003 splash screen is modified
'                       for supporting winNT/XP transparency feature
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Dim eachFileSizeBe As Double      'inbytes
Dim infile As String
Dim outfile As String
Dim TheFile As String   'pure name of file w/o path
Dim extention As Long
Dim kulparts As Long  'store total no of parts
Dim delFiles As Boolean  'source file deleted when true
Dim delFilessplit As Boolean 'split files deleted wen merged
Dim isStop As Boolean    'processing stop when true
Dim batch As Boolean  'create batch file when true
Dim prompt As Boolean 'prompt for removable media whentrue
 Dim destfolder As String
 
 'faltoo bouncing ball variables
 Dim BallX As Integer
Dim BallY As Integer
Dim BallDirx As Integer
Dim BallDiry As Integer
'help










Private Sub chkBatch_Click()
If chkBatch.Value = 1 Then
 MsgBox "This will create a batch file for merging " _
 & Chr(13) & " Now you can merge the files on any system (on which the VIN Split and Merge v1.0 is not installed)" & Chr(13) _
 & " by running the batch file.Merging by this method is very fast."
 batch = True
Else
 MsgBox "This will not create a batch file for merging " _
 & Chr(13) & " so you can not merge the files on any system (on which the VIN Split and Merge v1.0 is not installed)" & Chr(13) _
 & " To merge the files you must install this software on other system.It is recomended that do not turn off this feature.  "
 batch = False
End If

End Sub

Private Sub chkDel_Click()
If chkDel.Value = 1 Then
 MsgBox "Now the source file will be deleted after the splitting " & Chr(13) & _
  "If its file attributes are uncheked."
 delFiles = True
Else
 delFiles = False
End If
 
End Sub

Private Sub chkDelmerge_Click()
If chkDelmerge.Value = 1 Then
 MsgBox "Now the splitted files will be deleted after the merging "
 delFilessplit = True
Else
 delFilessplit = False
End If
End Sub

Private Sub chkMedia_Click()
'chkMedia.Value = Not chkMedia.Value
If chkMedia.Value = 1 Then
 MsgBox "Please Insert A Rmovable Storage Media (eg. Floppy Disk) " & Chr(13) & _
 "And Select the Media Type (3.5 Floppy Disk) from Size of Each Splitted File combo" & Chr(13) & _
 "Set the Destination Folder to A:\"
 prompt = True
Else
 prompt = False
 End If
End Sub

Private Sub cmbSize_Click()
Dim isSourcefileSelected As Boolean
isSourcefileSelected = False
Dim i As Integer
For i = 0 To listFile.ListCount - 1
   If listFile.Selected(i) = True Then
      isSourcefileSelected = True
       Exit For
    End If
Next
If isSourcefileSelected = False Then
 Command2.SetFocus
   MsgBox "First Select the file to be splitted" & Chr(13) & _
   "To select any file click it in source file box .Use Ctrl + Mouse for multiple selection."
   Exit Sub 'no file selected
End If

If cmbSize.ListIndex = 0 Then
 eachFileSizeBe = 1457664 - 1200 'subtracting FOR *.*.VIN & BATCH file
 optByte.Value = True
ElseIf cmbSize.ListIndex = 1 Then
 eachFileSizeBe = 100431872 - 1200
 optByte.Value = True
ElseIf cmbSize.ListIndex = 2 Then
 eachFileSizeBe = 1073741824 - 1200
 optByte.Value = True
ElseIf cmbSize.ListIndex = 3 Then
 eachFileSizeBe = 4096          'floppy disk cluster size = 4kb
 optByte.Value = True
ElseIf cmbSize.ListIndex = 4 Then
 eachFileSizeBe = 2048          '2kb   NTFS
 optByte.Value = True

ElseIf cmbSize.ListIndex = 5 Or cmbSize.ListIndex = 9 Then
 eachFileSizeBe = 4096          '4kb   fat32     / fat
 optByte.Value = True
 ElseIf cmbSize.ListIndex = 6 Or cmbSize.ListIndex = 10 Then
 eachFileSizeBe = 8192     '8kb  fat32  /fat
 optByte.Value = True
ElseIf cmbSize.ListIndex = 7 Or cmbSize.ListIndex = 11 Then
 eachFileSizeBe = 16384    '16 kb   fat32/ fat
  optByte.Value = True
ElseIf cmbSize.ListIndex = 8 Or cmbSize.ListIndex = 12 Then
 eachFileSizeBe = 32768  '32 kb      fat32/fat
 optByte.Value = True


Else
  If Trim(cmbSize.Text) = " " Then
   Exit Sub
  End If
 eachFileSizeBe = Val(cmbSize.Text)      'byte
 
 If optKb.Value = True Then
  eachFileSizeBe = eachFileSizeBe * 1024   'kb
 ElseIf optMb.Value = True Then
  eachFileSizeBe = eachFileSizeBe * 1024 * 1024  'mb
 End If
End If
 
'MsgBox cmbSize.ListIndex & eachFileSizeBe
Dim iner As Integer
    iner = FreeFile
    Open listFile.List(listFile.ListIndex) For Binary Access Read As #iner
Dim flen As Long
    
   flen = LOF(iner)
   txtRemarks.Text = "File ''" & listFile.List(listFile.ListIndex) & "'' Was splitted on " & Now & " Size of File is " & flen & " Bytes"
   Frame1.Caption = "Splitting Options : Selected file size = " & flen & " Bytes"
   If eachFileSizeBe < 1 Then Exit Sub   'division by zero
   txtParts.Text = Int(flen / eachFileSizeBe) + 1
   Close #iner
   
 '  MsgBox flen
End Sub

Private Sub cmbSize_KeyUp(KeyCode As Integer, Shift As Integer)
If IsNumeric(Chr(KeyCode)) = True Or KeyCode = 32 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 37 Or KeyCode = 39 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 86 Or KeyCode = 17 Or KeyCode = 46 Or KeyCode = 35 Then
  cmbSize_Click
  Exit Sub
End If

If IsNumeric(Chr(KeyCode)) = False Then

 cmbSize.Text = "" 'Left(txtParts.Text, Len(txtParts.Text) - 1)
End If

End Sub

Private Sub cmdMerge_Click()
'here initially outfile will be *.*.vin after calling openvin then it become *.*
'and infile are the splitted file
If Trim(TheFile) = "" Then Exit Sub

 isStop = False
    destfolder = Dir2.Path 'CommonDialog1.FileName
    If Len(destfolder) < 4 Then    'If "F:\" = 3
      destfolder = Left(destfolder, 2)  'if "F:\" it return "F:"
    End If

   cmdStopmerge.Enabled = True
   cmdMerge.Enabled = False
   
   If delFilessplit = True Then   'when delete splitted file after merge is true
    Dim pos As Integer
    Deletefile (outfile) 'delete the file *.*.vin
    pos = InStrRev(outfile, ".", -1, vbBinaryCompare)
    Deletefile (Left(outfile, pos - 1) & ".bat") 'extract file name with path and add .bat then delete it
   End If
   outfile = destfolder & "\" & TheFile
   Dim infileNumber As Integer, outfileNumber As Integer
    
    outfileNumber = FreeFile
    Open outfile For Binary Access Write As #outfileNumber
    'Dim inChar As Byte, outChar As Byte
    Dim startTime As Date
    startTime = Now
    For extention = 1 To kulparts
     DoEvents
     Progress.Value = Round((Progress.Max / kulparts) * extention)
     'Me.Caption = "Merging the File ''" & TheFile & "'' Processed (" & extention & "/" & kulparts & ")"
     'Label11(1).Caption = "Merging the File ''" & TheFile & "'' Processed (" & extention & "/" & kulparts & ")"
     infile = destfolder & "\" & TheFile & ".vin" & extention  'curently having extention file.ext.2
    ' MsgBox infile
     infileNumber = FreeFile
     Open infile For Binary Access Read As #infileNumber
        Put #outfileNumber, , Input$(LOF(infileNumber), infileNumber) 'outChar
         DoEvents
        Me.Caption = "Merging " & TheFile & " (" & extention & "/" & kulparts & ")"
        Label11(1).Caption = Me.Caption
        DoEvents
     Close infileNumber
        If isStop = True Then  'if stop pressed close merged file and exit
          Close outfileNumber
          Me.Caption = "VIN Split & Merge v1.0"
          Label11(1).Caption = ""
          Progress.Value = 0
          MsgBox "Splitting Process is Aborted"
          cmdStopmerge.Enabled = False
          cmdMerge.Enabled = True

          Exit Sub
        End If
       If delFilessplit = True Then    'now delete the splitted file
        Deletefile (infile)
       End If
   Next
   Close outfileNumber      'file merged
   Me.Caption = "VIN Split & Merge v1.0"
   Label11(1).Caption = ""
   Progress.Value = 0
   Text1.Text = DateDiff("s", startTime, Now) & " seconds"
   MsgBox "File " & TheFile & " is Merged from " & kulparts & Chr(13) & _
   "Parts And saved in Destination folder " & destfolder
   cmdStopmerge.Enabled = False
   cmdMerge.Enabled = True
End Sub

Private Sub cmdSplit_Click()
'On Error GoTo chupchap

cmbSize_Click
If listFile.ListCount < 1 Then Exit Sub   'no files open so exit
If Trim(cmbSize.Text) = "" Then        'size not given exit
  MsgBox "Please Specify the size of each splitted part"
  Exit Sub
End If
isStop = False
Dim pos As Long
Dim i As Integer
'infile = listFile.ListIndex
For i = 0 To i = listFile.ListCount - 1
 If listFile.Selected(i) = True Then
  Exit For
 End If
Next
cmdSplit.Enabled = False
cmdStopsplit.Enabled = True
'if multiple file selected execute this loop
 For i = 0 To listFile.ListCount - 1
  If listFile.Selected(i) = True Then    'only choose the selected file
     resetGlobalVar
     infile = listFile.List(i)      'name of source file with path
       '    MsgBox infile
     pos = Len(infile) - InStrRev(infile, "\", -1, vbBinaryCompare)
     TheFile = Right(infile, pos)    'extract only file name without path
     
     destfolder = Dir1.Path 'folder where splitted files be stored
    If Len(destfolder) < 4 Then    'If "F:\" = 3
      destfolder = Left(destfolder, 2)  'if "F:\" it return "F:"
    End If
    
     '  If outfile = "" Then Exit Sub   ' If infile = outfile Then
     '    MsgBox "This application will not overwrite the source file." & vbCrLf & _
                "Please select another output file name"
     '        GoTo GetOutFileName
      'End If
    Dim infileNumber As Integer, outfileNumber As Integer
    infileNumber = FreeFile
    Open infile For Binary Access Read As #infileNumber
         'Dim inChar As Byte, outChar As Byte
    Dim startTime As Date
    Dim flen As Long     'store length of file
    flen = LOF(infileNumber)
    If listFile.ListCount > 1 Then txtRemarks.Text = "File ''" & infile & "'' Was splitted on " & Now & " Size: " & flen & " Bytes"    'only when more than 1 file is selected
    kulparts = Int(flen / eachFileSizeBe) + 1
 '   MsgBox "Total no of parts will be " & kulparts
     startTime = Now
     Dim fnum As Integer
      fnum = FreeFile
      If batch = True Then      'create batch file
        Open destfolder & "\" & TheFile & ".bat" For Output As #fnum
          Print #fnum, "@echo off"
          Print #fnum, "copy /b " & Chr(34) & TheFile & ".vin" & extention & Chr(34) & " " & Chr(34) & TheFile & Chr(34) ' copy /b "*.*.vin1" "*.*"
      End If
    While Not EOF(infileNumber)
        'Get #infileNumber, , inChar
        'outChar = inChar
        DoEvents
        If (prompt = True) Then MsgBox "Insert a Blank Storage Media"
        Me.Caption = "Splitting ''" & TheFile & "'' Processing (" & extention & "/" & kulparts & ")"
        Label11(1).Caption = "Splitting ''" & TheFile & "'' Processed (" & extention & "/" & kulparts & ")"
        Progress.Value = Round((Progress.Max / kulparts) * extention)
        outfile = destfolder & "\" & TheFile & ".vin" & extention
      DoEvents
        extention = extention + 1    'increament the *.*.vinextention for splitted file
        If batch = True And (kulparts > extention Or kulparts = extention) Then
         Print #fnum, "copy /b "; Chr(34) & TheFile & Chr(34) & "+" & Chr(34) & TheFile & ".vin" & extention & Chr(34) 'copy /b "*.*"+"*.*.vin2"
        End If
        outfileNumber = FreeFile
        
        Open outfile For Binary Access Write As #outfileNumber
        Put #outfileNumber, , Input$(eachFileSizeBe, infileNumber) 'outChar'put all the datasize given by user on each file
        DoEvents
        Close outfileNumber
        If isStop = True Then  'if stop pressed close input file and exit
          Close infileNumber
          
          Dim tempa As Long
          tempa = createVin(destfolder) 'create extra file for comments and noof parts
          Me.Caption = "VIN Split & Merge v1.0"
          Label11(1).Caption = ""
          Progress.Value = 0
          MsgBox "Splitting Process is Aborted"
           If batch = True Then    'close batch file
            Print #fnum, "echo file" & TheFile & " successfully Merged"
            Print #fnum, "echo This batch file is created by VIN Split & Merge v1.0"
            Print #fnum, "echo Programmer vinod kotiya"
            Close fnum
           End If
          cmdSplit.Enabled = True
          cmdStopsplit.Enabled = False

          Exit Sub
        End If
        DoEvents
    Wend
    Close infileNumber
    DoEvents
    If batch = True Then    'close batch file
      Print #fnum, "echo file" & TheFile & " successfully Merged"
      Print #fnum, "echo This batch file is created by VIN Split & Merge v1.0"
      Print #fnum, "echo Programmer vinod kotiya"
      Close fnum
    End If
    DoEvents
    txtTime.Text = DateDiff("s", startTime, Now) & " seconds"
    Dim temp As Long
    temp = createVin(destfolder) 'create extra file for comments and noof parts
    If delFiles = True Then    'now delete the source file
      Deletefile (infile)
    End If
  '  MsgBox "file processed in " & DateDiff("s", startTime, Now) & " seconds"
   DoEvents
   MsgBox "File " & TheFile & " is splitted in to " & kulparts & Chr(13) & _
    "Parts And saved in Destination folder " & destfolder
   cmdSplit.Enabled = True
    cmdStopsplit.Enabled = False
  
 End If
Next    'end of biggest loop when multiple file selection

 Me.Caption = "VIN Split & Merge v1.0"
 Label11(1).Caption = " "
 Progress.Value = 0
 Exit Sub
 
chupchap:
 'Me.Caption = "Some error occured"
End Sub
Function createVin(outputfolder As String)
Dim fnum As Integer
fnum = FreeFile
Open outputfolder & "\" & TheFile & ".VIN" For Output As #fnum
Print #fnum, TheFile
Print #fnum, kulparts
Print #fnum, txtRemarks.Text & " Processing Time" & txtTime.Text & " Splitted by ''VIN Split & Merge'' - programmer VINOD KOTIYA"
Close #fnum

End Function






Private Sub cmdSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSplit.BackColor = &HE0E0E0
End Sub

Private Sub cmdStopsplit_Click()
isStop = True
End Sub



Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HE0E0E0
End Sub

Private Sub Command2_Click()
Dim pathname As String
Dim pos As Long
   CommonDialog1.Flags = cdlOFNAllowMultiselect 'Or cdlOFNExplorer  'Or cdlOFNLongNames
   CommonDialog1.Filter = "All Files|*.*"
    CommonDialog1.ShowOpen
    infile = CommonDialog1.FileName
    If Len(infile) = 0 Then
        MsgBox "No files selected"
        Exit Sub
    End If
' Extract path name:
' IF FILETITLE IS NOT EMPTY, THEN A SINGLE FILE
' HAS BEEN SELECTED. DISPLAY IT AND EXIT
    If CommonDialog1.FileTitle <> "" Then
        listFile.AddItem CommonDialog1.FileName
         txtRemarks.Text = "File ''" & infile & "'' Was splitted on " & Now
        Exit Sub
    End If
  
' FILETITLE IS NOT EMPTY, THEN MANY FILES WERE SELECTED
' AND WE MUST EXTRACT THEM FROM THE FILENAME PROPERTY
    pos = InStr(infile, " ")
    pathname = Left(infile, pos - 1)
    infile = Mid(infile, pos + 1)
' then extract each space delimited file name
    If Len(infile) = 0 Then
        listFile.AddItem "No files selected"
        Exit Sub
    Else
        pos = InStr(infile, " ")
        While pos > 0
            listFile.AddItem pathname & Left(infile, pos - 1)
            infile = Mid(infile, pos + 1)
            pos = InStr(infile, " ")
        Wend
        listFile.AddItem pathname & (infile)
' Add the last file's name to the list
' (the last file name isn't followed by a space)
        
    End If

'    txtRemarks.Text =
    
    
    
    'extract TheFile name w/o path
    
    
   ' If infile = "" Then Exit Sub
    'infile = listFile.ListIndex
    
   ' pos = Len(infile) - InStrRev(infile, "\", -1, vbBinaryCompare)
   ' TheFile = Right(infile, pos)
'GetOutFileName:
    
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HE0E0E0
End Sub



Private Sub Command3_Click(Index As Integer)
 Dim i As Integer
If Index = 0 Then

 For i = listFile.ListCount - 1 To 0 Step -1
  If listFile.Selected(i) = True Then
  listFile.RemoveItem (i)
  End If
 Next
Else

For i = listFile.ListCount - 1 To 0 Step -1
 listFile.RemoveItem (i)
Next

End If
End Sub





Private Sub cmdStopmerge_Click()
Dim reply As VbMsgBoxResult
If delFilessplit = True Then
     reply = MsgBox("If you abort the process then you will not get back the splitted file which have been deleted so far" & Chr(13) & _
          " Do you really want to stop the merging process", vbYesNo)
Else
    reply = MsgBox(" Do you really want to stop the merging process", vbYesNo)
End If
If reply = vbYes Then isStop = True

End Sub





Private Sub Command6_Click(Index As Integer)
If Index = 0 Then
 Dim temp As Long
 On Error GoTo Exeerror
 temp = Shell(App.Path & "\credit.exe", vbNormalFocus)
 Exit Sub
Exeerror:
  MsgBox "Application 'CREDIT.EXE' is not found in its " _
  & "Default directory CREDIT.exe "
ElseIf Index = 1 Then
 Dim b As Boolean
b = ShowHelp("splithelp.chm", True)
ElseIf Index = 2 Then
MsgBox "                     VINSOFT" & vbCrLf & _
                "               *************" & vbCrLf & _
                "Programmer : VINOD KOTIYA " & vbCrLf & _
" date created: 4/03/2003 " & vbCrLf & _
 " time: 3.00 AM to 8:00 AM " & vbCrLf & _
" web : http:\\vinodkotiya.tripod.com " & vbCrLf & _
"email : vinodkotiya24@rediffmail.com "

End If

End Sub

Private Sub Command6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6(Index).BackColor = &HE0E0E0
End Sub








Private Sub Dir1_Change()
 ChDir Dir1.Path
    'File1.Path = Dir1.Path
    'File1.Pattern = Combo1.Text
End Sub

Private Sub Dir2_Change()
 ChDir Dir2.Path
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

Private Sub Drive2_Change()
On Error GoTo vinerror
ChDrive Dir2.Path
    Dir2.Path = Drive2.Drive
    Dir2.Refresh
    Exit Sub
vinerror:
  MsgBox "There is no disk in drive"
End Sub

Private Sub Form_Load()
'init globals
     'faltoo
 
     BallDiry = -1
     BallDirx = -1
     

Drive1.Drive = "C:\"
Dir1.Path = "C:\"
eachFileSizeBe = 1457664
extention = 1 '.vin will stored for file containing filenames
delFiles = False
delFilessplit = False
batch = True
isStop = False
prompt = False
cmbSize.AddItem "1457664(3.5 Floppy Disk)"
cmbSize.AddItem "100431872(100 MB Zip Disk)"
cmbSize.AddItem "1073741824(1 GB Jaz Disk)"
cmbSize.AddItem "Cluster Size = 4KB for 3.5 Floppy"
cmbSize.AddItem "Cluster Size(NTFS)"
cmbSize.AddItem "Cluster Size(FAT 32 - under 8 GB)"
cmbSize.AddItem "Cluster Size(FAT 32 - 8 to 16 GB)"
cmbSize.AddItem "Cluster Size(FAT 32 - 16 to 32 GB)"
cmbSize.AddItem "Cluster Size(FAT 32 - more than 32 GB)"
cmbSize.AddItem "Cluster Size(FAT - under 256 MB)"
cmbSize.AddItem "Cluster Size(FAT - 257 to 512 MB)"
cmbSize.AddItem "Cluster Size(FAT - 513 to 1024 MB)"
cmbSize.AddItem "Cluster Size(FAT - 1025 to 2048 MB)"
End Sub
Private Sub openVin()
Dim fnum As Integer
Dim temp As String
fnum = FreeFile
Open outfile For Input As #fnum  'here output folder will work as input folder
Line Input #fnum, TheFile
Line Input #fnum, temp
kulparts = CLng(Trim(temp))
Line Input #fnum, temp
txtRem.Text = temp
Close #fnum

'MsgBox kulparts & TheFile
End Sub
Private Sub Command1_Click()
   CommonDialog1.Flags = cdlOFNExplorer 'CommonDialog1.Flags
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "vin Files(*.vin)|*.vin"
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then
      Text2.Text = "No file is selected for merging"
      Exit Sub
   End If
   outfile = CommonDialog1.FileName '*.*.vin
   Text2.Text = outfile
    'extract TheFolder name w/o slash
    'Dim pos As Long
'    pos = InStrRev(outfile, "\", -1, vbBinaryCompare)
 '   destfolder = Left(outfile, pos - 1)
  '  MsgBox destfolder
    resetGlobalVar
    openVin

    'If infile = "" Then Exit Sub
End Sub
Private Sub resetGlobalVar()
'only used for cmdsplit
'eachFileSizeBe = 100000
isStop = False
extention = 1
Progress.Value = 0
End Sub





Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSplit.BackColor = vbWhite
End Sub

Private Sub listFile_Click()
'MsgBox listFile.List(listFile.ListIndex)
cmdSplit.Enabled = True
cmdStopsplit.Enabled = False
cmbSize_Click
End Sub

Private Sub optByte_Click()
  cmbSize_Click
End Sub

Private Sub optKb_Click()
  cmbSize_Click
End Sub

Private Sub optMb_Click()
  cmbSize_Click
End Sub

Private Sub shpBall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BallDirx = BallDirx * -1
shpBall1.FillColor = vbRed
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'MsgBox PreviousTab & SSTab1.Tab
If SSTab1.Tab = 2 Then
 timBall.Interval = 70
Else
 timBall.Interval = 0
 End If
End Sub




Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbWhite
Command6(0).BackColor = vbWhite
Command6(1).BackColor = vbWhite
Command6(2).BackColor = vbWhite
Command2.BackColor = vbWhite

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dir2.SetFocus
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dir2.SetFocus
End Sub

Private Sub timBall_Timer()
shpBall1.FillColor = vbBlue
BallX = BallX + BallDirx * frmBball.ScaleWidth / 50
If BallX < 0 Then
  Beep
  BallX = 0
  BallDirx = 1
ElseIf BallX > frmBball.ScaleWidth - shpBall.Width Then
  Beep
  BallX = frmBball.ScaleWidth - shpBall.Width
  BallDirx = -1
End If
shpBall.Left = BallX
BallY = BallY + BallDiry * (frmBball.ScaleHeight - 615) / 50    '615 is height of text bar
If BallY < 0 Then
  Beep
  BallY = 0
  BallDiry = 1
ElseIf BallY > (frmBball.ScaleHeight) - shpBall.Height Then
  Beep
  BallY = (frmBball.ScaleHeight) - shpBall.Height
  BallDiry = -1
End If
shpBall.Top = BallY
End Sub

Private Sub Timer1_Timer()
MsgBox cmbSize.ListIndex
End Sub

Private Sub txtParts_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Or KeyCode = 8 Or KeyCode = 13 Or KeyCode = 37 Or KeyCode = 39 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 86 Or KeyCode = 17 Or KeyCode = 46 Or KeyCode = 35 Then
Exit Sub
End If

If IsNumeric(Chr(KeyCode)) = False Then
 'Exit Sub
 txtParts.Text = "" 'Left(txtParts.Text, Len(txtParts.Text) - 1)
 'VScroll1_Change
'VScroll1 = -30 '-1 * Val(txtParts.Text)
End If

End Sub


Private Sub Deletefile(filenm As String)
On Error GoTo vinerror
Dim fsys As New FileSystemObject
fsys.Deletefile filenm, True
 Exit Sub
vinerror:
  MsgBox "Unable to delete the file"
End Sub

Private Sub txtRem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dir2.SetFocus
End Sub
