VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "VISUAL BASIC RUNTIME by VINOD KOTIYA"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2520
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   855
      Left            =   4680
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3360
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   2355
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":030A
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "The Controls you can watch are installed and registered successfully in your system"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   960
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   1080
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "VISUAL BASIC RUNTIME by VINOD KOTIYA"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   5295
   End
   Begin VB.OLE OLE1 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

