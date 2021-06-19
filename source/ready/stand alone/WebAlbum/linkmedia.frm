VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmAlbum 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIN WebAlbum Maker"
   ClientHeight    =   6105
   ClientLeft      =   3555
   ClientTop       =   2100
   ClientWidth     =   6855
   ForeColor       =   &H0000FF00&
   Icon            =   "linkmedia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6855
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   16711680
      TabCaption(0)   =   "&Make"
      TabPicture(0)   =   "linkmedia.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Settings"
      TabPicture(1)   =   "linkmedia.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Settings"
         Height          =   3975
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   6615
         Begin VB.Frame Frame2 
            Caption         =   "Create Autorun"
            Height          =   375
            Index           =   5
            Left            =   3480
            TabIndex        =   60
            Top             =   3510
            Width           =   3015
            Begin VB.OptionButton optAuto 
               Caption         =   "No"
               Height          =   195
               Index           =   1
               Left            =   2040
               TabIndex        =   62
               Top             =   120
               Width           =   855
            End
            Begin VB.OptionButton optAuto 
               Caption         =   "Yes"
               Height          =   195
               Index           =   0
               Left            =   1320
               TabIndex        =   61
               Top             =   120
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame frBack 
            BackColor       =   &H008B3144&
            Caption         =   "Background color of webpage"
            ForeColor       =   &H00FFFF80&
            Height          =   735
            Left            =   120
            TabIndex        =   53
            ToolTipText     =   "vin"
            Top             =   240
            Width           =   3255
            Begin VB.Shape Shape1 
               BorderWidth     =   2
               FillColor       =   &H0000FF00&
               FillStyle       =   0  'Solid
               Height          =   150
               Left            =   120
               Shape           =   3  'Circle
               Top             =   240
               Width           =   150
            End
            Begin VB.Image imgPreviewSet 
               Height          =   225
               Left            =   120
               Top             =   240
               Width           =   240
            End
            Begin VB.Label frText 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "COLOR of text"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   54
               ToolTipText     =   "vino"
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Resolution of each picture on webpage"
            Height          =   735
            Index           =   0
            Left            =   3480
            TabIndex        =   48
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
            Begin VB.TextBox txtRes 
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   0
               Left            =   600
               MaxLength       =   3
               TabIndex        =   50
               TabStop         =   0   'False
               Text            =   "200"
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtRes 
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   1
               Left            =   2160
               MaxLength       =   3
               TabIndex        =   49
               TabStop         =   0   'False
               Text            =   "180"
               Top             =   360
               Width           =   375
            End
            Begin VB.VScrollBar vscrollRes 
               Height          =   255
               Index           =   0
               LargeChange     =   8
               Left            =   960
               Max             =   999
               Min             =   10
               TabIndex        =   7
               Top             =   360
               Value           =   200
               Width           =   255
            End
            Begin VB.VScrollBar vscrollRes 
               Height          =   255
               Index           =   1
               LargeChange     =   8
               Left            =   2520
               Max             =   999
               Min             =   10
               TabIndex        =   8
               Top             =   360
               Value           =   180
               Width           =   255
            End
            Begin VB.Label frText 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Width"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   52
               Top             =   360
               Width           =   735
            End
            Begin VB.Label frText 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Height"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   1440
               TabIndex        =   51
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "No. of media in a Row"
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   3255
            Begin VB.OptionButton optNo 
               Caption         =   "or No of columns"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   47
               Top             =   360
               Width           =   1575
            End
            Begin VB.VScrollBar vscrollRes 
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               LargeChange     =   8
               Left            =   2040
               Max             =   20
               Min             =   1
               TabIndex        =   9
               Top             =   360
               Value           =   3
               Width           =   255
            End
            Begin VB.TextBox txtRes 
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   1680
               MaxLength       =   2
               TabIndex        =   46
               TabStop         =   0   'False
               Text            =   "3"
               Top             =   360
               Width           =   375
            End
            Begin VB.OptionButton optNo 
               Caption         =   "Default"
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   45
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Alternative Text when link not shown"
            Height          =   735
            Left            =   3480
            TabIndex        =   41
            Top             =   1080
            Width           =   3015
            Begin VB.OptionButton optAlt 
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Width           =   255
            End
            Begin VB.TextBox txtAlt 
               Enabled         =   0   'False
               Height          =   285
               Left            =   360
               TabIndex        =   42
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton optAlt 
               Caption         =   "Name of media"
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   10
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.VScrollBar vscrollRes 
            Height          =   255
            Index           =   2
            LargeChange     =   8
            Left            =   2280
            Max             =   999
            Min             =   10
            TabIndex        =   11
            Top             =   1920
            Value           =   30
            Width           =   255
         End
         Begin VB.Frame Frame2 
            Caption         =   "Also Display the following media information below the each link"
            Height          =   735
            Index           =   2
            Left            =   0
            TabIndex        =   39
            Top             =   2760
            Width           =   6615
            Begin VB.CheckBox chkInfo 
               Caption         =   "Type"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "Size"
               Height          =   195
               Index           =   1
               Left            =   1560
               TabIndex        =   15
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "Path"
               Height          =   195
               Index           =   2
               Left            =   5160
               TabIndex        =   21
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "Parent Folder"
               Height          =   195
               Index           =   3
               Left            =   5160
               TabIndex        =   17
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "Date Created"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   18
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "Date Last Accessed"
               Height          =   195
               Index           =   5
               Left            =   1560
               TabIndex        =   19
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "Date Last Modified"
               Height          =   195
               Index           =   6
               Left            =   3360
               TabIndex        =   20
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "Dos Name (8.3)"
               Height          =   195
               Index           =   7
               Left            =   3360
               TabIndex        =   16
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Include following information in each webpage"
            Height          =   615
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   2160
            Width           =   6375
            Begin VB.TextBox txtInfo 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   37
               Text            =   "Enter you remarks here.."
               Top             =   240
               Width           =   1815
            End
            Begin VB.OptionButton optInfo 
               Caption         =   "Top"
               Height          =   255
               Index           =   0
               Left            =   3000
               TabIndex        =   36
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optInfo 
               Caption         =   "Both"
               Height          =   255
               Index           =   2
               Left            =   4440
               TabIndex        =   34
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optInfo 
               Caption         =   "Don't "
               Height          =   255
               Index           =   3
               Left            =   5520
               TabIndex        =   13
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.TextBox txtRes 
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   33
               TabStop         =   0   'False
               Text            =   "H3"
               Top             =   240
               Width           =   375
            End
            Begin VB.VScrollBar vscrollRes 
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   2280
               Max             =   7
               TabIndex        =   32
               Top             =   240
               Value           =   3
               Width           =   255
            End
            Begin VB.OptionButton optInfo 
               Caption         =   "Bottom"
               Height          =   255
               Index           =   1
               Left            =   3600
               TabIndex        =   35
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "At"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   38
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Mode "
            Height          =   375
            Index           =   4
            Left            =   3480
            TabIndex        =   29
            Top             =   1800
            Width           =   3015
            Begin VB.OptionButton optMode 
               Caption         =   "Turbo"
               Height          =   195
               Index           =   0
               Left            =   600
               TabIndex        =   30
               Top             =   120
               Width           =   855
            End
            Begin VB.OptionButton optMode 
               Caption         =   "Normal"
               Height          =   195
               Index           =   1
               Left            =   1800
               TabIndex        =   12
               Top             =   120
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.ComboBox cmbImgList 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   3530
            Width           =   975
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "Add.."
            Height          =   255
            Left            =   2880
            TabIndex        =   23
            Top             =   3530
            Width           =   495
         End
         Begin VB.TextBox txtRes 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   2
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "30"
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "No. of media per page"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   56
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "   List of Valid media Files"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   3600
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F8F8E4&
         BorderStyle     =   0  'None
         Caption         =   "Select Picture Folders to make WEB ALBUM"
         ForeColor       =   &H00000000&
         Height          =   4335
         Left            =   0
         TabIndex        =   24
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdOpt 
            BackColor       =   &H00FFFFFF&
            Caption         =   " #  About Me #"
            Height          =   255
            Index           =   0
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   3000
            Width           =   1455
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
            Left            =   135
            TabIndex        =   2
            Top             =   600
            Width           =   3360
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
            Top             =   270
            Width           =   3360
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Create MediaAlbum"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   960
            Width           =   2535
         End
         Begin VB.CommandButton cmdOpt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "?  Help ?"
            Height          =   255
            Index           =   1
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CommandButton cmdOpt 
            BackColor       =   &H00FFFFFF&
            Caption         =   " #  About  #"
            Height          =   255
            Index           =   2
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Timer Timer1 
            Interval        =   50
            Left            =   3600
            Top             =   1320
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
            Height          =   1980
            Left            =   4095
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.TextBox txtScanned 
            Height          =   2055
            Left            =   4080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Text            =   "linkmedia.frx":1D02
            Top             =   600
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label lblScan 
            BackStyle       =   0  'Transparent
            Caption         =   "No Picture Folder Selected .................."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   135
            TabIndex        =   27
            Top             =   3720
            Width           =   6480
         End
         Begin VB.Line lnTop 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            Index           =   1
            X1              =   135
            X2              =   3615
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line lnTop 
            BorderColor     =   &H00C0FFFF&
            BorderStyle     =   3  'Dot
            BorderWidth     =   4
            Index           =   2
            Visible         =   0   'False
            X1              =   3720
            X2              =   5280
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line lnTop 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   6
            Index           =   0
            X1              =   135
            X2              =   3615
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   8
            Index           =   0
            X1              =   135
            X2              =   6480
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line lnTop 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   6
            Index           =   3
            Visible         =   0   'False
            X1              =   3720
            X2              =   5520
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   8
            Index           =   1
            X1              =   3720
            X2              =   6480
            Y1              =   360
            Y2              =   360
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "VIN LinkMedia v1.0"
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
      TabIndex        =   57
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "VIN WebAlbum Maker v2.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   59
      Top             =   5400
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   6855
   End
   Begin VB.Menu Preview 
      Caption         =   "adrishya"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "Change Background color of webpage"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Change Forecolor of webpage"
         Index           =   1
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmAlbum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''--------------------------------------------------------------------------
'----------------------------------------------------------------------------
'-----------------------Programmer:- vinod kotiya -----------------------
'-------------------- date - 10-07-2003 ----------------------------------
'-------------------- time - 11.30 am to 5.30 pm + 7.30 pm to 00.00 am
'-------------------- total hours - 10.30 hrs. -----------------------------
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Dim InitialFolder As String
Dim totalFiles As Integer
' file attributes
Dim ThisFile As File
Dim thisfolder  As Folder
Dim Fsys As New FileSystemObject
Dim mediasinrow As Integer         ' store total no of medias in each row
Dim mediasinpage As Integer    ' store total no of medias in each page
Dim PageNo  As Integer    ' store current no of pages
Dim mediaWidth As Integer    ' store resolution  of medias
Dim mediaHeight As Integer     ' store resolution  of medias
Dim altText As String              ' store alternative text
Dim hcodeBACK As String    'store hex code of background color of page
Dim hcodeTEXT As String    'store hex code of text color of page
Dim starttime As Date
Dim Create As Boolean    ' true when making album
'''----- file linking
Dim firstFolder As String
Dim linkFile As New Collection   ' store the each filenames created and corresponding info are beloww
Dim ParentName  As New Collection    'relative to parent folder. if file is in parent then set 24051982
Dim dEpth As Integer     'count the depth . if file is inside parent then set 0

Dim scroll As String    'used scrolling text
Dim madhya As Integer     'contain splitted scroll no
Dim delay As Integer
Dim MoveOnce As Boolean 'used to change color of commands only once inframe_mouse move







Private Sub chkInfo_GotFocus(Index As Integer)
Call Frame2_MouseMove(2, 0, 0, 0, 0)
End Sub

Private Sub cmbImgList_GotFocus()
lblStatus.Caption = "These extention determines the type of media files to be added on your WebAlbum at the time of scanning folders"
End Sub

Private Sub cmdNew_Click()
Dim txt As String
txt = InputBox("Enter new media extention to be added : " & vbCrLf & "  use ' * ' then ' . ' then 'extention' " & vbCrLf & "  eg.  ", "New Extention", "*.bmp")
If Trim(txt) = "*.bmp" Or Trim(txt) = "" Then
 Exit Sub
Else
If Trim(Left(txt, 2)) <> "*." Then
  MsgBox " You must enter ' *. ' before extention eg . *.jpeg"
  Exit Sub
End If
On Error GoTo fileerror
 Dim Fsys As New Scripting.FileSystemObject
 Dim Tstream As TextStream
 Set Tstream = Fsys.OpenTextFile(App.Path & "\data\imgext.vin", ForAppending)
 Tstream.WriteLine (txt)
 Dim i As Integer
  For i = cmbImgList.ListCount - 1 To 0
  cmbImgList.RemoveItem i
 Next
  loadExt
End If
 Exit Sub
fileerror:
 MsgBox "An error occured while saving file " & vbCrLf & _
 "Probably there is no writing media "
End Sub

Private Sub loadExt()
Dim FNum As Integer
Dim currentline As String

On Error GoTo fileerror
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\imgext.vin" For Input As FNum    'dont use #1 for multiple file openings
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbImgList.AddItem currentline
    Wend
    Close #FNum
   Exit Sub
fileerror:
    MsgBox "Unkown error while opening file " & "imgext.vin" _
     & "file is effected by any fool "

End Sub
Private Sub loadSettings()
Dim FNum As Integer
Dim currentline As String

On Error GoTo fileerror
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\config.vin" For Input As FNum    'dont use #1 for multiple file openings
    'While Not EOF(FNum)
    Line Input #FNum, currentline  '<color>
       Line Input #FNum, currentline
        hcodeBACK = currentline
       Line Input #FNum, currentline
        hcodeTEXT = currentline
       Line Input #FNum, currentline
        frBack.BackColor = Val(currentline)
       Line Input #FNum, currentline
        frText(0).ForeColor = Val(currentline)
    Line Input #FNum, currentline '<resolution>
       Line Input #FNum, currentline
        vscrollRes(0).Value = Val(currentline)
       Line Input #FNum, currentline
        vscrollRes(1).Value = Val(currentline)
    Line Input #FNum, currentline '<no_of_medias>
       Line Input #FNum, currentline
        If Val(currentline) = 0 Then  'default
        optNo(1).Value = True
       Line Input #FNum, currentline 'nothing below
        Else    'not default
        optNo(0).Value = True
       Line Input #FNum, currentline 'something value below
        vscrollRes(3).Value = Val(currentline)
        End If
    Line Input #FNum, currentline '<alt>
       Line Input #FNum, currentline
        If Val(currentline) = 0 Then  'default
        optAlt(1).Value = True
       Line Input #FNum, currentline 'nothing below
        Else    'not default
        optAlt(0).Value = True
       Line Input #FNum, currentline 'something value below
        txtAlt.Text = currentline
        txtAlt.Enabled = True
        End If
   Line Input #FNum, currentline '<mediaper page>
       Line Input #FNum, currentline
        vscrollRes(2).Value = Val(currentline)
   Line Input #FNum, currentline '<mode>
       Line Input #FNum, currentline
        optMode(Val(currentline)).Value = True
   Line Input #FNum, currentline '<rem>
       Line Input #FNum, currentline
        If Val(currentline) = 3 Then  'default
        optInfo(3).Value = True
        Line Input #FNum, currentline 'skip below
        Line Input #FNum, currentline 'skip below
        Else    'not default
        optInfo(Val(currentline)).Value = True
        Line Input #FNum, currentline 'something value below
        txtInfo.Text = currentline
        Line Input #FNum, currentline
        vscrollRes(4).Value = Val(currentline)
        End If
 Line Input #FNum, currentline '<info>
       Line Input #FNum, currentline
        If currentline = "0" Then chkInfo(0).Value = Checked
       Line Input #FNum, currentline
        If currentline = "1" Then chkInfo(1).Value = Checked
       Line Input #FNum, currentline
        If currentline = "7" Then chkInfo(7).Value = Checked
       Line Input #FNum, currentline
        If currentline = "3" Then chkInfo(3).Value = Checked
       Line Input #FNum, currentline
        If currentline = "4" Then chkInfo(4).Value = Checked
       Line Input #FNum, currentline
        If currentline = "5" Then chkInfo(5).Value = Checked
       Line Input #FNum, currentline
        If currentline = "6" Then chkInfo(6).Value = Checked
       Line Input #FNum, currentline
        If currentline = "7" Then chkInfo(7).Value = Checked
  Line Input #FNum, currentline '<autorun>
       Line Input #FNum, currentline
        optAuto(Val(currentline)).Value = True
  Line Input #FNum, currentline '<credit>
  Line Input #FNum, currentline '<credit>
    'Wend
    Close #FNum
   Exit Sub
fileerror:
    MsgBox "Unkown error while opening file " & "config.vin" _
     & "file is effected by any fool "

End Sub
Private Sub SaveSettings()
Dim FNum As Integer
Dim currentline As String

On Error GoTo fileerror
   
    FNum = FreeFile    'getting file no for futures referance
    Open App.Path & "\data\config.vin" For Output As FNum    'dont use #1 for multiple file openings
    'While Not EOF(FNum)
    
    Print #FNum, "<COLOR HTML VB BACK TEXT>"
       Print #FNum, hcodeBACK
       Print #FNum, hcodeTEXT
       Print #FNum, frBack.BackColor
       Print #FNum, frText(0).ForeColor
    Print #FNum, "<RESOLUTION>"
       Print #FNum, vscrollRes(0).Value
       Print #FNum, vscrollRes(1).Value
    Print #FNum, "<NO_OF_mediaS>"
       
        If optNo(1).Value Then  'default
              Print #FNum, "0"
              Print #FNum, "" 'nothing below
        Else    'not default
        
         Print #FNum, "1" 'something value below
         Print #FNum, vscrollRes(3).Value
        End If
    Print #FNum, "<ALT>"
        If optAlt(1).Value Then  'default
          Print #FNum, "0"
          Print #FNum, " "
        Else    'not default
          Print #FNum, "1" 'something value below
          Print #FNum, txtAlt.Text
        End If
   Print #FNum, "<mediaS_PER_PAGE>"
       Print #FNum, vscrollRes(2).Value
   Print #FNum, "<MODE>"
       If optMode(0).Value = True Then
         Print #FNum, "0"
       Else
         Print #FNum, "1"
       End If
   Print #FNum, "<REM>"
       
       
        If optInfo(3).Value = True Then
        Print #FNum, "3" 'dont
        ElseIf optInfo(0).Value = True Then    'not default
        Print #FNum, "0"
        ElseIf optInfo(1).Value = True Then    'not default
        Print #FNum, "1"
        ElseIf optInfo(2).Value = True Then    'not default
        Print #FNum, "2"
        End If
        Print #FNum, txtInfo.Text
        Print #FNum, vscrollRes(4).Value
       
 Print #FNum, "<INFO>"
       If chkInfo(0).Value Then
        Print #FNum, "0"
       Else
        Print #FNum, "x"
       End If
       If chkInfo(1).Value Then
        Print #FNum, "1"
       Else
        Print #FNum, "x"
       End If
       If chkInfo(7).Value Then
        Print #FNum, "7"
       Else
        Print #FNum, "x"
       End If
       If chkInfo(3).Value Then
        Print #FNum, "3"
       Else
        Print #FNum, "x"
       End If
       If chkInfo(4).Value Then
        Print #FNum, "4"
       Else
        Print #FNum, "x"
       End If
       If chkInfo(5).Value Then
        Print #FNum, "5"
       Else
        Print #FNum, "x"
       End If
       If chkInfo(6).Value Then
        Print #FNum, "6"
       Else
        Print #FNum, "x"
       End If
       If chkInfo(2).Value Then
        Print #FNum, "2"
       Else
        Print #FNum, "x"
       End If
        
  Print #FNum, "<AUTORUN>"
       If optAuto(0).Value = True Then
         Print #FNum, "0"
       Else
         Print #FNum, "1"
       End If
  Print #FNum, "THIS FILE IS NESSECCARY FOR VIN WebAlbum Maker BY VINOD KOTIYA"
  Print #FNum, "http:\\vinodkotiya.tripod.com"
  Print #FNum, "email: vinodkotiya24@rediffmail.com"
  Print #FNum, "fone: +91-0755-2794428"
    'Wend
    Close #FNum
   Exit Sub
fileerror:
    MsgBox "An error occured while saving file " & "config.vin" & vbCrLf & _
      "Probably there is no writing media "

End Sub



Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = " If you wanna add more media file extentions then click here. " & _
"These extention determines the type of media files to be added on your WebAlbum at the time of scanning folders"
End Sub

Private Sub cmdOpt_Click(Index As Integer)
If Index = 1 Then    'help
 MsgBox "                                HELP " & vbCrLf & _
 "**************************************" & vbCrLf & vbCrLf & _
 "What's This:- This utility make the album of your" & vbCrLf & _
"              Pictures collection in to html format." & vbCrLf & _
"              You should place your all pictures in to any folder or group of folders." & vbCrLf & _
"              This utility also create AUTORUN if you want the album to be exported on CD ROM." & vbCrLf & _
"              The main.html is the page from which you can access whole album." & vbCrLf & vbCrLf & _
"How2Use:- Place your mouse cursor or use TAB to see more information in status bar.  " & vbCrLf & _
"*****" & vbCrLf
ElseIf Index = 2 Then    'about
 MsgBox "                                About " & vbCrLf & _
 "**************************************" & vbCrLf & _
 "                               VIN WebAlbum Maker " & vbCrLf & _
"                               Programmed By : - VINOD KOTIYA    " & vbCrLf & _
"                              Created On:- 10-07-2003 " & vbCrLf & _
"                              Time :- 11:30 AM to 05:30 PM + 7.30 PM to 00:00 AM  " & vbCrLf & _
"                              Total Hours :- 10.5 hr.       " & vbCrLf & _
"                              Proudly releasing version 1.0 " & vbCrLf & _
"///////////////////////////////////////////////////////////////////////////////" & vbCrLf & _
"  FIRST MODIFICATION :- Enable to save user's  settings " & vbCrLf & _
"                           Proudly releasing version 2.0" & vbCrLf & _
"//////////////////////////////////////////////////////////////////////////////"
ElseIf Index = 0 Then    'about me
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

Private Sub cmdOpt_GotFocus(Index As Integer)
If Index = 1 Then
 lblStatus.Caption = "Click to get Help ?. "
ElseIf Index = 2 Or Index = 0 Then
 lblStatus.Caption = "Click to get information about this product."
End If
End Sub

Private Sub cmdOpt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
cmdOpt_GotFocus (Index)
cmdOpt(Index).BackColor = &HF0E7D7
MoveOnce = True
End Sub

Private Sub Command1_Click()
If Create = False Then
  Create = True
  Command1.Caption = "Cancel Operation"
Else
  Create = False
  Command1.Caption = "Create WebAlbum"
  Label(1).Caption = "Process Aborted..........."
  Exit Sub
End If
Dim vinod As Boolean
    If optMode(1).Value = True Then
       lnTop(2).Visible = True
       lnTop(3).Visible = True
    End If
    starttime = Now
    vinod = checkonce
    If vinod = False Then Exit Sub
    ChDrive Drive1.Drive
    ChDir Dir1.Path
    InitialFolder = CurDir
    firstFolder = Dir1.Path
    totalFiles = 0
    Label(1).Visible = True
    lblStatus.Visible = False
    
    ScanFolders
    
    CreateMainPage
    
    If optAuto(0).Value Then CreateAutorun
        
    createData
    
   ' createAbout
    
    lblScan.Caption = "Process complete in " & DateDiff("s", starttime, Now) & " Seconds.Files scanned  " & totalFiles
    Label(1).Visible = False
    lblStatus.Visible = True
    lblStatus.Caption = Label(1).Caption
    'MsgBox "There are " & totalFiles & " under the " & InitialFolder & " folder"
    lnTop(2).Visible = False
    lnTop(3).Visible = False

End Sub
Function checkonce() As Boolean
Dim ext As String
Dim i As Integer
For i = cmbImgList.ListCount - 1 To 0 Step -1
    ext = ext & cmbImgList.List(i) & ";"
 Next
' MsgBox Left(ext, Len(ext) - 1)
File1.Pattern = Left(ext, Len(ext) - 1)  '"*.jpg ;*.bmp"
If Trim(txtRes(0).Text) = "" Or Trim(txtRes(1).Text) = "" Then
  MsgBox "Check the resolution of each medias"
  checkonce = False
ElseIf Trim(txtRes(3).Text) = "" And optNo(0).Value = True Then
MsgBox "Check the no of medias in  a Row "
    checkonce = False
 ElseIf Trim(txtAlt.Text) = "" And optAlt(0).Value = True Then
 MsgBox "Check the Alternative text "
    checkonce = False
 Else
     If optAlt(0).Value = True Then txtAlt = txtAlt.Text
     If optNo(0).Value = True Then mediasinrow = Val(txtRes(3).Text)
     mediasinpage = vscrollRes(2).Value
     checkonce = True
End If
End Function





Sub ScanFolders()

If Create = False Then Exit Sub

On Error GoTo chupchap
lnTop(0).X2 = lnTop(0).X1
lnTop(0).X2 = lnTop(0).X1
Dim subFolders As Integer
Dim i As Integer
Dim FNum As Integer
PageNo = 1      'set new page for each folder
   Set thisfolder = Fsys.GetFolder(File1.Path)          'for folder attributes
  
            'make collection
        'when start scanning every folder then put 'vin' in link file as marker and put the folder name in parentname
          linkFile.Add "vin"
          ParentName.Add thisfolder.Name
           
   'find the depth of this folder
   dEpth = 0
   i = InStr(1, Right(File1.Path, Len(File1.Path) - Len(firstFolder)), "\")
          While i > 0
                dEpth = dEpth + 1
                i = InStr(i + 1, Right(File1.Path, Len(File1.Path) - Len(firstFolder)), "\")
          Wend
Dim dot As String
dot = ""
For v = 1 To dEpth
  dot = dot & "../"
 Next
       
   Upper (dot)       'fill txtbox with starting tags
     txtScanned.Text = txtScanned.Text & " <center> <font size = 6 color = " & Chr(34) & "#00ffff" & Chr(34) & ">" & thisfolder.Name & " ( " & File1.ListCount & " )" & " </font></center><br><br><br>" & vbCrLf
       ' putting users information
       If optInfo(0).Value = True Or optInfo(2).Value = True Then txtScanned.Text = txtScanned.Text & "<H" & vscrollRes(4).Value & "> " & txtInfo.Text & "</H" & vscrollRes(4).Value & "> <br><br>"
     txtScanned.Text = txtScanned.Text & "<center><table cellspacing = 20>" & vbCrLf & "<tr>" & vbCrLf
   
For i = 0 To File1.ListCount - 1                 'scanning current folder
  'Text1.Text = Text1.Text & vbCrLf & File1.List(i)
     lnTop(0).X2 = lnTop(0).X1 + Round(6345 * ((i + 1) / File1.ListCount))      'progress bar
     lnTop(1).X2 = lnTop(1).X1 + Round(6345 * ((i + 1) / File1.ListCount))         'progress bar
     lblScan.Caption = " Scanning  folder " & thisfolder.Name & "  for Picture  " & File1.List(i)     'progress bar
     Label(1).Caption = "Picture no : " & i & "   Total Pictures Scanned : " & totalFiles + i + 1 & "   Time : " & DateDiff("s", starttime, Now) & " Seconds  "
      If optAlt(1).Value = True Then altText = File1.List(i)        'if user not give alternate text then use filename
       Set ThisFile = Fsys.GetFile(File1.Path & "\" & File1.List(i))             ' for file attributes
           
           '-----------------------making single cell for each media file
           txtScanned.Text = txtScanned.Text & "<td> <center><a href = " & Chr(34) & File1.List(i) & Chr(34) & "><br><FONT size = 4>" & File1.List(i) & "</FONT></a><font size = 2 color =" & Chr(34) & hcodeTEXT & Chr(34) & ">"
            
            If chkInfo(0).Value Then txtScanned.Text = txtScanned.Text & "<BR>Tupe " & ThisFile.Type
            If chkInfo(1).Value Then txtScanned.Text = txtScanned.Text & "<BR>SIZE " & Int(ThisFile.Size / 1024) & " KB"
            If chkInfo(7).Value Then txtScanned.Text = txtScanned.Text & "<br>DOS Name  " & ThisFile.ShortName
            If chkInfo(3).Value Then txtScanned.Text = txtScanned.Text & "<br>Parent Folder " & thisfolder.ParentFolder
            If chkInfo(4).Value Then txtScanned.Text = txtScanned.Text & "<br>Date Created " & ThisFile.DateCreated
            If chkInfo(5).Value Then txtScanned.Text = txtScanned.Text & "<br>Date Last Accessed " & ThisFile.DateLastAccessed
            If chkInfo(6).Value Then txtScanned.Text = txtScanned.Text & "<br>Date Last Modified " & ThisFile.DateLastModified
            If chkInfo(2).Value Then txtScanned.Text = txtScanned.Text & "<br>Path " & ThisFile.Path
           
            txtScanned.Text = txtScanned.Text & "  </font></center>" & vbCrLf & "   </td>"
         ''------------- 1 cell complete with media , name and information
            
        If (i + 1) Mod mediasinrow = 0 Then       'when 1 row com[plete switch to next row
              txtScanned.Text = txtScanned.Text & vbCrLf & "</tr><tr>"
       End If
       
       If (i + 1) Mod mediasinpage = 0 Then    'when 1 page complete
          txtScanned.Text = txtScanned.Text & "</tr></table></center> "
          If PageNo = 1 Then txtScanned.Text = txtScanned.Text & "<br><CENTER><FONT size = 6><A href=" & Chr(34) & Replace(thisfolder.Name, " ", "_") & PageNo + 1 & ".html" & Chr(34) & "><img src =" & Chr(34) & dot & "data/next.gif" & Chr(34) & "></img></A></font></center> "
          If PageNo > 1 Then txtScanned.Text = txtScanned.Text & "<br><CENTER><FONT size = 6><A href=" & Chr(34) & Replace(thisfolder.Name, " ", "_") & PageNo - 1 & ".html" & Chr(34) & "><img src =" & Chr(34) & dot & "data/prev.gif" & Chr(34) & "></img></A>&nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; <A href=" & Chr(34) & thisfolder.Name & PageNo + 1 & ".html" & Chr(34) & "><img src =" & Chr(34) & dot & "data/next.gif" & Chr(34) & "></img></A></font></center> <br><br>"
            If optInfo(1).Value = True Or optInfo(2).Value = True Then txtScanned.Text = txtScanned.Text & "<H" & vscrollRes(4).Value & "> " & txtInfo.Text & "</H" & vscrollRes(4).Value & "> <br><br>"
          Lower (dot)
          FNum = FreeFile
          Open File1.Path & "\" & Replace(thisfolder.Name, " ", "_") & PageNo & ".html" For Output As FNum
             Print #FNum, txtScanned.Text
          Close #FNum
          'make collection
          linkFile.Add Replace(thisfolder.Name, " ", "_") & PageNo & ".html"
           If Right(File1.Path, Len(File1.Path) - Len(firstFolder)) = "" Then
             ParentName.Add ""       '"same folder " balnk
           Else
             ParentName.Add Replace(Right(File1.Path, Len(File1.Path) - Len(firstFolder) - 1), "\", "/") & "/"    'eg pictures/the wallpaper/
           End If
                     
           Upper (dot)       'fill txtbox with starting tags
           txtScanned.Text = txtScanned.Text & " <center> <font size = 6 color = " & Chr(34) & "#00ffff" & Chr(34) & ">" & thisfolder.Name & " ( " & File1.ListCount & " )" & " </font></center><br><br><br>" & vbCrLf
               If optInfo(0).Value = True Or optInfo(2).Value = True Then txtScanned.Text = txtScanned.Text & "<H" & vscrollRes(4).Value & "> " & txtInfo.Text & "</H" & vscrollRes(4).Value & "> <br><br>"
           txtScanned.Text = txtScanned.Text & "<center><table cellspacing = 20>" & vbCrLf & "<tr>" & vbCrLf
            PageNo = PageNo + 1
         End If
       DoEvents
Next          'scanning each file completes
     
     txtScanned.Text = txtScanned.Text & "</tr></table></center><br> "
     'putting users information
     If optInfo(1).Value = True Or optInfo(2).Value = True Then txtScanned.Text = txtScanned.Text & "<H" & vscrollRes(4).Value & "> " & txtInfo.Text & "</H" & vscrollRes(4).Value & "> <br><br>"
   Lower (dot)
   FNum = FreeFile
   Open File1.Path & "\" & Replace(thisfolder.Name, " ", "_") & PageNo & ".html" For Output As FNum
    Print #FNum, txtScanned.Text
   Close #FNum
   'make collection
          linkFile.Add Replace(thisfolder.Name, " ", "_") & PageNo & ".html"
           If Right(File1.Path, Len(File1.Path) - Len(firstFolder)) = "" Then
             ParentName.Add ""       '"same folder "    blank
           Else
             ParentName.Add Replace(Right(File1.Path, Len(File1.Path) - Len(firstFolder) - 1), "\", "/") & "/"    'eg pictures/the wallpaper/
           End If
    
    GoTo Skip
chupchap:
lblStatus.Caption = "An unexpected error occured during operation ............."
 MsgBox "An unexpected error occured." & vbCrLf & "Probably there is no writing media." & vcrlf & "Chances of file handling error is higher. " & vbCrLf & _
 "Also note that u can't create WebAlbum directly from pictures CD. " & vbCrLf & "Now attempting to scan another folder ......... "

Skip:
    totalFiles = totalFiles + File1.ListCount
    subFolders = Dir1.ListCount
    If subFolders > 0 Then
        For i = 0 To subFolders - 1
            ChDir Dir1.List(i)
            Dir1.Path = Dir1.List(i)
 '             MsgBox "Now scanning folder   " & Dir1.Path
            File1.Path = Dir1.Path 'Dir1.List(i)
            frmAlbum.Refresh
            ScanFolders
        Next
    End If
    File1.Path = Dir1.Path
    MoveUp
 

End Sub
Private Sub Upper(dotDot As String)
'Dim v As Integer
txtScanned.Text = "<HTML>" & vbCrLf & "    <HEAD> " & "         <TITLE>VIN Web Album</TITLE> " & vbCrLf & _
"    </HEAD>" & vbCrLf & "  <BODY bgcolor = " & Chr(34) & hcodeBACK & Chr(34) & "  TEXT = " & Chr(34) & hcodeTEXT & Chr(34) & _
 " LINK = " & Chr(34) & hcodeTEXT & Chr(34) & " vlink = " & Chr(34) & hcodeTEXT & Chr(34) & " alink = " & Chr(34) & hcodeTEXT & Chr(34) & ">" & vbCrLf

 txtScanned.Text = txtScanned.Text & " <center><font size = 2 color = " & Chr(34) & "#ffffff" & Chr(34) & " ><A href=" & Chr(34) & "http://vinodkotiya.tripod.com" & Chr(34) & "><B>http://vinodkotiya.tripod.com</B></A>&nbsp;|&nbsp;<A  " & _
"href= " & Chr(34) & dotDot & "about.html" & Chr(34) & "<B>ABOUT / Credit </B></A>&nbsp;|&nbsp;&nbsp;|&nbsp;<A href=" & Chr(34) & "mailto:vinodkotiya24@rediffmail.com" & Chr(34) & "><B>Mail me</B></A>&nbsp;|&nbsp;<A " & _
"href= " & Chr(34) & "http://vinodkotiya.tripod.com" & Chr(34) & "><B>VINSOFT</B></A>" & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
"<A href=" & Chr(34) & dotDot & "main.html" & Chr(34) & "><img src =" & Chr(34) & dotDot & "data/home.gif" & Chr(34) & "></img></A></font></center><hr>"

End Sub
Private Sub Lower(dotDot As String)
'Dim v As Integer
  
txtScanned.Text = txtScanned.Text & "<br><hr> <center><font size = 2 color = " & Chr(34) & "#ffffff" & Chr(34) & " ><A href=" & Chr(34) & "http://vinodkotiya.tripod.com" & Chr(34) & "><img src =" & Chr(34) & dotDot & "data/vinsoft.gif" & Chr(34) & "></img></A>" & _
"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A  " & _
"href= " & Chr(34) & dotDot & "about.html" & Chr(34) & "<B>ABOUT / Credit </B></A>&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;<A href=" & Chr(34) & "mailto:vinodkotiya24@rediffmail.com" & Chr(34) & "><B>Mail me</B></A>" & _
"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
"<A href=" & Chr(34) & dotDot & "main.html" & Chr(34) & "><img src =" & Chr(34) & dotDot & "data/home.gif" & Chr(34) & "></img></A></font></center>"

txtScanned.Text = txtScanned.Text & vbCrLf & "  </BODY>" & vbCrLf & "</HTML>"
End Sub
Sub MoveUp()
    If Dir1.List(-1) <> InitialFolder Then
        ChDir Dir1.List(-2)
        Dir1.Path = Dir1.List(-2)
    End If
End Sub


Private Sub CreateMainPage()
Dim v As Integer
Upper ("")
txtScanned.Text = txtScanned.Text & "<br><br><center> <font size = 7><b><u> MediaAlbum Home Page</u></b> </font></center><br><br>" & vbCrLf
txtScanned.Text = txtScanned.Text & "<h2> Click on any of the following links ..... </h2><br><br>" & vbCrLf
txtScanned.Text = txtScanned.Text & vbCrLf
For v = 1 To linkFile.Count
   If linkFile.Item(v) = "vin" Then
     txtScanned.Text = txtScanned.Text & "<font size = 4 align = left>&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; " & ParentName.Item(v) & "</font><br>" & vbCrLf
   Else
     txtScanned.Text = txtScanned.Text & "<CENTER><a href = " & Chr(34) & ParentName.Item(v) & linkFile.Item(v) & Chr(34) & " >" & Left(linkFile.Item(v), Len(linkFile.Item(v)) - 5) & "</a></center><BR> " & vbCrLf     'remove  .html
   End If
Next

Lower ("")
 v = FreeFile
   Open InitialFolder & "\" & "main.html" For Output As v
    Print #v, txtScanned.Text
   Close #v
  
End Sub

Private Sub CreateAutorun()
Dim v As Integer
 txtScanned.Text = "[autorun]" & vbCrLf & "open = start main.html" & vbCrLf & "icon = vin.ico"
v = FreeFile
   Open InitialFolder & "\" & "autorun.inf" For Output As v
    Print #v, txtScanned.Text
   Close #v
 
End Sub
Private Sub createData()
lblScan.Caption = "Now Copying data files ..........."
Dim Fsys As New FileSystemObject
On Error GoTo vinerror
If Fsys.FolderExists(InitialFolder & "\" & "data") = False Then 'if folder not exist create it
  Fsys.CreateFolder InitialFolder & "\" & "data"
End If
'Fsys.CopyFolder App.Path & "\data", InitialFolder & "\" & "data", True
'Fsys.CopyFile "c:\cdata\*.txt", "c:\windows\vinbakup", True
Fsys.CopyFile App.Path & "\data\vin.ico", InitialFolder & "\", True        'copy icon file
Fsys.CopyFile App.Path & "\data\vinsoft.gif", InitialFolder & "\" & "data\", True       'copy gif file
Fsys.CopyFile App.Path & "\data\next.gif", InitialFolder & "\" & "data\", True       'copy gif file
Fsys.CopyFile App.Path & "\data\prev.gif", InitialFolder & "\" & "data\", True       'copy gif file
Fsys.CopyFile App.Path & "\data\home.gif", InitialFolder & "\" & "data\", True       'copy gif file
Fsys.CopyFile App.Path & "\data\make.gif", InitialFolder & "\" & "data\", True       'copy gif file
Fsys.CopyFile App.Path & "\data\set.gif", InitialFolder & "\" & "data\", True       'copy gif file
Fsys.CopyFile App.Path & "\data\vin.jpg", InitialFolder & "\" & "data\", True       'copy jpg file
Fsys.CopyFile App.Path & "\data\about.html", InitialFolder & "\", True       'copy html file
'Fsys.CopyFile App.Path & "\data\zzzz", InitialFolder & "\", True        'copy shortcut

Exit Sub
vinerror:
 MsgBox "file handling error occured."

End Sub

Private Sub Command1_GotFocus()
lblStatus.Caption = "If you have select the pictures folder then click to generate your Web Album. "
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1_GotFocus
Command1.BackColor = &HF0E7D7
MoveOnce = True
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Dir1_Change()
    ChDir Dir1.Path
    File1.Path = Dir1.Path
   lnTop(0).X2 = lnTop(0).X1
   lnTop(1).X2 = lnTop(1).X1
'MsgBox Replace(Dir1.Path, " ", "_")

End Sub

Private Sub Dir1_Click()
  lblScan.Caption = "Selected Folder is " & Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Dir1_GotFocus()
Call Dir1_MouseMove(0, 0, 0, 0)
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If x = -1 Then
 lblStatus.Caption = "Don't Try to create web album directly from CD Drive because of absance of writing media. "
 Exit Sub
End If
lblStatus.Caption = "Select the folder which contain pictures or subfolders with pictures " & _
" The main page of your webAlbum will be saved in the selected folder and the supporting " & _
"pages will be saved in corresponding sub folder.You can access all from main page "
End Sub

Private Sub Drive1_Change()
On Error GoTo vinerror
    ChDrive Dir1.Path
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    Exit Sub
vinerror:
   MsgBox "There is no disk in drive " & Drive1.Drive

End Sub


Private Sub Drive1_GotFocus()
Call Dir1_MouseMove(0, 0, -1, 0)
End Sub

Private Sub File1_Click()
Set ThisFile = Fsys.GetFile(Dir1.Path & "\" & File1.List(File1.ListIndex))
MsgBox "Created " & ThisFile.DateCreated & vbCrLf & _
" Size  " & Int(ThisFile.Size / 1024) & " KB " & vbCrLf & _
"DOS Name  " & ThisFile.ShortName & vbCrLf
Set thisfolder = Fsys.GetFolder(Dir1.Path)
MsgBox "Folder " & File1.ListCount

End Sub

Private Sub Form_Load()
'loadReg
'makeTranslusent
loadExt
    ChDrive App.Path
    ChDir App.Path
mediasinrow = 3
'mediaWidth = 200
'mediaHeight = 180
'altText As String
hcodeBACK = "#31448b"
hcodeTEXT = "#ffffff"
lnTop(0).X2 = lnTop(0).X1
lnTop(1).X2 = lnTop(1).X1
lnTop(2).X2 = lnTop(2).X1
lnTop(3).X2 = lnTop(3).X1
madhya = 1
scroll = " VIN Web Album Maker v2.0 : by Vinod Kotiya    ********************"
loadSettings
frBack.ToolTipText = hcodeBACK
frText(0).ToolTipText = hcodeTEXT
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'updateReg
SaveSettings
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = "If size of each medias are more then 50 KB or resolution above 800 X 600 then " & _
 " medias per page should be less then 40 otherwise it will take too much time to load the webpage on 128 MB RAM " & _
 "But you can use more then 40 per page if medias are of small size or small resolution."
End Sub

Private Sub Frame2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 If Index = 0 Then
   lblStatus.Caption = "Set the resolution of each medias to be displayed on webAlbum as a link. You can rich to the original medias by clicking on these links."
 ElseIf Index = 1 Then
   lblStatus.Caption = "Set the no of medias to be displayed in each row or it is the no of columns.Click on default if you want default no of columns."
  ElseIf Index = 2 Then
  lblStatus.Caption = "Determines the which information should be displayed alongwith each media."
  ElseIf Index = 3 Then
  lblStatus.Caption = "If you want to include any message in web album then type it in the text box also specify its size and position."
 ElseIf Index = 4 Then
  lblStatus.Caption = "Determines the execution speed for scanning files. Turbo mode is recommended if medias are more then 0.1 million."
  ElseIf Index = 5 Then
  lblStatus.Caption = "Determines wheather the autorun will be created or not after making Album." & vbCrLf & _
  "If you burn webalbum on a CD then autorun will start the page main.html whenever CD runs"
 End If
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = "When Browser is unable to load any media then an alternative text will be displayed inplace of media . You can type any alternative text here otherwise the Name of the media will be used as alternative text (Recommended.)"
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MoveOnce Then
  Call Dir1_MouseMove(0, 0, 0, 0)
  Command1.BackColor = vbWhite
  For Button = 0 To 2   'used temporary
   cmdOpt(Button).BackColor = vbWhite
  Next
 MoveOnce = False
End If
End Sub

Private Sub frBack_DblClick()
PopupMenu Preview
End Sub

Private Sub frBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Preview
End Sub

Private Sub frBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = " Write click or click on green ball to set the Background color and Text color of the webpages.Move your mouse over color to see the hexadecimal color code. "
End Sub

Private Sub frText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Preview
End Sub

Private Sub imgPreviewSet_Click()
PopupMenu Preview
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 If Index = 2 Then lblStatus.Caption = "These extention determines the type of media files to be added on your WebAlbum at the time of scanning folders"
End Sub

Private Sub mnuPreview_Click(Index As Integer)
' since ullu ka pattha vb accept BGR color
'so you have to do some kasrat to get RGB color for web pages
'show the BGR color in VB and store RGB color for webpage

Dim CDFlags As Long
Dim Lal As Integer, Hara As Integer, Nila As Integer
Dim Rang As Long

On Error GoTo ColorError

    CDFlags = &H2 + &H8 + &H1 'CDFlags + Check2(i).Value * Val(Check2(i).Tag)

    CommonDialog1.Flags = CDFlags
    CommonDialog1.CancelError = True
        CommonDialog1.ShowColor
   If Index = 0 Then
 
    Rang& = CommonDialog1.Color      'obtained BGR color
    'now convert it in to RGB color
    frBack.BackColor = Rang&    'long value of color
   Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    hcodeBACK = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)
   frBack.ToolTipText = hcodeBACK
   ElseIf Index = 1 Then
   Rang& = CommonDialog1.Color
   frText(0).ForeColor = Rang&
   Lal = Rang& Mod 256
    Hara = ((Rang& And &HFF00FF00) / 256&)
    Nila = (Rang& And &HFF00000) / 65536
    hcodeTEXT = "#" & Hex(Lal) & Hex(Hara) & Hex(Nila)
    frText(0).ToolTipText = hcodeTEXT
   End If
     
    
    
      'txtDefault.Text = hcodeBACK & "  " & hcodeTEXT ' (frText.ForeColor)
    Exit Sub
ColorError:
    If Err.Number = 32755 Then
        MsgBox "You have not select any color"
    
    Else
        MsgBox "An error occured"
    End If

End Sub

Private Sub optAlt_GotFocus(Index As Integer)
Call Frame3_MouseMove(4, 0, 0, 0)
If Index = 1 Then
  txtAlt.Enabled = False
Else
  txtAlt.Enabled = True
End If
End Sub

Private Sub optInfo_Click(Index As Integer)
If Index = 3 Then
 txtInfo.Enabled = False
 vscrollRes(4).Enabled = False
 txtRes(4).Enabled = False
Else
 txtInfo.Enabled = True
 vscrollRes(4).Enabled = True
 txtRes(4).Enabled = True
End If
End Sub

Private Sub optMode_Click(Index As Integer)
If Index = 0 Then
  Timer1.Interval = 0
  frmAlbum.Caption = scroll
Else
  Timer1.Interval = 50
 End If
End Sub

Private Sub optMode_GotFocus(Index As Integer)
Call Frame2_MouseMove(4, 0, 0, 0, 0)
End Sub

Private Sub optNo_Click(Index As Integer)
If Index = 1 Then
  mediasinrow = 3
  txtRes(3).Enabled = False
   vscrollRes(3).Enabled = False
Else
   txtRes(3).Enabled = True
   vscrollRes(3).Enabled = True
 End If
End Sub

Private Sub optNo_GotFocus(Index As Integer)
Call Frame2_MouseMove(1, 0, 0, 0, 0)
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblStatus.Caption = "Place the cursor or set the focus (using TAB ) over any control to get information ????  ???? ???? ???? ....... ......... ........"
End Sub

Private Sub Timer1_Timer()
''lntop2 green lntop3 white
'If lnTop(2).X2 < Line1(1).X2 Then
  lnTop(2).X2 = lnTop(2).X2 + 64
  lnTop(3).X2 = lnTop(3).X2 + 64
'End If

If (lnTop(2).X2 - lnTop(2).X1) > ((Line1(1).X2 - Line1(1).X1) / 2) Then
'Timer1.Interval = 0
   lnTop(2).X1 = lnTop(2).X1 + 64
   lnTop(3).X1 = lnTop(3).X1 + 64
End If
If lnTop(2).X2 > Line1(1).X2 Then
     lnTop(2).X1 = Line1(1).X1
     lnTop(3).X1 = Line1(1).X1
     lnTop(2).X2 = lnTop(2).X1
     lnTop(3).X2 = lnTop(3).X1
End If
DoEvents
delay = delay + 1
'//// show scrolling
If delay = 5 Then
 delay = 0
  frmAlbum.Caption = Mid$(scroll, madhya, Len(scroll) - madhya)
   frmAlbum.Caption = frmAlbum.Caption & Mid$(scroll, 1, madhya) 'temp
   madhya = madhya + 1
    If madhya > Len(scroll) Then madhya = 1
  
End If
 '/////////////////////

End Sub

Private Sub txtInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
txtInfo.SelStart = 0
txtInfo.SelLength = Len(txtInfo.Text)

End Sub

Private Sub txtRes_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vscrollRes(Index).SetFocus
End Sub

Private Sub vscrollRes_Change(Index As Integer)
If vscrollRes(Index).Value > vscrollRes(Index).Min And vscrollRes(Index).Value < vscrollRes(Index).Max Then
   txtRes(Index).Text = vscrollRes(Index).Value
   If Index = 3 Then mediasinrow = vscrollRes(Index).Value
   If Index = 4 Then txtRes(Index).Text = "H" & vscrollRes(Index).Value
End If
End Sub

Private Sub vscrollRes_GotFocus(Index As Integer)
    If Index = 0 Then
      Call Frame2_MouseMove(0, 0, 0, 0, 0)
    ElseIf Index = 2 Then
     Call Frame1_MouseMove(0, 0, 0, 0)
    ElseIf Index = 3 Then
     Call Frame2_MouseMove(1, 0, 0, 0, 0)
    End If
    
End Sub



