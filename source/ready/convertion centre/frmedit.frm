VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Delete (VIN Conversion Centre)"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "frmedit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Restore All Data Files"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<<----------    Delete This Unit      "
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cmbDel 
      Height          =   315
      Left            =   1440
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add/Delete In Selected Group"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton opt 
         BackColor       =   &H00FFFF80&
         Caption         =   "Currency"
         Height          =   255
         Index           =   6
         Left            =   6120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FFFF80&
         Caption         =   "Volume"
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FFFF80&
         Caption         =   "Area"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FFFF80&
         Caption         =   "Temperature"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FFFF80&
         Caption         =   "Mass"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FFFF80&
         Caption         =   "Length"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FFFF80&
         Caption         =   "Time"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbOther 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtAns 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtShort 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Any Existing Unit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Add A New Unit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   2640
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      Height          =   735
      Left            =   120
      Top             =   2640
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   1875
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1455
      Left            =   120
      Top             =   960
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label lblOne 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Solve This"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Short Form of New Unit (eg. sec,gm)"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of New Unit (eg. second,Gram)"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      Height          =   1455
      Left            =   120
      Top             =   960
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datafile As String
Dim converter(60) As Double    'holds the inverse multiple w.r.t mks like 100,corresponding for cm
Dim units(60) As String    'holds the unit shortform
Dim answer As String


Private Sub cmdAdd_Click()
Dim iofcmbOther As Integer    'store index of cmbother which will deleted
iofcmbOther = cmbOther.ListIndex
answer = Str(converter(iofcmbOther) / CDbl(txtAns.Text))
Dim Fsys As New FileSystemObject
Dim Tstream As TextStream
Set Tstream = Fsys.OpenTextFile(datafile, ForAppending)
Tstream.WriteBlankLines (1)
Tstream.WriteLine (txtName.Text)
Tstream.WriteLine (txtShort.Text)
Tstream.Write (Trim(answer))
Tstream.Close

End Sub

Private Sub cmdDel_Click()
Dim iofDel As Integer 'store index to be deleted
iofDel = cmbDel.ListIndex
Dim Fsys As New FileSystemObject
Dim Tstream As TextStream
Set Tstream = Fsys.OpenTextFile(datafile, ForWriting)
Dim i As Integer
On Error GoTo FileError
    'FNum = FreeFile
    'Open datafile For Output As #1
     For i = 0 To cmbDel.ListCount - 1
      If i <> iofDel Then
       
       Tstream.WriteLine (cmbDel.List(i))
       Tstream.WriteLine (units(i))
       Tstream.Write (converter(i))
       If i = cmbDel.ListCount - 2 Then Exit For   'for last item not write a line
       Tstream.WriteBlankLines (1)
       'Print #FNum, cmbDel.List(i)
       'Print #FNum, Chr(13)
       'Print #FNum, units(i)
       'Print #FNum, Chr(13)
       'Print #FNum, converter(i)
       'Print #FNum, Chr(13)
      End If
      Next
    Tstream.Close
    cmbDel.RemoveItem iofDel
    Exit Sub
FileError:
  MsgBox "can't delete"
End Sub

Private Sub Command1_Click()
Dim Fsys As New FileSystemObject
Dim reply As Integer
On Error GoTo vinerror
If Fsys.FolderExists("c:\windows\vinbakup") = False Then 'if folder not exist create it
  MsgBox "Backup was not created so unable to Restore"
End If
'Fsys.CopyFolder App.Path & "\data", "c:\windows\vinbakup", True
reply = MsgBox("Are you sure to restore all the data files to previous state", vbYesNo)
If reply = vbYes Then
Fsys.CopyFile "c:\windows\vinbakup\time.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\length.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\mass.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\area.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\temperature.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\volume.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\factors.vin", App.Path & "\data\", True
Fsys.CopyFile "c:\windows\vinbakup\others.vin", App.Path & "\data\", True
MsgBox "Restore successfully"
End If
Exit Sub
vinerror:
 MsgBox "file handling error occured"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOne.Caption = "1 " & txtName.Text '& " = "
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load Form1
Form1.Visible = True
Unload Me
End Sub

Private Sub opt_Click(Index As Integer)
Dim i As Integer

If cmbOther.ListCount > 0 Then

 For i = cmbOther.ListCount - 1 To 0 Step -1
  cmbOther.RemoveItem (i)
  cmbDel.RemoveItem (i)
 Next
End If

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Shape1.Visible = True
Shape2.Visible = True
Shape3.Visible = True
Shape4.Visible = True
lblOne.Visible = True
txtName.Visible = True
txtShort.Visible = True
txtAns.Visible = True
cmbOther.Visible = True
cmbDel.Visible = True
cmdDel.Visible = True
cmdAdd.Visible = True
If Index = 0 Then
 datafile = App.Path & "\data\time.vin"
ElseIf Index = 1 Then
 datafile = App.Path & "\data\length.vin"
ElseIf Index = 2 Then
 datafile = App.Path & "\data\mass.vin"
ElseIf Index = 3 Then
 datafile = App.Path & "\data\temperature.vin"
ElseIf Index = 4 Then
 datafile = App.Path & "\data\area.vin"
ElseIf Index = 5 Then
 datafile = App.Path & "\data\volume.vin"
ElseIf Index = 6 Then
' datafile = App.Path & "\data\currency.vin"
MsgBox "Currently Disabled"
Exit Sub
End If
loadcombo
End Sub
Private Sub loadcombo()
Dim FNum As Integer
Dim currentline As String
Dim tmpdouble As Double
Dim i As Integer
i = 0

On Error GoTo FileError
   
    FNum = FreeFile    'getting file no for futures referance
    Open datafile For Input As FNum
    While Not EOF(FNum)
        Line Input #FNum, currentline
        cmbOther.AddItem currentline
        cmbDel.AddItem currentline
        Line Input #FNum, units(i)
        Line Input #FNum, currentline
        tmpdouble = CDbl(currentline)
        converter(i) = tmpdouble
        'MsgBox currentline
        i = i + 1
    Wend
  '  txt = Input(LOF(FNum), #FNum)
    Close #FNum
   Exit Sub

FileError:
    MsgBox "Unkown error while opening file " & "mass.vin" _
     & "file is effected by any fool "

End Sub

Private Sub txtName_Change()
lblOne.Caption = "1 " & txtName.Text
cmdAdd.Caption = "Add " & txtName.Text & " As New Unit"
End Sub
