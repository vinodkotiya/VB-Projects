VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmShow 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   -645
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox totalfound 
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Text            =   "store total founds"
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin RichTextLib.RichTextBox rtxtWeb 
      Height          =   2415
      Left            =   1800
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmShow.frx":0000
   End
   Begin VB.CommandButton frPerson 
      Caption         =   "Format"
      Height          =   255
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtsname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   11
      Text            =   " Surname (Kotiya)"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Text            =   " Name (Vinod)"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtPost 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Text            =   " Designation (Student)"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Text            =   " Address (s-2 Shrimaya Apartment Sector -B/363 Sarvdharm Colony)"
      Top             =   1560
      Width           =   8295
   End
   Begin VB.TextBox txtArea 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Text            =   "Area (Kolar Road)"
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtCity 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   6
      Text            =   "City   (Bhopal)"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtStd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Text            =   "STD(0755)"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtFoneo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Text            =   "Fone(R) 2794428"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtfoner 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Text            =   "Fone(O)   2794428"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   2
      Text            =   "Mobile (9********)"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Text            =   "Emails:- vinner24@hotmail.com ; www.webduniya.com"
      Top             =   2640
      Width           =   8295
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Height          =   2055
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   8535
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
 rtxtWeb.Text = rtxtWeb.Text + "''" + frmStart.txtSearch.Text + "''</marquee></font></h2>"
Dim i As Integer
On Error GoTo SQLError
Form3.Show
Form3.Refresh
MDIForm1.Caption = "Search Results For '" & frmStart.txtSearch.Text & "'"
   
Form1.Datafone.Refresh
Form1.Datafone.Recordset.MoveLast
Form1.Datafone.Recordset.MoveFirst
For i = 1 To Form1.Datafone.Recordset.RecordCount     'run loop upto last record
  Form1.Datafone.Recordset.FindNext Form1.GenerateSQL
    If Form1.Datafone.Recordset.NoMatch Then  'if no more matched record found then please exit from loop
        Form3.ProgressBar1.Value = 100   'full progress
        Beep
        Unload Form3                  'unload progressbar
         If i - 1 = 0 And frmStart.txtSearch.Text <> "" Then
          MsgBox "No match found | Please modify your search string ." _
           & "Recomended Strings are : " & frmStart.txtSearch.Text & "*  ," _
           & " *" & frmStart.txtSearch.Text & "  *" & frmStart.txtSearch.Text & " *  ," _
           & Mid(frmStart.txtSearch.Text, 1, 3) & "*"
        End If
        txtResult.Visible = True
        txtResult.Text = "Search Complete !!!! " _
                        & String(25, " ") & "There are " & Str(i - 1) & " Records found for the '" _
                        & frmStart.txtSearch.Text & "'"
        rtxtWeb.Text = rtxtWeb.Text + "</body></html>                "
        totalfound.Text = Str(i - 1)    'use globaly total found
       
           
        Exit Sub
     End If
'loading controls jane kya hoga rama re jaane kya hoga moula re
  Load Shape1(i)
  Load frPerson(i)
  Load txtName(i)
  Load txtsname(i)
  Load txtPost(i)
  Load txtAddress(i)
  Load txtArea(i)
  Load txtCity(i)
  Load txtStd(i)
  Load txtFoneo(i)
  Load txtfoner(i)
  Load txtMobile(i)
    Load txtEmail(i)
 'setting position of controls at runtime
 Shape1(i).Top = Shape1(i - 1).Top + 2200
  frPerson(i).Top = frPerson(i - 1).Top + 2200
   txtName(i).Top = txtName(i - 1).Top + 2200
   txtsname(i).Top = txtsname(i - 1).Top + 2200
   txtPost(i).Top = txtPost(i - 1).Top + 2200
   txtAddress(i).Top = txtAddress(i - 1).Top + 2200
   txtArea(i).Top = txtArea(i - 1).Top + 2200
   txtCity(i).Top = txtCity(i - 1).Top + 2200
   txtStd(i).Top = txtStd(i - 1).Top + 2200
   txtFoneo(i).Top = txtFoneo(i - 1).Top + 2200
   txtfoner(i).Top = txtfoner(i - 1).Top + 2200
   txtMobile(i).Top = txtMobile(i - 1).Top + 2200
   txtEmail(i).Top = txtEmail(i - 1).Top + 2200
'making visible
Shape1(i).Visible = True
  frPerson(i).Visible = True
   txtName(i).Visible = True
   txtsname(i).Visible = True
   txtPost(i).Visible = True
   txtAddress(i).Visible = True
   txtArea(i).Visible = True
   txtCity(i).Visible = True
   txtStd(i).Visible = True
  txtFoneo(i).Visible = True
   txtfoner(i).Visible = True
   txtMobile(i).Visible = True
   txtEmail(i).Visible = True
  Beep
'giving values to controls at runtime from form1
  frPerson(i).Caption = "Person" & Str(i)
   txtName(i).Text = Form1.Text1.Text
   txtsname(i).Text = Form1.Text2.Text
   txtPost(i).Text = Form1.Text3.Text
   txtAddress(i).Text = Form1.Text4.Text
   txtArea(i).Text = Form1.Text5.Text
   txtCity(i).Text = Form1.Text6.Text
   txtStd(i).Text = Form1.Text7.Text
   txtFoneo(i).Text = Form1.Text8.Text
   txtfoner(i).Text = Form1.Text9.Text
   txtMobile(i).Text = Form1.Text10.Text
   txtEmail(i).Text = Form1.Text11.Text
   'write on the rtxtWeb for web page
   writeweb (i)
   'write on the rtxtText for Text File
   
  ' writeText (i)
   frmShow.Height = txtEmail(i).Top + txtEmail(i).Height
  ' i = i + 1
  'display progress bar
  If Form3.ProgressBar1.Value < 51 Then
   Form3.ProgressBar1.Value = i * 2
  End If
Next
Unload Form3                  'unload progressbar
SQLError:
    MsgBox Err.Description

End Sub

Private Sub writeweb(i As Integer)

 rtxtWeb.Text = rtxtWeb.Text + frPerson(i).Caption
 rtxtWeb.Text = rtxtWeb.Text + "</br><center><b>"
 rtxtWeb.Text = rtxtWeb.Text + txtName(i).Text
 rtxtWeb.Text = rtxtWeb.Text + "  "
 rtxtWeb.Text = rtxtWeb.Text + txtsname(i).Text
 rtxtWeb.Text = rtxtWeb.Text + "</b></br>"
 rtxtWeb.Text = rtxtWeb.Text + txtPost(i).Text
 rtxtWeb.Text = rtxtWeb.Text + "</br>"
 rtxtWeb.Text = rtxtWeb.Text + txtAddress(i).Text
  rtxtWeb.Text = rtxtWeb.Text + txtArea(i).Text
  rtxtWeb.Text = rtxtWeb.Text + "</br>"
  rtxtWeb.Text = rtxtWeb.Text + txtCity(i).Text
  rtxtWeb.Text = rtxtWeb.Text + "</br>"
  rtxtWeb.Text = rtxtWeb.Text + "Fone: (" + txtStd(i).Text
  rtxtWeb.Text = rtxtWeb.Text + ") " + txtFoneo(i).Text
  rtxtWeb.Text = rtxtWeb.Text + "   "
 rtxtWeb.Text = rtxtWeb.Text + txtfoner(i).Text
 rtxtWeb.Text = rtxtWeb.Text + "      "
  rtxtWeb.Text = rtxtWeb.Text + txtMobile(i).Text
  rtxtWeb.Text = rtxtWeb.Text + "</br>"
  rtxtWeb.Text = rtxtWeb.Text + "Email:  " + txtEmail(i).Text
  rtxtWeb.Text = rtxtWeb.Text + "</center></br>"
rtxtWeb.Text = rtxtWeb.Text + "<hr>"
End Sub

'Private Sub writeText(i As Integer)

 'rtxtText.Text = rtxtText.Text + frPerson(i).Caption
 'rtxtText.Text = rtxtText.Text + "\par "
'rtxtText.Text = rtxtText.Text + txtName(i).Text
 
'rtxtText.Text = rtxtText.Text + txtsname(i).Text
' rtxtText.Text = rtxtText.Text + "\par "
'rtxtText.Text = rtxtText.Text + txtPost(i).Text
' rtxtText.Text = rtxtText.Text + "\par "
'rtxtText.Text = rtxtText.Text + txtAddress(i).Text
' rtxtText.Text = rtxtText.Text + txtArea(i).Text
 'rtxtText.Text = rtxtText.Text + "\par "
 'rtxtText.Text = rtxtText.Text + txtCity(i).Text
 'rtxtText.Text = rtxtText.Text + "\par "
 'rtxtText.Text = rtxtText.Text + "Fone: (" + txtStd(i).Text
 'rtxtText.Text = rtxtText.Text + ") " + txtFoneo(i).Text
 'rtxtText.Text = rtxtText.Text + "   "
'rtxtText.Text = rtxtText.Text + txtfoner(i).Text
'rtxtText.Text = rtxtText.Text + "      "
' rtxtText.Text = rtxtText.Text + txtMobile(i).Text
' rtxtText.Text = rtxtText.Text + "\par "
' rtxtText.Text = rtxtText.Text + "Email:  " + txtEmail(i).Text
' rtxtText.Text = rtxtText.Text + "\par "

'End Sub

Private Sub frPerson_Click(Index As Integer)

End Sub

Private Sub txtResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtName(0).SetFocus
End Sub
