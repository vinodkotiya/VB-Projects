VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfCal 
      Height          =   2175
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"calandermaker.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbYears 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim year As String


Private Sub Command1_Click()

'MsgBox Weekday(Date)
makecalander
 rtfCal.Text = rtfCal.Text & "</TABLE> "
  rtfCal.SaveFile "C:\TEMP.HTML", rtfText
  
End Sub

Private Sub Form_Load()
cmbYears.AddItem "2002"
cmbYears.AddItem "2003"
cmbYears.AddItem "2004"
cmbYears.AddItem "2005"
cmbYears.AddItem "2006"
cmbYears.AddItem "2007"
cmbYears.AddItem "2008"
cmbYears.AddItem "2009"
cmbYears.AddItem "2010"
cmbYears.ListIndex = 0
'loadrtf
End Sub
Private Sub makecalander()

Dim thismonth As Integer   'store the current month no for comparison when date exceed from 31

Dim j As Integer  'for changing month

bigtable
rtfCal.Text = rtfCal.Text & "<TR>"
loadweekdays
rtfCal.Text = rtfCal.Text & "</TR><tr>"
year = "1-" & "1" & "-" & cmbYears.List(cmbYears.ListIndex)
fill1to7 (year)
year = "1-" & "2" & "-" & cmbYears.List(cmbYears.ListIndex)
fill1to7 (year)
year = "1-" & "3" & "-" & cmbYears.List(cmbYears.ListIndex)
fill1to7 (year)
rtfCal.Text = rtfCal.Text & "</TR><tr>"
year = "1-" & "1" & "-" & cmbYears.List(cmbYears.ListIndex)
fill7to14 (year)
year = "1-" & "2" & "-" & cmbYears.List(cmbYears.ListIndex)
fill7to14 (year)
year = "1-" & "3" & "-" & cmbYears.List(cmbYears.ListIndex)
fill7to14 (year)
rtfCal.Text = rtfCal.Text & "</TR><tr>"
End Sub
Private Sub fill7to14(year As String)
Dim daysuptofilled As Integer
Dim i As Integer
daysuptofilled = 7 - WEEKDAY(year) + 1
For i = daysuptofilled To 7 + daysuptofilled
      rtfCal.Text = rtfCal.Text + "<TD>" & i & "</TD>"
Next
End Sub
Private Sub fill1to7(year As String)
'For j = 1 To 12
Dim i As Integer     'for chnging days
Dim WeekDaySpace As Integer
'loadrtf
'thismonth = Month(year)

Dim ONLYONE As Boolean
ONLYONE = True 'NOW U CAN ENTER IN FOR LOOP TO CREATE BLANKS IF PEHLI TARIKH NOT SUN
 For i = 1 To i - WEEKDAY(year) + 1
 
  'If Month(year) <> thismonth Then
  ' Exit For                             'IF MONTH CROSS 31 THEN EXIT
  'End If
  
  If ONLYONE = True Then
   For WeekDaySpace = WEEKDAY(year) To 2 Step -1   'CREATE BLANKS IF PEHLI TARIKH NOT SUNDAY
      rtfCal.Text = rtfCal.Text + "<TD>YO</TD>"
   Next
   ONLYONE = False
  End If
  
  rtfCal.Text = rtfCal.Text & "<TD><B>" & Day(year) & "</B></TD>"
  
  'If WEEKDAY(year) = 7 Then            'WEN SAT COME CHANGE ROW
  ' rtfCal.Text = rtfCal.Text & "</TR>"
  'End If
  
  'year = DateAdd("d", 1, year) 'INCREAMENT THE MONTHS DAYS
  
  'If WEEKDAY(year) = 1 Then            'WEN SUN COME START ROW
  ' rtfCal.Text = rtfCal.Text & "<TR>"
  'End If

 Next
'   rtfCal.Text = rtfCal.Text & "</table><br><table>"

End Sub
Private Sub loadweekdays()
Dim i As Integer
For i = 1 To 3
rtfCal.Text = rtfCal.Text + Chr(32) & Chr(32) & Chr(32) & "<TD><font color =red><B>SUN</B></FONT></TD>" & Chr(13) & _
        Chr(32) & Chr(32) & Chr(32) & "<TD><B>MON</B></TD>" & Chr(13) & _
        Chr(32) & Chr(32) & Chr(32) & "<TD><B>TUE</B></TD>" & Chr(13) & _
        Chr(32) & Chr(32) & Chr(32) & "<TD><B>WED</B></TD>" & Chr(13) & _
        Chr(32) & Chr(32) & Chr(32) & "<TD><B>THU</B></TD>" & Chr(13) & _
        Chr(32) & Chr(32) & Chr(32) & "<TD><B>FRI</B></TD>" & Chr(13) & _
        Chr(32) & Chr(32) & Chr(32) & "<TD><B>SAT</B></TD>" & Chr(13)
Next
End Sub
Private Sub bigtable()
  rtfCal.Text = rtfCal.Text + Chr(32) & Chr(32) & Chr(32) & "<TR width = 100%><font color =black><B><td width=33%>January</TD><TD width=33%> February </TD><TD width=330>March</TD></B></Font></TR>" & Chr(13)
End Sub
