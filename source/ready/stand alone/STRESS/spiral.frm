VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Spiral"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "spiral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   2760
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   9120
      Left            =   0
      ScaleHeight     =   604
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   804
      TabIndex        =   0
      Top             =   0
      Width           =   12120
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Close"
         Height          =   255
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8040
         Width           =   840
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1200
         Top             =   1320
      End
      Begin VB.Label txtScroll 
         BackColor       =   &H00000000&
         Caption         =   "NOW IT IS TIME TO TAKE A BREAK.RELAX YOUR EYES,MOVE YOUR NECK LEFT AND RIGHT 20 TIMES ,SIT IN A WRITE POSTURE :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   8400
         Width           =   10920
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScrollLab As Long
'Dim makenotop As Boolean
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
 '   ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
  '  ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Dim toleft As Boolean  'it is false when tip come from right
'but become true when going toleft after rest

'for tips

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "data\tipofmin.vin"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long

'Const HWND_TOPMOST = -1
'Const HWND_NOTOPMOST = -2
'Const SWP_SHOWWINDOW = &H40

Private Sub Command1_Click()

frmStart.Show
'Load frmStart
'frmStart.Visible = True
Unload Me
End Sub



Private Sub Form_Load()
    PenColor = RGB(0, 0, 0)
    ScrollLab = 1
    trignometry = 3
   Dim retValue As Long
    'Load Form1
  '  retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, _
   '        1024, 768, SWP_SHOWWINDOW)
    
    If Screen.Height > 9000 Then
     Form1.Height = 768
     Form1.Width = 1024
     Picture1.Height = 768
     Picture1.Width = 1024
     txtScroll.Width = Picture1.Width - txtScroll.Left
   End If
    Randomize
    ' Read in the tips file and display a tip at random.
    'If LoadTips(App.Path & "\" & TIP_FILE) = False Then
     '   MsgBox " Tip of the minute file " & TIP_FILE & " was not found? "
      '  txtScroll.Caption = " Tip of the minute file " & TIP_FILE & " was not found? Please run troubleshooter"
    'End If
    LoadTips (App.Path & "\" & TIP_FILE)
   txtScroll.Top = Picture1.ScaleHeight - 40
    Command1.Top = Picture1.ScaleHeight - 50
    Command1.Left = Picture1.ScaleWidth - 100
     End Sub


Private Sub Timer1_Timer()
DrawRoullette    'draw ek ke upar ek


End Sub
Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        txtScroll.Caption = Tips.Item(CurrentTip)
    End If
End Sub
Private Sub DoNextTip()
    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd(43433)) + 1)
    Form1.DisplayCurrentTip
 End Sub
 Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub WOption_Click(Index As Integer)
   
End Sub

Private Sub Timer2_Timer()
If ScrollLab > 750 Then      'to display scroll label
Form1.Picture1.Cls          'it is done b/c picture box
ScrollLab = 0           'will clear after 15 sec  'and label need 100 milisec
'used to set trignometry
  If trignometry = 1 Then
    trignometry = 2       'for tan
  ElseIf trignometry = 2 Then
    trignometry = 3      'for sqr sincos
  ElseIf trignometry = 3 Then
    trignometry = 1     'for sincos
  End If
End If
ScrollLab = ScrollLab + 1
If ScrollLab Mod 375 = 0 Then DoNextTip

'If makenotop = True Then
 'notop
 'makenotop = False
'End If
'Form1.txtScroll.Left = Form1.txtScroll.Left - 3
'If Form1.txtScroll.Left + Form1.txtScroll.Width < 0 Then
' Form1.txtScroll.Left = Form1.txtScroll.Width  'ScrollLab = 0
'End If

End Sub

Private Sub Timer3_Timer()


End Sub
Private Sub nottop()
Dim reetValue As Long
   
  '  reetValue = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, _
   '           1024, 768, SWP_SHOWWINDOW)
End Sub


