VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "NotePad On Top"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8820
   Icon            =   "Notepadontop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Editor 
      Height          =   5025
      HideSelection   =   0   'False
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   15
      Width           =   8790
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -15
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2.54052e-29
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu FileNew 
         Caption         =   "New"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu FileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu EditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu EditSelect 
         Caption         =   "Select All"
      End
      Begin VB.Menu EditSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu EditFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu ProcessMenu 
      Caption         =   "F&ormat"
      Begin VB.Menu ProcessUpper 
         Caption         =   "Upper Case"
      End
      Begin VB.Menu ProcessLower 
         Caption         =   "Lower Case"
      End
      Begin VB.Menu ProcessNumber 
         Caption         =   "Number Lines"
      End
   End
   Begin VB.Menu CustomMenu 
      Caption         =   "&View"
      Begin VB.Menu CustomFont 
         Caption         =   "Font"
      End
      Begin VB.Menu CustomPage 
         Caption         =   "Page Color"
      End
      Begin VB.Menu CustomText 
         Caption         =   "Text Color"
      End
   End
   Begin VB.Menu ontop 
      Caption         =   "&Set On Top"
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "About"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "http:\\vinodkotiya.tripod.com"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OpenFile As String
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40

Private Sub Form_Load()
 ontop_Click
End Sub


Private Sub CustomFont_Click()
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    Editor.Font = CommonDialog1.FontName
    Editor.FontBold = CommonDialog1.FontBold
    Editor.FontItalic = CommonDialog1.FontItalic
    Editor.FontSize = CommonDialog1.FontSize
End Sub

Private Sub CustomPage_Click()
    CommonDialog1.ShowColor
    Editor.BackColor = CommonDialog1.Color
End Sub

Private Sub CustomText_Click()
On Error Resume Next
    CommonDialog1.ShowColor
    Editor.ForeColor = CommonDialog1.Color
End Sub

Private Sub EditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Editor.SelText
End Sub

Private Sub EditCut_Click()

    Clipboard.SetText Editor.SelText
    Editor.SelText = ""

End Sub

Private Sub EditFind_Click()
    Form2.Show
End Sub

Private Sub EditPaste_Click()
    If Clipboard.GetFormat(vbCFText) Then
        Editor.SelText = Clipboard.GetText
    Else
        MsgBox "Invalid Clipboard format."
    End If
End Sub

Private Sub EditSelect_Click()
    Editor.SelStart = 0
    Editor.SelLength = Len(Editor.Text)
End Sub

Private Sub FileExit_Click()
    End
End Sub

Private Sub FileNew_Click()
    Editor.Text = ""
    OpenFile = ""
End Sub

Private Sub FileOpen_Click()
Dim FNum As Integer
Dim txt As String

On Error GoTo FileError
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.DefaultExt = "TXT"
    CommonDialog1.Filter = "Text files|*.TXT|All files|*.*"
    CommonDialog1.ShowOpen
    FNum = FreeFile
    Open CommonDialog1.FileName For Input As #1
    txt = Input(LOF(FNum), #FNum)
    Close #FNum
    Editor.Text = txt
    OpenFile = CommonDialog1.FileName
    Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & CommonDialog1.FileName
    OpenFile = ""
    
End Sub

Private Sub FileSave_Click()
Dim FNum As Integer
Dim txt As String

    If OpenFile = "" Then
        FileSaveAs_Click
        Exit Sub
    End If
On Error GoTo FileError
    FNum = FreeFile
    Open OpenFile For Output As #1
    Print #FNum, Editor.Text
    Close #FNum
    Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file " & OpenFile
    OpenFile = ""

End Sub

Private Sub FileSaveAs_Click()
Dim FNum As Integer
Dim txt As String

On Error GoTo FileError
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.DefaultExt = "TXT"
    CommonDialog1.Filter = "Text files|*.TXT|All files|*.*"
    CommonDialog1.ShowSave
    FNum = FreeFile
    Open CommonDialog1.FileName For Output As #1
    Print #FNum, Editor.Text
    Close #FNum
    OpenFile = CommonDialog1.FileName
    Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file " & CommonDialog1.FileName
    OpenFile = ""
End Sub


Private Sub Form_Resize()
    Editor.Width = Form1.Width - 15 * Screen.TwipsPerPixelX
    Editor.Height = Form1.Height - 50 * Screen.TwipsPerPixelY
End Sub


Private Sub mnuHelp_Click(Index As Integer)
If Index = 0 Then
 MsgBox "Hey ||| Are you sure that u need any help."
ElseIf Index = 1 Then
MsgBox "This Notepad is made for vinod kotiya by vinod kotiya" & vbCrLf & _
  "So that i can debugg my projects by pasting check list on top of all windows." & vbCrLf & _
  "Thus i can utilize my monitor's large view area."
End If
End Sub

Private Sub ProcessLower_Click()
Dim Sel1 As Integer, Sel2 As Integer
    
    Sel1 = Editor.SelStart
    Sel2 = Editor.SelLength
    Editor.SelText = LCase$(Editor.SelText)
    Editor.SelStart = Sel1
    Editor.SelLength = Sel2
End Sub

Private Sub ProcessNumber_Click()
Dim tmpText As String, tmpLine As String
Dim firstChar As Integer, lastChar As Integer
Dim currentLine As Integer

firstChar = 1
currentLine = 1
lastChar = InStr(Editor.Text, Chr$(10))
While lastChar > 0
    tmpLine = Format$(currentLine, "000") & "  " & Mid$(Editor.Text, firstChar, lastChar - firstChar + 1)
    currentLine = currentLine + 1
    firstChar = lastChar + 1
    lastChar = InStr(firstChar, Editor.Text, Chr$(10))
    tmpText = tmpText + tmpLine
Wend
Editor.Text = tmpText
End Sub

Private Sub ProcessUpper_Click()
Dim Sel1, Sel2 As Integer

    Sel1 = Editor.SelStart
    Sel2 = Editor.SelLength
    Editor.SelText = UCase$(Editor.SelText)
    Editor.SelStart = Sel1
    Editor.SelLength = Sel2
End Sub

Private Sub ontop_Click()
 If ontop.Caption = "Set On Top" Then
   SetWindowPos Me.hwnd, HWND_TOPMOST, 30, 20, 600, 400, SWP_SHOWWINDOW
   ontop.Caption = "Remove from Top"
  Else
  SetWindowPos Me.hwnd, -2, 30, 20, 600, 400, SWP_SHOWWINDOW
   ontop.Caption = "Set On Top"
  End If
End Sub
