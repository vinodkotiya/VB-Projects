VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FF8080&
   Caption         =   "Search Results"
   ClientHeight    =   9105
   ClientLeft      =   5940
   ClientTop       =   645
   ClientWidth     =   9120
   Icon            =   "mdifrmfone1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu menuSaveWeb 
      Caption         =   "Save As &Web Page"
   End
   Begin VB.Menu menuSaveText 
      Caption         =   "Save As &Text File"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Initialize()

MDIForm1.Show
MDIForm1.Top = 20
    MDIForm1.Enabled = False
    frmShow.Show
    MDIForm1.Height = frmShow.Height + 70 * Screen.TwipsPerPixelY
    MDIForm1.Enabled = True
End Sub

Private Sub menuSaveText_Click()
Dim FNum As Integer
Dim txt As String
Dim i As Integer
On Error GoTo FileError
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.DefaultExt = "VIN"
    CommonDialog1.Filter = "Vin files |*.vin |Text files|*.TXT|All files|*.*"
    CommonDialog1.ShowSave
    FNum = FreeFile
    
    Open CommonDialog1.FileName For Output As #1
    For i = 0 To Val(frmShow.totalfound.Text) Step 1
    
         Print #FNum, frmShow.txtName(i).Text & "   " & frmShow.txtsname(i).Text
         Print #FNum, frmShow.txtPost(i).Text
         'Print #FNum, frmShow.txtAddress(i).Text
         Print #FNum, frmShow.txtArea(i).Text
         Print #FNum, frmShow.txtCity(i).Text
         Print #FNum, frmShow.txtStd(i).Text & " " & frmShow.txtFoneo(i).Text _
          & "  "; frmShow.txtfoner(i).Text & "  " & frmShow.txtMobile(i).Text
         'Print #FNum, frmShow.txtEmail(i).Text
         Print #FNum, "***********************************************************************"
    Next
     Print #FNum, "Fone directory by vinod kotiya "
     Print #FNum, txt
    
    Close #FNum
    'OpenFile = CommonDialog1.FileName
    Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file " & CommonDialog1.FileName
    'OpenFile = ""
End Sub



Private Sub menuSaveWeb_Click()
Dim txt As String
Dim FNum As Integer

On Error GoTo FileError:
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = "html"
    CommonDialog1.Filter = "HTML Documents|*.html|RTF Files|*.RTF|Text Files|*.TXT|All Files|*.*"
    CommonDialog1.ShowSave
    
    frmShow.rtxtWeb.SaveFile CommonDialog1.FileName, rtfText
    'OpenFile = CommonDialog1.FileName
    Exit Sub
    
FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & CommonDialog1.FileName
    'OpenFile = ""

End Sub
