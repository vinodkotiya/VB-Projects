Attribute VB_Name = "HTMLPadModule"
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40

Sub RenderDocument()
On Error GoTo vinerror
    HTMLEdit.WebBrowser1.Document.Script.Document.Clear
    HTMLEdit.WebBrowser1.Document.Script.Document.Write HTMLEdit.RichTextBox1.Text
    HTMLEdit.WebBrowser1.Document.Script.Document.Close
    Exit Sub
vinerror:
End Sub


