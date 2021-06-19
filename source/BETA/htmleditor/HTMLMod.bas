Attribute VB_Name = "HTMLPadModule"
Sub RenderDocument()

    HTMLEdit.WebBrowser1.Document.Script.Document.Clear
    HTMLEdit.WebBrowser1.Document.Script.Document.Write HTMLEdit.RichTextBox1.Text
    HTMLEdit.WebBrowser1.Document.Script.Document.Close
    Exit Sub

End Sub

