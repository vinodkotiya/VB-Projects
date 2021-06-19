Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports PortalModulePage
Imports Microsoft.VisualBasic

Public Class Options : Inherits PortalModulePage

    Protected Sub Colors_Change(sender As Object, evt As EventArgs)

        Dim button As HtmlInputRadioButton = CType(sender, HtmlInputRadioButton)

        If ((button.Checked) And (String.Compare(UserState("UserId"),"ANONYMOUS") <> 0)) Then
            Dim colorValues As String = button.Value
            Dim colorList() As String = Split(colorValues, ",")

            UserState("ColorScheme") = colorList(0)
            UserState("HeadColor") = colorList(1)
            UserState("SubheadColor") = colorList(2)
            UserState("BackColor") = colorList(3)
            UserState("LeftColor") = colorList(4)
        End If
    End Sub

    Protected Sub Submit_Click(sender As Object, evt As EventArgs)
        Response.Redirect("/Quickstart/aspplus/samples/portal/VB/default.aspx")
    End Sub

End Class