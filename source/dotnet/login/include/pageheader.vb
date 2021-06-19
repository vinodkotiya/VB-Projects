Imports System
Imports System.Web
Imports PortalModuleControl

Public Class PageHeader : Inherits PortalModuleControl

    Private _ShowSignOut As Boolean = True

    Public Property ShowSignOut As Boolean
        Get
            Return _ShowSignOut
        End Get
        Set
            _ShowSignOut = Value
        End Set
    End Property

    Protected Sub SignOff_Click(sender As Object, e As EventArgs )

       Dim PageIndex As HttpCookie = New HttpCookie("_PageIndex", "0")

       PageIndex.Path = "/"
       PageIndex.Expires = new DateTime(2002, 10, 10)
       Response.AppendCookie(PageIndex)
       System.Web.Security.FormsAuthentication.SignOut()
       Response.Redirect("/Quickstart/aspplus/samples/portal/VB/default.aspx")

    End Sub

End Class