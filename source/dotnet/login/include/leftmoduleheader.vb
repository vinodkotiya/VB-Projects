Imports System
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports PortalModuleControl
Imports Microsoft.VisualBasic

Public Class LeftModuleHeader : Inherits PortalModuleControl

    Public Title As String = "Left Module"
    Private _EditPage As String = "#edit"
    Private _ShowEditButton As Boolean = True
    Private _ShowCloseButton As Boolean = True
    Public ModuleSource As String = ""
    Public anchorEditPage as HtmlAnchor

    Public Property EditPage As String
        Get
            Return _EditPage
        End Get
        Set
            _EditPage = Value
        End Set
    End Property

    Public Property ShowEditButton As Boolean
        Get
            Return _ShowEditButton
        End Get
        Set
            _ShowEditButton = Value
        End Set
    End Property

    Public Property ShowCloseButton As Boolean
        Get
            Return _ShowCloseButton
        End Get
        Set
            _ShowCloseButton = Value
        End Set
    End Property

    Protected Sub Page_Load(sender As Object, e As EventArgs)

        If (String.Compare(UserState("UserId"),"ANONYMOUS") = 0) Then
            _EditPage = "/Quickstart/aspplus/samples/portal/VB/login.aspx"
        End If

        anchorEditPage.Attributes("style") = "color:" + UserState("HeadColor")
        DataBind()

    End Sub

   Protected Sub CloseButton_Click(sender As Object, e As EventArgs)

        If (String.Compare(UserState("UserId"),"ANONYMOUS") <> 0) Then

            Dim pageIndex As Integer = 0

            If (Not Request.Cookies("_PageIndex") Is Nothing) Then
                pageIndex = Int32.Parse(Request.Cookies("_PageIndex").Value)
      End If

      Dim leftModules As String = UserState("PageModules_" + pageIndex.ToString() + "L")
            Dim moduleList() As String = Split(leftModules, ";")

            Dim s As String = ""
      Dim i as Integer

            For i=0 To moduleList.Length-1
                If (String.Compare(ModuleSource, moduleList(i)) <> 0) Then
                    s += moduleList(i) + ";"
        End If
            Next i

            UserState("PageModules_" + pageIndex.ToString() + "L") = TrimEnd(s, ";")
            Response.Redirect("/Quickstart/aspplus/samples/portal/VB/default.aspx")
        Else
            Response.Redirect("/Quickstart/aspplus/samples/portal/VB/login.aspx")
        End If

    End Sub

  Private Function TrimEnd(source As String, trimchar as String) As String

    Dim tc() as Char = trimchar.ToCharArray()
    TrimEnd = source.TrimEnd(tc)

  End Function

End Class
