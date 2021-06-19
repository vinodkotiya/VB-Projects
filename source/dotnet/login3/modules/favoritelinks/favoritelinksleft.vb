Imports System
Imports System.Collections
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports PortalModuleControl
Imports Microsoft.VisualBasic

Public Class FavoriteLinksLeft : Inherits PortalModuleControl

	Public mySpan as HtmlContainerControl

    Protected Sub Page_Load(sender As Object, e As EventArgs)

		Dim dl As ArrayList = New ArrayList()
		Dim links As String = UserState("FavoriteLinksLeft_List")
		Dim i As Integer
 
		If (links <> "") Then
			Dim s As String
			Dim linkList() As String = Split(links, ",")

			For i=0 To linkList.Length-1 Step 2
				If (String.Compare(linkList(i),"CATEGORY") = 0) Then 
				  s = "<b><u><p>" + linkList(i+1) + "</b></u><p>"
				Else
				  s = "<img src='/Quickstart/aspplus/samples/portal/VB/images/bullet.gif' align='middle'> <a target='_new' href='" + linkList(i+1) + "'>" + linkList(i) + "</a><br>"
				End If
				dl.Add(s)		
			Next i
		End If

		Dim innerHtml As String = ""

		for i=0 To dl.Count-1
			innerHtml += dl(i).ToString()
		Next i

		innerHtml += ""
		mySpan.InnerHtml = innerHtml

	End Sub

End Class
