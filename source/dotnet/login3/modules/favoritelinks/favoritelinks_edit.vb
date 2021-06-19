Imports System
Imports System.Collections
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports PortalModulePage
Imports Microsoft.VisualBasic

Public Class FavoriteLinksEdit : Inherits PortalModulePage

	Public selectSize As Integer = 5
	Public mySelect As HtmlSelect

	Protected Sub Page_Load(sender As Object, e As EventArgs)

		Dim statefield As String = "FavoriteLinks" + Request.QueryString("side") + "_List" ' <empty> or "Left"

		If (String.Compare(Request.HttpMethod, "Post", true) <> 0) Then
			'populate items in select control
			Dim links As String = UserState(statefield)

			If (links <> "") Then
				Dim linkvalue, linktext As String
				Dim linkList() As String = Split(links, ",")
				Dim i as Integer

				selectSize = CInt(IIf(linkList.Length/2 < 17, linkList.Length/2, 16))

				For i=0 To linkList.Length-1 Step 2
					If (String.Compare(linkList(i),"CATEGORY") = 0) Then
						linkvalue = "CATEGORY," + linkList(i+1)
						linktext = "---" + linkList(i+1) + "---"
					Else
						linkvalue = linkList(i) + "," + linkList(i+1)
						linktext = linkList(i)
					End If
					mySelect.Items.Add(New ListItem(linktext, linkvalue))
				Next i
			End If
		Else
			Dim links As String = Request.Form("mySelect")
			If links Is Nothing Then
			    UserState(statefield) = ""
			Else
			    UserState(statefield) = links
			End If
			Response.Redirect("/Quickstart/aspplus/samples/portal/VB/default.aspx")
		End If

	End Sub

End Class

