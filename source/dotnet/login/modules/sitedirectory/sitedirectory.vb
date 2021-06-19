Imports System
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Data
Imports System.Data.SqlClient
Imports PortalModuleControl
Imports Microsoft.VisualBasic

Public Class SiteDirectory : Inherits PortalModuleControl

  Public myDataGrid as DataList

    Protected Sub Page_Load(sender As Object, e As EventArgs)

    Dim LinkIndices As String = UserState("SiteDirectory_Links")
    If (LinkIndices = "") Then LinkIndices = "0"

    Dim hshTable as NameValueCollection = CType(Context.GetConfig("system.web/dsnstore"), NameValueCollection)
    Dim dsn as String = CType(hshTable.Item("portaldb"), String)

    Dim myAdapter As SqlDataAdapter = new SqlDataAdapter()
    myAdapter.SelectCommand = New SqlCommand()
    myAdapter.SelectCommand.Connection = New SqlConnection(dsn)
    myAdapter.SelectCommand.CommandText = "select LinkIndex, LinkName, LinkRef, LinkDescription, UserData from SiteDirectory where LinkIndex IN (" + LinkIndices + ")"

    Dim myDataSet As DataSet = New DataSet()
    myAdapter.Fill(myDataSet,"Results")

    Dim rows As DataRowCollection = myDataSet.Tables(0).Rows
    Dim index As Integer = 0
    Dim Source As ArrayList = New ArrayList()
    Dim i,j As Integer

    For i=0 To rows.Count-1
      Dim propertyBag As Hashtable = new Hashtable()

      For j=0 to myDataSet.Tables(0).Columns.Count-1
        Dim value As Object = rows(i).Item(myDataSet.Tables(0).Columns(j))
                if (String.Compare(myDataSet.Tables(0).Columns(j).ToString(), "LinkRef") = 0) Then
          propertyBag(myDataSet.Tables(0).Columns(j).ToString()) = IIf(value Is Nothing, "" , Request.ApplicationPath + "/" + value.ToString())
        Else
          propertyBag(myDataSet.Tables(0).Columns(j).ToString()) = IIf(value Is Nothing, "" , value.ToString())
        End If
      Next j

      Source.Add (propertyBag)
    Next i

    myDataGrid.DataSource = Source
    DataBind()

    End Sub

End Class
