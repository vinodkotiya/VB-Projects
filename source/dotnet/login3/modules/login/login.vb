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

Public Class Login : Inherits PortalModuleControl

  Public UserName as TextBox
  Public Password as HtmlInputText
  Public ErrorMsg as HtmlContainerControl

  Public Function Authenticate(user As String, pass As String) as Boolean

    Dim authenticated As Boolean = False

    Try
      Dim hshTable as NameValueCollection = CType(Context.GetConfig("system.web/dsnstore"), NameValueCollection)
      Dim dsn as String = CType(hshTable.Item("portaldb"), String)

      Dim myCommand as SqlCommand = new SqlCommand()
      myCommand.Connection = new SqlConnection(dsn)
      myCommand.Connection.Open()
      myCommand.CommandText = "sp_ValidateUser"
      myCommand.CommandType = CommandType.StoredProcedure

      Dim myUserId as SQLParameter = new SQLParameter("@UserId", SqlDbType.NVarChar, 20)
      myUserId.Value = user.Trim()
      myCommand.Parameters.Add(myUserId)

      Dim myPassword as SQLParameter = new SQLParameter("@Password",SqlDbType.NVarChar, 15)
      myPassword.Value = pass.Trim()
      myCommand.Parameters.Add(myPassword)

      Dim IsValid as SQLParameter = new SQLParameter("@IsValid",SqlDbType.Int)
      IsValid.Direction = ParameterDirection.Output
      myCommand.Parameters.Add(IsValid)
      myCommand.ExecuteNonQuery()

      if (CType(IsValid.Value, Integer) = 1) then authenticated = True

    Catch e As Exception
    End Try

    Authenticate = authenticated
    End Function

    Protected Sub SubmitBtn_Click(sender As Object, e As EventArgs)

       If (Authenticate(UserName.Text, Password.Value)) = True Then
           System.Web.Security.FormsAuthentication.SetAuthCookie(UserName.Text, True)
           Response.Redirect("default.aspx")
       Else
           ErrorMsg.Visible = True
       End If

    End Sub

End Class
