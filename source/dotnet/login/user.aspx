<%@ Page Language="VB" Inherits="User" Src="User.vb" Description="Create Account Page"%>

<%@ Register TagPrefix="Portal" TagName="PageHeader" Src="include/PageHeader.ascx" %>
<%@ Register TagPrefix="Portal" TagName="PageSubHeader" Src="include/PageSubHeader.ascx" %>

<html>

<body bgcolor="ffffff" style="margin:0,0,0,0">
<form method=post runat=server>

<Portal:PageHeader ShowSignOut="false" runat="server"/>
<Portal:PageSubHeader Title="Create New Account" runat="server"/>

<table border=0 width="100%" cellspacing=0 cellpadding=0 style="padding:0,0,0,0">
    <tr align=left>
     <td style="padding-left:15">
        &nbsp;<br>
       <font face=geneva,arial size=-1>
           Please complete the following required fields to create a new user account.
           When you are finished, click the "Create Account" button to proceed.<p>
       </font>
     </td>
   </tr>
</table>

<center>

    <!-- sign-in -->
    <table border=0 width=600>
    <tr><td colspan=3>
        <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tr><td>
            <font face=geneva,arial size=-1><b>Sign-In Information</b><i> (Required)</i></font>
        </td></tr>
        </table>
    </td></tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>User Id:</font>
      </td>
      <td>
        <asp:TextBox id=userid width=200px maxlength=20 runat=server />
      </td>
      <td>
        <asp:RequiredFieldValidator id="useridReqVal"
            ControlToValidate="userid"
            Display="Dynamic"
            Font-Name="Verdana" Font-Size="12"
            runat=server>
            *
        </asp:RequiredFieldValidator>
        <asp:Label runat=server id="Err" Text="Please select a different User Id " EnableViewState="false" visible="false"
           style="font-color:black;font:8pt verdana,arial;color:red" />
      </td>
    </tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>Password:</font>
      </td>
      <td>
        <asp:TextBox id=passwd TextMode="Password" maxlength=20 runat=server/>
      </td>
      <td>
          <asp:RequiredFieldValidator id="passwdReqVal"
              ControlToValidate="passwd" ErrorMessage="Password"
              Display="Dynamic"
              Font-Name="Verdana" Font-Size="12"
              runat=server>
              *
          </asp:RequiredFieldValidator>
        <asp:CompareValidator id="CompareValidator1"
            ControlToValidate="passwd2" ControlToCompare="passwd"
            ErrorMessage="Re-enter Password"
            Display="Dynamic"
            Font-Name="Arial" Font-Size="11"
            runat=server>
            Password fields don't match
        </asp:CompareValidator>
      </td>
    </tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>Re-enter Password:</font>
      </td>
      <td>
        <asp:TextBox id=passwd2 TextMode="Password" maxlength=20 runat=server/>
      </td>
      <td>
        <asp:RequiredFieldValidator id="passwd2ReqVal"
            ControlToValidate="passwd2" ErrorMessage="Password"
            Display="Dynamic"
            Font-Name="Verdana" Font-Size="12"
            runat=server>
            *
        </asp:RequiredFieldValidator>
      </td>
    </tr>
    <tr><td colspan=3>&nbsp;</td></tr>

    <!-- personalization information -->
    <tr><td colspan=3>
        <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tr><td><font face=geneva,arial size=-1>
            <b>Personal Information</b><i> (Optional)</i></font>
        </td></tr>
        </table>
    </td></tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>First Name:</font>
      </td>
      <td>
        <asp:TextBox id=fn maxlength=20 width=200px runat=server />
      </td>
      <td>
      </td>
    </tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>Last Name:</font>
      </td>
      <td>
        <asp:TextBox id=ln maxlength=40 width=200px runat=server />
      </td>
      <td>
      </td>
    </tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>Address:</font>
      </td>
      <td>
        <asp:TextBox id=address width=200px runat=server />
      </td>
      <td>
      </td>
    </tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>City:</font>
      </td>
      <td>
        <asp:TextBox id=city width=200px runat=server />
      </td>
      <td>
      </td>
    </tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>State:</font>
      </td>
      <td>
        <asp:TextBox id=adrstate width=30px maxlength=2 runat=server />&nbsp;
        <font face=Arial size=2>Zip Code:</font>&nbsp;
        <ASP:TextBox id=zip width=60px maxlength=5 runat=server />
      </td>
      <td>
        <asp:RegularExpressionValidator id="RegularExpressionValidator1"
            ASPClass="RegularExpressionValidator" ControlToValidate="zip"
            ValidationExpression="[0-9]{5}"
            Display="Static"
            Font-Name="Arial" Font-Size="11"
            runat=server>
            Zip code must be 5 numeric digits
        </asp:RegularExpressionValidator>
      </td>
    </tr>
    <tr>
      <td align=right>
        <font face=Arial size=2>Phone:</font>
      </td>
      <td>
        <asp:TextBox id="phone" maxlength=20 runat="server" />
      </td>
    </tr>
    </table>

    <br>

    <table width="100%" >
        <tr bgcolor="<%= UserState("SubheadColor")%>">
            <td align="center" style="padding:15,15,15,15;border-color:black;border-style:solid;border-width:1;border-right:0;border-left:0">
                <input type="submit" value="Create Account" OnServerClick="Create_User" runat="server"/>
                <input type="submit" value="Cancel" OnServerClick="Cancel_Click" runat="server"/>&nbsp;
            </td>
        </tr>
    </table>

    <p>

</center>
</form>
</body>
</html>
