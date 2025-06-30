<%@ Control language="vb" Inherits="DotNetNuke.Modules.Admin.Security.Signin" CodeBehind="Signin.ascx.vb" AutoEventWireup="false" Explicit="True" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table cellspacing="0" cellpadding="3" border="0" summary="SignIn Design Table">
	<tr>
		<td colspan="2" class="SubHead" align="center"><dnn:label id="plUsername" controlname="txtUsername" runat="server" text="UserName:"></dnn:label></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><asp:textbox id="txtUsername" columns="9" width="130" cssclass="NormalTextBox" runat="server" /></td>
	</tr>
	<tr>
		<td colspan="2" class="SubHead" align="center"><dnn:label id="plPassword" controlname="txtPassword" runat="server" text="Password:"></dnn:label></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><asp:textbox id="txtPassword" columns="9" width="130" textmode="password" cssclass="NormalTextBox"
				runat="server" /></td>
	</tr>
	<tr id="rowVerification1" runat="server" visible="false">
		<td colspan="2" class="SubHead" align="center"><dnn:label id="plVerification" controlname="txtVerification" runat="server" text="Verification Code:"></dnn:label></td>
	</tr>
	<tr id="rowVerification2" runat="server" visible="false">
		<td colspan="2" align="center"><asp:textbox id="txtVerification" columns="9" width="130" cssclass="NormalTextBox" runat="server" /></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><asp:checkbox id="chkCookie" class="Normal" resourcekey="Remember" text="Remember Login" runat="server" /></td>
	</tr>
	<tr>
		<td align="right" width="50%" colSpan="1" id="TDLogin" runat="server"><asp:Button ID="cmdLogin" resourcekey="cmdLogin" cssclass="StandardButton" text="Login" runat="server"
				Width="80" />&nbsp;</td>
		<td align="left" width="50%" id="TDRegister" runat="server">&nbsp;<asp:Button id="cmdRegister" resourcekey="cmdRegister" cssclass="StandardButton" text="Register"
				runat="server" Width="80" /></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><asp:Button id="cmdSendPassword" resourcekey="cmdSendPassword" cssclass="StandardButton" text="Password Reminder"
				runat="server" Width="140" /></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><asp:label id="lblLogin" cssclass="Normal" runat="server" /></td>
	</tr>
</table>
<br>
