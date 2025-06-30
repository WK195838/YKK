<%@ Register TagPrefix="dnn" TagName="User" Src="~/controls/User.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Address" Src="~/controls/Address.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Control language="vb" CodeBehind="ManageUsers.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.Users.ManageUsers" %>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<table cellspacing="1" cellpadding="0" summary="Manage Users Design Table">
	<tr>
		<td class="SubHead" width="350" valign="top" colspan="2"><dnn:user id="userControl" runat="server" /></td>
		<td rowspan="4">&nbsp;&nbsp;</td>
		<td rowspan="4" class="NormalBold" valign="top" width="350"><dnn:address runat="server" id="addressUser" /></td>
	</tr>
	<tr id="FillerRow1" runat="server">
		<td>&nbsp;</td>
	</tr>
	<tr id="FillerRow2" runat="server">
		<td>&nbsp;</td>
	</tr>
	<tr id="rowAuthorized" runat="server">
		<td class="SubHead" width="100"><dnn:label id="plAuthorized" runat="server" controlname="chkAuthorized"></dnn:label></td>
		<td><asp:checkbox id="chkAuthorized" runat="server" tabindex="30"></asp:checkbox></td>
	</tr>
</table>
<br>
<dnn:sectionhead id="dshPreferences" runat="server" text="Preferences" section="tblPreferences" resourcekey="Preferences"
	includerule="True" cssclass="Head" isexpanded="True"></dnn:sectionhead>
<table id="tblPreferences" runat="server" cellspacing="1" cellpadding="0" width="600" summary="Preferences">
	<tr>
		<td class="SubHead" width="140"><dnn:label id="plLocale" runat="server" controlname="cboLocale" text="Preferred Language:"></dnn:label></td>
		<td class="NormalBold" nowrap>
			<asp:dropdownlist tabindex="31" id="cboLocale" runat="server" cssclass="NormalTextBox" width="300"></asp:dropdownlist></td>
	</tr>
	<tr>
		<td class="SubHead" width="140"><dnn:label id="plTimeZone" runat="server" controlname="cboTimeZone" text="Time Zone:"></dnn:label></td>
		<td class="NormalBold" nowrap>
			<asp:dropdownlist id="cboTimeZone" tabindex="32" runat="server" cssclass="NormalTextBox" width="300"></asp:dropdownlist></td>
	</tr>
</table>
<br>
<asp:panel id="PasswordManagementRow" runat="server">
<dnn:sectionhead id="dshPassword" runat="server" cssclass="Head" text="Reset Password" resourcekey="ResetPassword"
	includerule="True" section="tblPassword"></dnn:sectionhead>
<table id="tblPassword" cellSpacing="0" cellPadding="4" width="600" summary="Password Management"
	border="0" runat="server">
	<tr vAlign="top" height="*">
		<td id="MessageCell" colSpan="2" runat="server"></td>
	</tr>
	<tr>
		<td class="SubHead" width="175">
			<dnn:label id="plNewPassword" runat="server" text="New Password:" controlname="txtNewPassword"></dnn:label></td>
		<td class="NormalBold" noWrap>
			<asp:textbox id="txtNewPassword" tabIndex="33" runat="server" cssclass="NormalTextBox" maxlength="20"
				size="25" textmode="Password"></asp:textbox>&nbsp;*</td>
	</tr>
	<tr>
		<td class="SubHead" width="175">
			<dnn:label id="plNewConfirm" runat="server" text="Confirm New Password:" controlname="txtNewConfirm"></dnn:label></td>
		<td class="NormalBold" noWrap>
			<asp:textbox id="txtNewConfirm" tabIndex="34" runat="server" cssclass="NormalTextBox" maxlength="20"
				size="25" textmode="Password"></asp:textbox>&nbsp;*</td>
	</tr>
	<tr>
		<td colSpan="2">
			<asp:linkbutton class="CommandButton" id="cmdUpdatePassword" runat="server" text="Update Password"
				resourcekey="cmdUpdatePassword"></asp:linkbutton></td>
	</tr>
</table>
</asp:panel>
<br>
<p>
	<asp:linkbutton class="CommandButton" id="cmdUpdate" resourcekey="cmdUpdate" runat="server" text="Update"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdUnlock" resourcekey="cmdUnlock" runat="server" text="Unlock Account"
		CausesValidation="False" Visible="False"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdCancel" resourcekey="cmdCancel" runat="server" text="Cancel"
		causesvalidation="False" borderstyle="none"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdDelete" resourcekey="cmdDelete" runat="server" text="Delete"
		causesvalidation="False"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdManage" resourcekey="cmdManage" runat="server" text="Manage Security Roles"
		causesvalidation="False"></asp:linkbutton>
</p>
<asp:panel id="pnlAudit" runat="server">
	<table summary="Manage Users Design Table" border="0">
		<tr>
			<td class="SubHead" vAlign="middle" width="150">
				<dnn:label id="plCreatedDate" controlname="lblCreatedDate" runat="server" text="Creation Date:"></dnn:label></td>
			<td vAlign="middle">
				<asp:label id="lblCreatedDate" runat="server" cssclass="Normal"></asp:label></td>
		</tr>
		<tr>
			<td class="SubHead" vAlign="middle" width="150">
				<dnn:label id="plLastLoginDate" controlname="lblLastLoginDate" runat="server" text="Last Login Date:"></dnn:label></td>
			<td vAlign="middle">
				<asp:label id="lblLastLoginDate" runat="server" cssclass="Normal"></asp:label></td>
		</tr>
	</table>
</asp:panel>
<p>
	<asp:label id="lblMessage" runat="server" cssclass="NormalRed"></asp:label>
</p>
