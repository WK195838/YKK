<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Control language="vb" CodeBehind="EditContacts.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Contacts.EditContacts" %>
<table width="560" cellspacing="0" cellpadding="0" border="0" summary="Edit Contacts Design Table">
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plNameField" runat="server" controlname="NameField" suffix=":"></dnn:label></td>
		<td align="left" width="400">
			<asp:textbox id="NameField" cssclass="NormalTextBox" width="390" columns="30" maxlength="50"
				runat="server" />
			<asp:requiredfieldvalidator resourcekey="valName.ErrorMessage" cssclass="NormalRed" display="Dynamic" runat="server"
				errormessage="You Must Enter a Valid Name" controltovalidate="NameField" id="valName" />
		</td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plRoleField" runat="server" controlname="RoleField" suffix=":"></dnn:label></td>
		<td><asp:textbox id="RoleField" cssclass="NormalTextBox" width="390" columns="30" maxlength="100"
				runat="server" /></td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plEmailField" runat="server" controlname="EmailField" suffix=":"></dnn:label></td>
		<td><asp:textbox id="EmailField" cssclass="NormalTextBox" width="390" columns="30" maxlength="100"
				runat="server" />
			<asp:regularexpressionvalidator id="valEmail" runat="server" cssclass="NormalRed" controltovalidate="EmailField" errormessage="You Must Enter a Valid Email"
				display="Dynamic" resourcekey="valEmail.ErrorMessage" validationexpression="[\w\.-]+(\+[\w-]*)?@([\w-]+\.)+[\w-]+"></asp:regularexpressionvalidator></td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plContact1Field" runat="server" controlname="Contact1Field" suffix=":"></dnn:label></td>
		<td><asp:textbox id="Contact1Field" cssclass="NormalTextBox" width="390" columns="30" maxlength="250"
				runat="server" /></td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plContact2Field" runat="server" controlname="Contact2Field" suffix=":"></dnn:label></td>
		<td><asp:textbox id="Contact2Field" cssclass="NormalTextBox" width="390" columns="30" maxlength="250"
				runat="server" /></td>
	</tr>
</table>
<p>
	<asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" runat="server" cssclass="CommandButton" text="Update"
		borderstyle="none"></asp:linkbutton>&nbsp;
	<asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" cssclass="CommandButton" text="Cancel"
		borderstyle="none" causesvalidation="False"></asp:linkbutton>&nbsp;
	<asp:linkbutton id="cmdDelete" resourcekey="cmdDelete" runat="server" cssclass="CommandButton" text="Delete"
		borderstyle="none" causesvalidation="False"></asp:linkbutton>
</p>
<portal:audit id="ctlAudit" runat="server" />
