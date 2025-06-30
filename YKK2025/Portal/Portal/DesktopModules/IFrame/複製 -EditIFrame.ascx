<%@ Control language="vb" CodeBehind="EditIFrame.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.IFrame.EditIFrame" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table width="560" cellspacing="0" cellpadding="0" border="0" summary="Edit IFrame Design Table">
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plSrc" runat="server" controlname="txtSrc" suffix=":"></dnn:label></td>
		<td width="400"><asp:textbox id="txtSrc" cssclass="NormalTextBox" columns="50" runat="server" />
			<asp:regularexpressionvalidator id="valURL" runat="server" cssclass="NormalRed" controltovalidate="txtSrc" errormessage="You Must Enter a Valid URL"
				display="Dynamic" resourcekey="valURL.ErrorMessage" validationexpression="^((http|https)\://)?[a-zA-Z0-9\-\.]+\.[a-zA-Z0-9]{2,3}(:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,\'/\\\+&amp;%\$#\=~])*[^\.\,\)\(\s]$"></asp:regularexpressionvalidator>
		</td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plWidth" runat="server" controlname="txtWidth" suffix=":"></dnn:label></td>
		<td width="400"><asp:textbox id="txtWidth" cssclass="NormalTextBox" columns="50" runat="server" /></td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plHeight" runat="server" controlname="txtHeight" suffix=":"></dnn:label></td>
		<td width="400"><asp:textbox id="txtHeight" cssclass="NormalTextBox" columns="50" runat="server" /></td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plTitle" runat="server" controlname="txtTitle" suffix=":"></dnn:label></td>
		<td width="400">
			<asp:textbox id="txtTitle" cssclass="NormalTextBox" columns="50" runat="server" />
		</td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plScrolling" runat="server" controlname="cboScrolling" suffix=":"></dnn:label></td>
		<td width="400">
			<asp:dropdownlist id="cboScrolling" runat="server" cssclass="NormalTextBox">
				<asp:listitem>auto</asp:listitem>
				<asp:listitem>no</asp:listitem>
				<asp:listitem>yes</asp:listitem>
			</asp:dropdownlist>
		</td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plBorder" runat="server" controlname="cboBorder" suffix=":"></dnn:label></td>
		<td width="400">
			<asp:dropdownlist id="cboBorder" runat="server" cssclass="NormalTextBox">
				<asp:listitem>no</asp:listitem>
				<asp:listitem>yes</asp:listitem>
			</asp:dropdownlist>
		</td>
	</tr>
</table>
<p>
	<asp:linkbutton class="CommandButton" id="cmdUpdate" resourcekey="cmdUpdate" runat="server" borderstyle="none"
		text="Update"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdCancel" resourcekey="cmdCancel" runat="server" borderstyle="none"
		text="Cancel" causesvalidation="False"></asp:linkbutton>
</p>
