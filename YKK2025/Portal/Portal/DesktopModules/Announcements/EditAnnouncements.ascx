<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Tracking" Src="~/controls/URLTrackingControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="URL" Src="~/controls/URLControl.ascx" %>
<%@ Control language="vb" CodeBehind="EditAnnouncements.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Announcements.EditAnnouncements" %>
<table cellSpacing="0" cellPadding="0" width="600" summary="Edit Announcements Design Table">
	<tr vAlign="top">
		<td class="SubHead" width="150"><dnn:label id="plTitle" runat="server" controlname="txtTitle" suffix=":"></dnn:label></td>
		<td width="450">
			<asp:textbox id="txtTitle" runat="server" maxlength="100" Columns="30" width="200" cssclass="NormalTextBox"></asp:textbox>
			&nbsp;
			<asp:CheckBox ID="chkAddDate" resourcekey="AddDate" Runat="server" CSSClass="SubHead" Text="Add Date?"
				TextAlign="Right"></asp:CheckBox>
			<br>
			<asp:requiredfieldvalidator id="valTitle" resourcekey="Title.ErrorMessage" runat="server" CssClass="NormalRed"
				ControlToValidate="txtTitle" ErrorMessage="You Must Enter A Title For The Announcement" Display="Dynamic"></asp:requiredfieldvalidator>
		</td>
	</tr>
	<tr vAlign="top">
		<td class="SubHead" width="150"><dnn:label id="plDescription" runat="server" controlname="txtDescription" suffix=":"></dnn:label></td>
		<td>
			<dnn:texteditor id="teDescription" runat="server" width="450" height="200"></dnn:texteditor>
			<br>
			<asp:requiredfieldvalidator id="valDescription" resourcekey="Description.ErrorMessage" runat="server" CssClass="NormalRed"
				ControlToValidate="teDescription" ErrorMessage="You Must Enter A Description Of The Announcement" Display="Dynamic"></asp:requiredfieldvalidator>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="SubHead" width="150"><dnn:label id="plURL" runat="server" controlname="ctlURL" suffix=":"></dnn:label></td>
		<td width="325">
			<portal:url id="ctlURL" runat="server" width="225" shownone="true" />
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="SubHead" width="150"><dnn:label id="plExpires" runat="server" controlname="txtExpires" suffix=":"></dnn:label></td>
		<td width="325">
			<asp:textbox id="txtExpires" runat="server" Columns="20" width="225" cssclass="NormalTextBox"
				Text=""></asp:textbox>
			&nbsp;
			<asp:hyperlink id="cmdCalendar" resourcekey="Calendar" CssClass="CommandButton" Runat="server">Calendar</asp:hyperlink>
			<asp:comparevalidator id="valExpires" resourcekey="Expires.ErrorMessage" runat="server" CssClass="NormalRed"
				ControlToValidate="txtExpires" ErrorMessage="<br>You have entered an invalid date!" Display="Dynamic"
				Type="Date" Operator="DataTypeCheck"></asp:comparevalidator>
		</td>
	</tr>
	<tr>
		<td class="SubHead" width="150"><dnn:label id="plViewOrder" runat="server" controlname="txtViewOrder" suffix=":"></dnn:label></td>
		<td>
			<asp:textbox id="txtViewOrder" runat="server" maxlength="3" Columns="20" width="300" CssClass="NormalTextBox"></asp:textbox>
			<asp:comparevalidator id="valViewOrder" resourcekey="ViewOrder.ErrorMessage" runat="server" CssClass="NormalRed"
				ControlToValidate="txtViewOrder" ErrorMessage="<br>View order must be an integer value." Display="Dynamic"
				Type="Integer" Operator="DataTypeCheck"></asp:comparevalidator>
		</td>
	</tr>
</table>
<p>
	<asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" runat="server" CssClass="CommandButton" Text="Update"
		BorderStyle="none"></asp:linkbutton>&nbsp;
	<asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" CssClass="CommandButton" Text="Cancel"
		BorderStyle="none" CausesValidation="False"></asp:linkbutton>&nbsp;
	<asp:linkbutton id="cmdDelete" resourcekey="cmdDelete" runat="server" CssClass="CommandButton" Text="Delete"
		BorderStyle="none" CausesValidation="False"></asp:linkbutton>
</p>
<portal:Audit id="ctlAudit" runat="server" />
<br>
<br>
<portal:Tracking id="ctlTracking" runat="server" />
