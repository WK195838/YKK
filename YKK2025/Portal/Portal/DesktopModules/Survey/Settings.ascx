<%@ Control Inherits="DotNetNuke.Modules.Survey.Settings" CodeBehind="Settings.ascx.vb" language="vb" AutoEventWireup="false" Explicit="true" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table cellSpacing="0" cellPadding="4" border="0">
	<tr>
		<td class="SubHead" width="200"><dnn:label id="plClosingDate" text="Survey Closing Date:" controlname="txtClosingDate" runat="server" /></td>
		<td valign="top">
			<asp:TextBox id="txtClosingDate" Runat="server" CssClass="NormalTextBox"></asp:TextBox>
			<asp:HyperLink id=cmdCalendar resourcekey="Calendar" CssClass="CommandButton" Runat="server">Calendar</asp:HyperLink>
			<asp:CompareValidator id=valClosingDate runat="server" CssClass="NormalRed" resourcekey="valClosingDate.ErrorMessage" ControlToValidate="txtClosingDate" ErrorMessage="<br>Invalid Closing Date!" Operator="DataTypeCheck" Type="Date" Display="Dynamic"></asp:CompareValidator>
		</td>
	</tr>
	<tr>
		<td class="SubHead" width="200"><dnn:label id="plGraphWidth" text="Maximum Bar Graph Width:" controlname="txtGraphWidth" runat="server" /></td>
		<td valign="top"><asp:TextBox id="txtGraphWidth" Runat="server" CssClass="NormalTextBox"></asp:TextBox></td>
	</tr>
	<tr>
		<td class="SubHead" width="200"><dnn:label id="plPersonal" text="Vote Tracking:" controlname="rblstPersonal" runat="server" /></td>
		<td vAlign="top" height="36">
			<asp:RadioButtonList id="rblstPersonal" CssClass="NormalTextBox" runat="server" Width="216px">
				<asp:ListItem resourcekey="No" Value="0" Selected="true">Vote tracking via cookie</asp:ListItem>
				<asp:ListItem resourcekey="Yes" Value="1">1 Vote/Registered User</asp:ListItem>
			</asp:RadioButtonList>
		</td>
	</tr>
	<tr>
		<td class="SubHead" width="200"><dnn:label id="plSurveyResults" text="Survey Results:" controlname="rblstSurveyResults" runat="server" /></td>
		<td vAlign="top" height="36">
			<asp:RadioButtonList id="rblstSurveyResults" CssClass="NormalTextBox" Width="216px" runat="server" repeatdirection="Horizontal">
				<asp:ListItem resourcekey="PublicResults" Value="0" Selected="true">Public</asp:ListItem>
				<asp:ListItem resourcekey="PrivateResults" Value="1">Private</asp:ListItem>
			</asp:RadioButtonList>
		</td>
	</tr>
</table>
