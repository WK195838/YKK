<%@ Control Language="vb" AutoEventWireup="false" Codebehind="LanguagePack.ascx.vb" Inherits="DotNetNuke.Services.Localization.LanguagePack" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<table>
	<tr>
		<td class="SubHead" vAlign="middle" width="150"><dnn:label id="lbLocale" text="Resource Locale" controlname="cboLanguage" runat="server"></dnn:label></td>
		<td vAlign="top"><asp:dropdownlist id="cboLanguage" runat="server"></asp:dropdownlist>&nbsp;
			<asp:linkbutton id="cmdCreate" runat="server" resourcekey="cmdCreate" Text="Create" CssClass="CommandButton"></asp:linkbutton>&nbsp;
			<asp:linkbutton id="cmdCancel" runat="server" resourcekey="cmdCancel" Text="Cancel" CssClass="CommandButton"></asp:linkbutton></td>
	</tr>
</table>
<p></p>
<asp:panel id="pnlLogs" runat="server" Visible="False">
	<dnn:sectionhead id="dshBasic" text="Language Pack Log" runat="server" resourcekey="LogTitle" section="divLog"
		includerule="True" cssclass="Head"></dnn:sectionhead>
	<DIV id="divLog" runat="server">
		<asp:Label id="lblLink" Runat="server" CssClass="Normal"></asp:Label>
		<HR>
		<asp:Label id="lblMessage" runat="server"></asp:Label></DIV>
</asp:panel>
