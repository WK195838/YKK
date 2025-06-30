<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="URL" Src="~/controls/URLControl.ascx" %>
<%@ Control language="vb" CodeBehind="EditRss.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.News.EditRss" %>
<table cellspacing="0" cellpadding="2" border="0" summary="Edit RSS Design Table">
	<tr>
		<td class="SubHead" width="200"><dnn:label id="plRSS" runat="server" controlname="ctlRSSxml" suffix=":"></dnn:label></td>
		<td><portal:url id="ctlRSSxml" runat="server" width="300" showtabs="False" showurls="True" showfiles="True"
				urltype="F" showlog="False" shownewwindow="False" showtrack="False" required="True" /></td>
	</tr>
	<tr>
		<td class="SubHead" width="200"><dnn:label id="plXSL" runat="server" controlname="ctlRSSxsl" suffix=":"></dnn:label></td>
		<td><portal:url id="ctlRSSxsl" runat="server" width="300" showtabs="False" showurls="True" showfiles="True"
				urltype="F" showlog="False" shownewwindow="False" showtrack="False" required="True" /></td>
	</tr>
</table>
<dnn:sectionhead id="dshBasic" cssclass="Head" runat="server" text="Security Options (optional)"
	section="tblBasic" resourcekey="lblSecurityTitle" includerule="True"></dnn:sectionhead>
<table cellspacing="0" cellpadding="2" border="0" summary="Edit RSS Design Table" id="tblBasic"
	runat="server">
	<tr>
		<td class="SubHead" colspan="2" width="150"><asp:label id="lblSecurity" runat="server" resourcekey="lblSecurity">Account Information</asp:label></td>
	</tr>
	<tr>
		<td class="SubHead"><dnn:label id="lblDomain" runat="server" controlname="txtAccount" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtAccount" runat="server"></asp:TextBox></td>
	</tr>
	<tr>
		<td class="SubHead" width="150"><dnn:label id="lblPassword" runat="server" controlname="txtPassword" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtPassword" runat="server" TextMode="Password"></asp:TextBox></td>
	</tr>
</table>
<P>
	<asp:linkbutton class="CommandButton" id="cmdUpdate" runat="server" resourcekey="cmdUpdate" text="Update"
		borderstyle="none"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdCancel" runat="server" resourcekey="cmdCancel" text="Cancel"
		borderstyle="none" causesvalidation="False"></asp:linkbutton></P>
