<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="URL" Src="~/controls/URLControl.ascx" %>
<%@ Control language="vb" CodeBehind="EditXml.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.XML.EditXml" %>
<br>
<table ellspacing="0" cellpadding="0" border="0" summary="Edit XML Design Table">
	<tr>
		<td class="SubHead"><dnn:label id="plXML" runat="server" controlname="ctlURLxml" suffix=":"></dnn:label></td>
		<td ><portal:url id="ctlURLxml" runat="server" width="300" showtabs="False" showurls="True" showfiles="True" urltype="F" showlog="False" shownewwindow="False" showtrack="False" required="True" /></td>
	</tr>
	<tr>
		<td class="SubHead"><dnn:label id="plXSL" runat="server" controlname="ctlURLxsl" suffix=":"></dnn:label></td>
		<td ><portal:url id="ctlURLxsl" runat="server" width="300" showtabs="False" showurls="True" showfiles="True" urltype="F" showlog="False" shownewwindow="False" showtrack="False" required="True" /></td>
	</tr>
</table>
	<asp:linkbutton class="CommandButton" id="cmdUpdate" resourcekey="cmdUpdate" runat="server" borderstyle="none"
		text="Update"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdCancel" resourcekey="cmdCancel" runat="server" borderstyle="none"
		text="Cancel" causesvalidation="False"></asp:linkbutton>
