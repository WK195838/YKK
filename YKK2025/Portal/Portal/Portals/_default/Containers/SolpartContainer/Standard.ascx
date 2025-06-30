<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="SOLPARTACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<%@ Register TagPrefix="dnn" TagName="VISIBILITY" Src="~/Admin/Containers/Visibility.ascx" %>
<table style="BORDER-COLLAPSE: collapse;" width="100%" border="0" ID="Table1">
	<tr>
		<td><img src="<%= SkinPath %>spacer.gif" width="7px"></td>
		<td>
			<table style="BORDER-COLLAPSE: collapse;" width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td><dnn:SOLPARTACTIONS runat="server" id="dnnSOLPARTACTIONS" /></td>
					<TD WIDTH="100%" STYLE="color: Black; font-family: Arial; font-size: 20;"><dnn:TITLE runat="server" id="dnnTITLE" cssclass="dialogtitle" /></TD>
					<td><dnn:VISIBILITY runat="server" id="dnnVISIBILITY" /></td>
				</tr>
				<TR>
					<td align="center" colspan="4" id="ContentPane" runat="server"></td>
		</td>
	</tr>
	<tr>
		<td><img src="<%= SkinPath %>spacer.gif" width="7px"></td>
	</tr>
</table>
</td>
<td><img src="<%= SkinPath %>spacer.gif" width="7px"></td>
</tr></table>
