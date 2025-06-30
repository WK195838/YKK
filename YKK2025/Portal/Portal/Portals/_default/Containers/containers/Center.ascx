<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="titlebar3">
		<tr>
			<td class="title3"><dnn:TITLE runat="server" id="dnnTITLE" /></td>
			<td class="title3"><dnn:ACTIONS runat="server" id="dnnACTIONS" /></td>
		</tr>
	</table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="box3">
		<tr>
			<td id="ContentPane" class="CenterContainer" runat="server"></td>
		</tr>
	</table>

