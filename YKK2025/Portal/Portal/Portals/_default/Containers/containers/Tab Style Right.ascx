<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="titlebar2">
		<tr>
			<td class="title1"><dnn:TITLE runat="server" id="dnnTITLE" /></td>
			<td class="tab1"><dnn:ACTIONS runat="server" id="dnnACTIONS" /></td>
		</tr>
	</table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="box2">
		<tr>
			<td id="ContentPane" class="RightContainer" runat="server"></td>
		</tr>
	</table>

