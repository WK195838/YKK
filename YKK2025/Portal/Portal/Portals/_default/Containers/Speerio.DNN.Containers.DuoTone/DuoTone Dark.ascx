<%@ Control AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Container" %>
<%@ Register TagPrefix="dnn" TagName="Icon" Src="~/admin/Containers/Icon.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Actions" Src="~/admin/Containers/Actions.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Title" Src="~/admin/Containers/Title.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Visibility" Src="~/admin/Containers/Visibility.ascx"%>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width=19><img src="images/C1_01.gif" width=19 height=37 runat="server"></td>
		<td bgcolor="#78A2B0" nowrap> <dnn:Actions runat="server" id="dnnActions" /> 
					<dnn:Title runat="server" id="dnnTitle" />
		</td>
		<td width=20><img src="images/C1_03.gif" width=20 height=37 runat="server"></td>
	</tr>
	<tr>
		<td width=19><img src="images/C1_04.gif" width=19 height=12 runat="server"></td>
		<td><img src="images/C1_05.gif" style="width:100%" height=12 runat="server"></td>
		<td width=20><img src="images/C1_06.gif" width=20 height=12 runat="server"></td>
	</tr>
	<tr>
		<td bgcolor="#9DBEC8" width=19></td>
		<td bgcolor="#9DBEC8" id="ContentPane" runat="server"></td>
		<td bgcolor="#9DBEC8" width=20></td>
	</tr>
	<tr>
		<td width=19><img src="images/C1_10.gif" width=19 height=20 runat="server"></td>
		<td><img src="images/C1_11.gif" style="width:100%" height=20 runat="server"></td>
		<td width=20><img src="images/C1_12.gif" width=20 height=20 runat="server"></td>
	</tr>
</table>
