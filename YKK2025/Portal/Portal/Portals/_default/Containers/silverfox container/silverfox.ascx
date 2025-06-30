<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="VISIBILITY" Src="~/Admin/Containers/Visibility.ascx" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<table border="0" cellpadding="0" width="100%" cellspacing="0" id="table1">
	<tr>
		<td width="14" height="46" background="<%= SkinPath %>topleft.gif" style="float: left">&nbsp;</td>
		<TD BACKGROUND="<%= SkinPath %>topmid.gif" HEIGHT=46 ALIGN=LEFT VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR HEIGHT=6><TD COLSPAN=3></TD></TR>
				<TR><TD VALIGN=TOP><dnn:VISIBILITY runat="server" id="dnnVISIBILITY" /></TD><TD><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD><TD ALIGN=LEFT VALIGN=MIDDLE><dnn:TITLE runat="server" id="dnnTITLE" /></TD></TR>
			</TABLE>
		</TD>
		<td background="<%= SkinPath %>topright.gif" width="14">&nbsp;</td>
	</tr>
	<tr>
		<td width="14" background="<%= SkinPath %>left.gif">&nbsp;</td>
		<TD BACKGROUND="<%= SkinPath %>container_05.jpg" ALIGN=LEFT VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR><TD id="ContentPane" runat="server" ALIGN=LEFT VALIGN=TOP></TD></TR>
			</TABLE>
		</TD>

		<td background="<%= SkinPath %>right.gif">&nbsp;</td>
	</tr>
	<tr>
		<td width="14" background="<%= SkinPath %>bLeft.gif" height="24">&nbsp;</td>
		<td background="<%= SkinPath %>bMid.gif">&nbsp;</td>
		<td background="<%= SkinPath %>bright.gif">&nbsp;</td>
	</tr>
</table>
