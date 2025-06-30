<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="VISIBILITY" Src="~/Admin/Containers/Visibility.ascx" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<table border="0" cellpadding="0" width="100%" cellspacing="0" id="table1">
	<tr>
		<td width="14" height="30" background="<%= SkinPath %>DlgHdrL-bl.gif" style="float: left">&nbsp;</td>
		<TD BACKGROUND="<%= SkinPath %>DlgHdrC-bl.gif" HEIGHT="30" ALIGN="LEFT" VALIGN="TOP">
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0">
				<TR HEIGHT="6">
					<TD COLSPAN="3"></TD>
				</TR>
				<TR>
					<TD VALIGN="TOP"><dnn:VISIBILITY runat="server" id="dnnVISIBILITY" /></TD>
					<TD><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
					<TD ALIGN="LEFT" VALIGN="MIDDLE" STYLE="color: White; font-family: Arial; font-size: 19;"><dnn:TITLE runat="server" id="dnnTITLE" cssclass="dialogtitle" /></TD>
				</TR>
			</TABLE>
		</TD>
		<td background="<%= SkinPath %>DlgHdrR-bl.gif" width="14">&nbsp;</td>
	</tr>
	<tr>
		<td width="4" background="<%= SkinPath %>DlgBrdL.gif">&nbsp;</td>
		<TD ALIGN="LEFT" VALIGN="TOP">
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" ID="Table3">
				<TR>
					<TD id="ContentPane" runat="server" ALIGN="LEFT" VALIGN="TOP" bgcolor="whitesmoke"></TD>
				</TR>
			</TABLE>
		</TD>
		<td background="<%= SkinPath %>DlgBrdL.gif">&nbsp;</td>
	</tr>
	<tr>
		<td width="14" background="<%= SkinPath %>DlgBrdBL.gif" height="22">&nbsp;</td>
		<td background="<%= SkinPath %>DlgBrdBC.gif" bgcolor="whitesmoke">&nbsp;</td>
		<td background="<%= SkinPath %>DlgBrdBR.gif">&nbsp;</td>
	</tr>
</table>

