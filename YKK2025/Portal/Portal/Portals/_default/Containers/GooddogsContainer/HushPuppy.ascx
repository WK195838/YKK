<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="VISIBILITY" Src="~/Admin/Containers/Visibility.ascx" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>container_01.jpg" WIDTH=14 HEIGHT=36></TD>
		<TD BACKGROUND="<%= SkinPath %>container_02.jpg" HEIGHT=36 ALIGN=LEFT VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR HEIGHT=6><TD COLSPAN=3></TD></TR>
				<TR><TD VALIGN=TOP><dnn:VISIBILITY runat="server" id="dnnVISIBILITY" /></TD><TD><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD><TD ALIGN=LEFT VALIGN=MIDDLE><dnn:TITLE runat="server" id="dnnTITLE" /></TD></TR>
			</TABLE>
		</TD>
		<TD BACKGROUND="<%= SkinPath %>container_03.jpg" WIDTH=17 HEIGHT=36 ALIGN=LEFT VALIGN=MIDDLE></TD>
	</TR>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>container_04.jpg" WIDTH=14></TD>
		<TD BACKGROUND="<%= SkinPath %>container_05.jpg" ALIGN=LEFT VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR><TD id="ContentPane" runat="server" ALIGN=LEFT VALIGN=TOP></TD></TR>
			</TABLE>
		</TD>
		<TD BACKGROUND="<%= SkinPath %>container_06.jpg" WIDTH=17></TD>
	</TR>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>container_07.jpg" WIDTH=14 HEIGHT=28 ALT=""></TD>
		<TD BACKGROUND="<%= SkinPath %>container_08.jpg" HEIGHT=28></TD>
		<TD BACKGROUND="<%= SkinPath %>container_09.jpg" WIDTH=17 HEIGHT=28 ALT=""></TD>
	</TR>
</TABLE>

