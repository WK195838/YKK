<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<%@ Register TagPrefix="dnn" TagName="VISIBILITY" Src="~/Admin/Containers/Visibility.ascx" %>
<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>journal_01.gif" WIDTH=19 HEIGHT=34 ALT=""><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
		<TD BACKGROUND="<%= SkinPath %>journal_02.gif" HEIGHT=34 ALT=""><dnn:TITLE runat="server" id="dnnTITLE" /></TD>
		<TD BACKGROUND="<%= SkinPath %>journal_03.gif" WIDTH=37 HEIGHT=34 ALT=""><dnn:VISIBILITY runat="server" id="dnnVISIBILITY" /></TD>
		<TD BACKGROUND="<%= SkinPath %>journal_04.gif" WIDTH=15 ROWSPAN=4></TD>
	</TR>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>journal_05.gif" WIDTH=19 ALT=""></TD>
		<TD BACKGROUND="<%= SkinPath %>journal_06.gif" ALT="">
			<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR>
					<TD ID="ContentPane" RUNAT=SERVER></TD>
				</TR>
			</TABLE>
		</TD>
		<TD BACKGROUND="<%= SkinPath %>journal_07.gif" WIDTH=37 ALT=""></TD>
	</TR>
	<TR>
		<TD><IMG SRC="<%= SkinPath %>journal_08.gif" WIDTH=19 HEIGHT=26 ALT=""></TD>
		<TD BACKGROUND="<%= SkinPath %>journal_09.gif" HEIGHT=26 ALT=""></TD>
		<TD><IMG SRC="<%= SkinPath %>journal_10.gif" WIDTH=37 HEIGHT=26 ALT=""></TD>
	</TR>
	<TR HEIGHT=12>
		<TD BACKGROUND="<%= SkinPath %>journal_11.gif" COLSPAN=3></TD>
	</TR>
</TABLE>

