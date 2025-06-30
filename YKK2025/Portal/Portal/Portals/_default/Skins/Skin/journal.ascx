<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="HELP" Src="~/Admin/Skins/Help.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<%@ Register TagPrefix="dnn" TagName="PRIVACY" Src="~/Admin/Skins/Privacy.ascx" %>
<HTML>
<HEAD>
<TITLE>skin.1</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
</HEAD>
<BODY BGCOLOR=#6B6C6F LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<!-- ImageReady Slices (skin.1.psd) -->
<TABLE WIDTH=800 HEIGHT="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>

	<TR HEIGHT=23>
		<TD BACKGROUND="<%= SkinPath %>journal_01.jpg" COLSPAN=4 ALIGN=LEFT VALIGN=MIDDLE>&nbsp;<dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="SelectedTab" /></TD>
		<TD ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=23 ALT=""></TD>
	</TR>

	<TR HEIGHT=20>
		<TD BACKGROUND="<%= SkinPath %>journal_02.jpg" COLSPAN=4 ALIGN=LEFT VALIGN=MIDDLE>&nbsp;<dnn:LOGIN runat="server" id="dnnLOGIN" CssClass="SelectedTab" />&nbsp;&nbsp;<dnn:USER runat="server" id="dnnUSER" CssClass="SelectedTab" />&nbsp;&nbsp;<dnn:HELP runat="server" id="dnnHELP" CssClass="SelectedTab" /></TD>
		<TD ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=20 ALT=""></TD>
	</TR>

	<TR HEIGHT=88>
		<TD BACKGROUND="<%= SkinPath %>journal_03.jpg" COLSPAN=4 ALIGN=LEFT VALIGN=MIDDLE>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<dnn:LOGO runat="server" id="dnnLOGO" /></TD>
		<TD><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=88 ALT=""></TD>
	</TR>

	<TR HEIGHT=39>
		<TD BACKGROUND="<%= SkinPath %>journal_10.jpg" ROWSPAN=3 ALIGN=LEFT VALIGN=TOP HEIGHT=101><IMG SRC="<%= SkinPath %>journal_04.jpg" WIDTH=30 HEIGHT=101 ALT=""></TD>
		<TD COLSPAN=3 ALIGN=LEFT VALIGN=TOP HEIGHT=39><IMG SRC="<%= SkinPath %>journal_05.jpg" WIDTH=770 HEIGHT=39 ALT=""></TD>
		<TD ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=39 ALT=""></TD>
	</TR>

	<TR HEIGHT=34>
		<TD ROWSPAN=2 ALIGN=LEFT VALIGN=TOP BGCOLOR=WHITE><IMG SRC="<%= SkinPath %>journal_06.jpg" WIDTH=166 HEIGHT=62 ALT=""></TD>
		<TD BACKGROUND="<%= SkinPath %>journal_12.jpg" ROWSPAN=3 ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>journal_07.jpg" WIDTH=49 HEIGHT=140 ALT=""></TD>
		<TD ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>journal_08.jpg" WIDTH=555 HEIGHT=34 ALT=""></TD>
		<TD><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=34 ALT=""></TD>
	</TR>

	<TR HEIGHT=28>

		<TD BACKGROUND="<%= SkinPath %>journal_09.jpg" ALIGN=LEFT VALIGN=TOP ROWSPAN=3>

			<TABLE WIDTH="100%" HEIGHT="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR WIDTH="100%">
					<TD ID="ContentPane" RUNAT=SERVER  ALIGN=LEFT VALIGN=TOP>
					</TD>
				</TR>
			</TABLE>

		</TD>
		<TD><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=28 ALT=""></TD>
	</TR>

	<TR HEIGHT=78>
		<TD BACKGROUND="<%= SkinPath %>journal_10.jpg" ROWSPAN=3 ALIGN=LEFT VALIGN=TOP></TD>
		<TD BACKGROUND="<%= SkinPath %>journal_11.jpg" ROWSPAN=3 ALIGN=LEFT VALIGN=TOP><dnn:MENU runat="server" id="dnnMENU" Display="Vertical" BackColor="white" /></TD>
		<TD ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=78 ALT=""></TD>
	</TR>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>journal_12.jpg" ROWSPAN=2 ALIGN=LEFT VALIGN=TOP></TD>
		<TD ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=257 ALT=""></TD>
	</TR>
	<TR VALIGN=BOTTOM HEIGHT=33>
		<TD BACKGROUND="<%= SkinPath %>journal_13.jpg" ALIGN=CENTER VALIGN=MIDDLE><dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="OtherTabs" />&nbsp;&nbsp;<dnn:TERMS runat="server" id="dnnTERMS" CssClass="OtherTabs" />&nbsp;&nbsp;<dnn:PRIVACY runat="server" id="dnnPRIVACY" CssClass="OtherTabs" /></TD>
		<TD ALIGN=LEFT VALIGN=TOP><IMG SRC="<%= SkinPath %>images/spacer.jpg" WIDTH=1 HEIGHT=33 ALT=""></TD>
	</TR>
</TABLE>
<!-- End ImageReady Slices -->
</BODY>
</HTML>
