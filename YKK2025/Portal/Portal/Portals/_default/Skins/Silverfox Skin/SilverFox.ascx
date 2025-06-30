<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<%@ Register TagPrefix="dnn" TagName="PRIVACY" Src="~/Admin/Skins/Privacy.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<TABLE WIDTH="100%" HEIGHT="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>

	<!-- DATE/LOGIN ROW -->
	<TR WIDTH="100%" BGCOLOR="#C2B789">
		<TD ALIGN=LEFT  background="<%= SkinPath %>topbar.gif" bordercolor="#000000" bgcolor="#FFFFFF" height="30">
			&nbsp;&nbsp; <dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="SelectedTab" />
		</TD>
		<TD ALIGN=RIGHT  background="<%= SkinPath %>topbar.gif" bordercolor="#000000" bgcolor="#FFFFFF">
			<dnn:USER runat="server" id="dnnUSER" CssClass="SelectedTab" /></TD>
	</TR>

	<!-- LOGO ROW -->
	<TR WIDTH="100%" BGCOLOR="white">
		<TD ALIGN=CENTER COLSPAN=2 background="<%= SkinPath %>banner.jpg" bgcolor="#F2F2F9">
			<p align="left"><dnn:LOGO runat="server" id="dnnLOGO" /></TD>
	</TR>

	<!-- MENU/USER ROW -->
	<TR WIDTH="100%"   BGCOLOR= #FFFFFF>
	
		<TD ALIGN=LEFT  background="<%= SkinPath %>bar.jpg">
			<TABLE width="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0 height="28">
				<TR>
					<td width ="20" background="<%= SkinPath %>menubar.gif"></td>
					<TD WIDTH=9 background ="<%= SkinPath %>menubar.gif"></TD>
					<TD bgcolor="#F3F3F7" background="<%= SkinPath %>menubar.gif" align="center"><dnn:MENU runat="server" id="dnnMENU" /></TD>
					<TD width="9"  align="left" background ="<%= SkinPath %>menubar.gif" ></TD>
				</TR>
			</TABLE>
		</TD>
		<TD ALIGN=RIGHT  background="<%= SkinPath %>menubar.gif" bgcolor="#F3F3F7">
			<dnn:LOGIN runat="server" id="dnnLOGIN" CssClass="SelectedTab" /></TD>
	</TR>

	<!-- CONTENT ROW -->
	<TR WIDTH="100%">
		<TD ALIGN=LEFT VALIGN=TOP WIDTH="100%" HEIGHT="100%" BGCOLOR=WHITE COLSPAN=4>

			<TABLE WIDTH="100%" BORDER=0 CELLPADDING=6 CELLSPACING=0>
				<TR>
					<TD WIDTH="100%" HEIGHT="100%">
						<TABLE WIDTH="100%" HEIGHT="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0 BGCOLOR=WHITE>
							<TR HEIGHT="6px">
								<TD COLSPAN=3></TD>
							</TR>
							<TR>
								<TD id="LeftPane" runat="server" WIDTH=200 visible="false" ALIGN=LEFT VALIGN=TOP></TD>
								<TD id="ContentPane" runat="server" HEIGHT="100%" visible="true" ALIGN=LEFT VALIGN=TOP></TD>
								<TD id="RightPane" runat="server" WIDTH=200 visible="false" ALIGN=LEFT VALIGN=TOP></TD>
							</TR>
							<TR HEIGHT="6px">
								<TD COLSPAN=3></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<!-- FOOTER ROW -->
	<TR WIDTH="100%" BGCOLOR="white">
		<TD ALIGN=LEFT VALIGN=MIDDLE COLSPAN=4 ></TD>
	</TR>
	<TR WIDTH="100%" BGCOLOR="#C2B789">
		<TD ALIGN=LEFT VALIGN=MIDDLE WIDTH="100%" COLSPAN=4 >
			<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0 VALIGN=MIDDLE>
				<TR>
					<TD WIDTH="100%">
						<TABLE BORDER=0 WIDTH="100%" CELLPADDING=0 CELLSPACING=0>
							<TR>
								<TD COLSPAN=2></TD>
							</TR>
							<tr>
								<TD ALIGN=LEFT VALIGN=BOTTOM width="1%" background="<%= SkinPath %>topbar.gif">
								&nbsp;</TD>
								<TD ALIGN=RIGHT width="100%" bgcolor="#FFFFFF" background="<%= SkinPath %>topbar.gif" height="30">
								<p align="center"><dnn:TERMS runat="server" id="dnnTERMS" />&nbsp;&nbsp;<dnn:PRIVACY runat="server" id="dnnPRIVACY" />&nbsp;&nbsp;</TD>
							</tr>
							<TR>
								<TD ALIGN=LEFT VALIGN=BOTTOM bgcolor="#F3F3F7" colspan="2">
								<p align="center">&nbsp;<dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="OtherTabs" /></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

</TABLE>
