<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<%@ Register TagPrefix="dnn" TagName="PRIVACY" Src="~/Admin/Skins/Privacy.ascx" %>
<TABLE WIDTH="100%" HEIGHT="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>

	<!-- DATE/LOGIN ROW -->
	<TR WIDTH="100%" BGCOLOR="#C2B789">
		<TD ALIGN=LEFT VALIGN=MIDDLE STYLE="border-bottom: black 1px solid;">
			&nbsp;&nbsp;<dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="SelectedTab" />
		</TD>
		<TD ALIGN=RIGHT VALIGN=MIDDLE STYLE="border-bottom: black 1px solid;">
			<dnn:LOGIN runat="server" id="dnnLOGIN" CssClass="SelectedTab" />&nbsp;&nbsp;
		</TD>
	</TR>
	<TR WIDTH="100%" BGCOLOR="white">
		<TD ALIGN=LEFT VALIGN=MIDDLE COLSPAN=2 STYLE="border-bottom: black 1px solid;"><IMG SRC="<%= SkinPath %>spacer.gif" HEIGHT=16 BORDER=0></TD>
	</TR>

	<!-- LOGO ROW -->
	<TR WIDTH="100%" BGCOLOR="white">
		<TD ALIGN=CENTER VALIGN=MIDDLE COLSPAN=2>
			<dnn:LOGO runat="server" id="dnnLOGO" />
		</TD>
	</TR>

	<!-- MENU/USER ROW -->
	<TR WIDTH="100%" BGCOLOR="white">
		<TD ALIGN=LEFT VALIGN=MIDDLE COLSPAN=2 STYLE="border-top: black 1px solid;"><IMG SRC="<%= SkinPath %>spacer.gif" HEIGHT=16 BORDER=0></TD>
	</TR>
	<TR WIDTH="100%" BGCOLOR="#C2B789">
		<TD ALIGN=LEFT VALIGN=MIDDLE STYLE="border-bottom: black 1px solid; border-top: black 1px solid;">
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD WIDTH=6></TD><TD><dnn:MENU runat="server" id="dnnMENU" /></TD></TR></TABLE>
		</TD>
		<TD ALIGN=RIGHT VALIGN=MIDDLE STYLE="border-bottom: black 1px solid; border-top: black 1px solid;">
			<dnn:USER runat="server" id="dnnUSER" CssClass="SelectedTab" />&nbsp;&nbsp;
		</TD>
	</TR>

	<!-- CONTENT ROW -->
	<TR WIDTH="100%">
		<TD ALIGN=LEFT VALIGN=TOP WIDTH="100%" HEIGHT="100%" BGCOLOR=WHITE COLSPAN=2>

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
		<TD ALIGN=LEFT VALIGN=MIDDLE COLSPAN=2 STYLE="border-top: black 1px solid;"><IMG SRC="<%= SkinPath %>spacer.gif" HEIGHT=16 BORDER=0></TD>
	</TR>
	<TR WIDTH="100%" BGCOLOR="#C2B789">
		<TD ALIGN=LEFT VALIGN=MIDDLE WIDTH="100%" COLSPAN=2 STYLE="border-top: black 1px solid; border-bottom: black 1px solid;">
			<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0 VALIGN=MIDDLE>
				<TR>
					<TD WIDTH="100%">
						<TABLE BORDER=0 WIDTH="100%" CELLPADDING=0 CELLSPACING=0>
							<TR>
								<TD COLSPAN=2></TD>
							</TR>
							<TR>
								<TD ALIGN=LEFT VALIGN=BOTTOM>&nbsp;&nbsp;<dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="OtherTabs" /></TD>
								<TD ALIGN=RIGHT VALIGN=BOTTOM><dnn:TERMS runat="server" id="dnnTERMS" />&nbsp;&nbsp;<dnn:PRIVACY runat="server" id="dnnPRIVACY" />&nbsp;&nbsp;</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

</TABLE>
