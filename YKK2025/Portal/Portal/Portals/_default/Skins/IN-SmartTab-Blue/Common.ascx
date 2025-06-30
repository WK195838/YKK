<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SOLPARTMENU1" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SEARCH" Src="~/Admin/Skins/Search.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BREADCRUMB" Src="~/Admin/Skins/BreadCrumb.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SOLPARTMENU2" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LINKS" Src="~/Admin/Skins/Links.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<%@ Register TagPrefix="dnn" TagName="PRIVACY" Src="~/Admin/Skins/Privacy.ascx" %>
<%@ Register TagPrefix="dnn" TagName="DOTNETNUKE" Src="~/Admin/Skins/DotNetNuke.ascx" %>
<TABLE class="pagemaster" align="center" border="0" cellspacing="0" cellpadding="0">
	<TR>
		<TD valign="top">
			<TABLE class="skinmaster" border="0" cellspacing="0" cellpadding="0">
				<TR>
					<TD valign="top">
						<TABLE class="skinheader" cellSpacing="0" cellPadding="0" border="0">
							<TR>
								<TD class="skinheader_left"><dnn:LOGO runat="server" id="dnnLOGO" /></TD>
								<TD class="skinheader_right">
									<dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="currentdate" /><br>
									<dnn:USER runat="server" id="dnnUSER" CssClass="link_text" />&nbsp;|&nbsp;<dnn:LOGIN runat="server" id="dnnLOGIN" CssClass="link_text" />
								</TD>
							</TR>
							<TR>
								<TD valign="bottom" colspan="2" class="mainmenu_tab">
									<dnn:SOLPARTMENU1 runat="server" id="dnnSOLPARTMENU1" MenuEffectsMouseOverExpand="False" UseRootBreadcrumbArrow="false" UsearRows="false" RootMenuItemLeftHtml="&lt;span&gt;&lt;span class=IN_RootMenuItemLeft&gt;&lt;/span&gt;&lt;span class=IN_RootMenuItem&gt;" RootMenuItemRightHtml="&lt;/span&gt;&lt;span class=IN_RootMenuItemRight&gt;&lt;/span&gt;&lt;/span&gt;" RootMenuItemCssClass="MainMenu_RootMenuItem" RootMenuItemSelectedCssClass="MainMenu_RootMenuItemSelected" RootMenuItemBreadCrumbCssClass="MainMenu_RootMenuItemBreadCrumb" SelectedBorderColor="blue" />
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD>
						<TABLE class="skindetailheader" cellSpacing="0" cellPadding="0" border="0">
							<TR>
								<TD class="skindetailheader_left"></TD>
								<TD class="skindetailheader_right">
									<dnn:SEARCH runat="server" id="dnnSEARCH" CssClass="link_text" />
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD valign="top" height="100%">
						<TABLE class="skindetail" cellSpacing="0" cellPadding="0">
							<TR vAlign="top">
								<TD class="skindetail_left">
									<TABLE cellSpacing="0" cellPadding="3" width="100%" height="100%">
										<TR vAlign="top">
											<TD class="skindetail_left_breadcrumb">
												<dnn:BREADCRUMB runat="server" id="dnnBREADCRUMB" RootLevel="0" CssClass="skindetail_left_breadcrumb_link" Separator="&lt;!----&gt;" />
											</TD>
										</TR>
										<TR vAlign="top">
											<TD>
												<dnn:SOLPARTMENU2 runat="server" id="dnnSOLPARTMENU2" RootOnly="true" UseRootBreadcrumbArrow="false" UsearRows="false" Display="Vertical" Level="child" RootMenuItemCssClass="ChildMenu_RootMenuItem" RootMenuItemSelectedCssClass="ChildMenu_RootMenuItemSelected" />
											</TD>
										</TR>
										<TR vAlign="top" height="100%">
											<TD id="LeftPane" valign="top" runat="server"></TD>
										</TR>
									</TABLE>
								</TD>
								<TD class="skindetail_main">
									<TABLE cellSpacing="0" cellPadding="3" width="100%" height="100%">
										<TR>
											<TD id="TopPane" runat="server" valign="top" colspan="2"></TD>
										</TR>
										<TR vAlign="top" height="100%">
											<TD id="ContentPane" runat="server" valign="top" class="cntentpane"></TD>
											<TD id="RightPane" runat="server" valign="top" class="rightpane"></TD>
										</TR>
										<TR>
											<TD id="BottomPane" runat="server" valign="bottom" colspan="2"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD valign="top">
						<TABLE class="skindetailfooter" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD>
									<dnn:LINKS runat="server" id="dnnLINKS" CssClass="link_text" Separator="&nbsp;|&nbsp;" Level="same" />
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD valign="top">
						<TABLE class="skinfooter" cellSpacing="0" cellPadding="2" width="100%">
							<TR>
								<TD>
									<dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="copyright" /><br>
									<dnn:TERMS runat="server" id="dnnTERMS" CssClass="link_text" />&nbsp;|&nbsp;<dnn:PRIVACY runat="server" id="dnnPRIVACY" CssClass="link_text" /><br>
									<dnn:DOTNETNUKE runat="server" id="dnnDOTNETNUKE" CssClass="link_text" /><br>
									<a href="http://www.dotnetnuke.com/" class="poweredbydnn"></a>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>



