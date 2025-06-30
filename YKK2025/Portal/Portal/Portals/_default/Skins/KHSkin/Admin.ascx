<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BREADCRUMB" Src="~/Admin/Skins/BreadCrumb.ascx" %>
<%@ Register TagPrefix="dnn" TagName="PRIVACY" Src="~/Admin/Skins/Privacy.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" WIDTH="95%" ALIGN="CENTER">
	<TR>
    	<TD CLASS="LogoPane"><dnn:LOGO runat="server" id="dnnLOGO" /></TD>
		
    <TD HEIGHT="83" CLASS="HeadPane" WIDTH="500" ALIGN="RIGHT" VALIGN="Bottom"><dnn:LOGIN runat="server" id="dnnLOGIN" CssClass="Login" /> | <dnn:USER runat="server" id="dnnUSER" CssClass="Login" /><BR><dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="Login" /></TD>
	</TR>
	<TR>
	   	<TD CLASS="MenuPane" HEIGHT="23"><dnn:MENU runat="server" id="dnnMENU" /></TD>
    	
    <TD VALIGN="TOP" ALIGN="Left" CLASS="BreadcrumbPane"><dnn:BREADCRUMB runat="server" id="dnnBREADCRUMB" CssClass="Breadcrumb" Separator=" ~" RootLevel="1" />&nbsp;</TD>
	</TR>
	<TR>
		<TD COLSPAN="2">
			<TABLE CELLPADDING="0" CELLSPACING="0" CLASS="ContentTable"  HEIGHT=100% WIDTH="100%" ALIGN="CENTER">
				<TR>
					<TD CLASS="ContentPane" VALIGN="TOP" ID="ContentPane" RUNAT="server"></TD>		
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
	   	
    <TD CLASS="BottomLeft" HEIGHT="23" ALIGN="RIGHT"><dnn:PRIVACY runat="server" id="dnnPRIVACY" /></TD>
    	
    <TD CLASS="BottomRight"><dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="Copyright" /></TD>
	</TR>
</TABLE>
