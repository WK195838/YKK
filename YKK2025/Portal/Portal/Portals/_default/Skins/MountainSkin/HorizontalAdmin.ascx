<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU2" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BREADCRUMB" Src="~/Admin/Skins/BreadCrumb.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="HOSTNAME" Src="~/Admin/Skins/HostName.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<!---------------------------------Mountain Horizontal Admin Skin--------------------------------->
<!-----------------------------------Created By John D. Cooper------------------------------------>
<!--------------------------------------www.johndcooper.com--------------------------------------->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
	<TR>
		<TD BGCOLOR="#4474E3" HEIGHT="15" ALIGN="Right" CLASS="OtherTabs"><dnn:LOGIN runat="server" id="dnnLOGIN" />&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:USER runat="server" id="dnnUSER" />&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="OtherTabs" /></TD>
	</TR>
	<TR>
		<TD CLASS="HorizLogo" WIDTH="100%" VALIGN="MIDDLE"><IMG SRC="<%= SkinPath %>spacer.gif" HEIGHT="115" WIDTH="1"><dnn:LOGO runat="server" id="dnnLOGO" /></TD>
	</TR>
	<TR>
	    	<TD BGCOLOR="#64769E" HEIGHT="10"><dnn:MENU2 runat="server" id="dnnMENU2" /></TD>
	</TR>
	<TR>
		<TD ALIGN="RIGHT" CLASS="BreadCrumb" COLSPAN="2" HEIGHT="10"><dnn:BREADCRUMB runat="server" id="dnnBREADCRUMB" CssClass="Breadcrumb" Separator=">>" /></TD>
	</TR>
	<TR>
    		<TD ID="ContentPane" RUNAT="server"  CLASS="ContentPane" WIDTH="100%" HEIGHT="100%" VALIGN="TOP"></TD>		
	</TR>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>TopBar.gif" HEIGHT="20" VALIGN="MIDDLE" ALIGN="CENTER" CLASS="OtherTabs"><dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="OtherTabs" />&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:HOSTNAME runat="server" id="dnnHOSTNAME" />&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:TERMS runat="server" id="dnnTERMS" /></TD>
	</TR>

</TABLE>
