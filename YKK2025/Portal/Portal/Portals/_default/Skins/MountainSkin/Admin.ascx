<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="HOSTNAME" Src="~/Admin/Skins/HostName.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<!--------------------------------------Mountain Admin Skin--------------------------------------->
<!-----------------------------------Created By John D. Cooper------------------------------------>
<!--------------------------------------www.johndcooper.com--------------------------------------->
<TABLE  WIDTH="100%"  HEIGHT="100%"  BORDER="0" CELLPADDING="0"  CELLSPACING="0">
	<TR>
		<TD CLASS="LogoPane" COLSPAN="2" HEIGHT="0"><dnn:LOGO runat="server" id="dnnLOGO" /></TD>
	</TR>
	<TR>
		
		
    <TD BACKGROUND="<%= SkinPath %>TopBar.gif" HEIGHT="30" VALIGN="MIDDLE" ALIGN="Right" CLASS="OtherTabs"><dnn:LOGIN runat="server" id="dnnLOGIN" />&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:USER runat="server" id="dnnUSER" /></TD>
		
    <TD BACKGROUND="<%= SkinPath %>TopBar.gif" HEIGHT="30" WIDTH="20%" VALIGN="MIDDLE" ALIGN="LEFT" CLASS="OtherTabs">&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="OtherTabs" /></TD>
	</TR>

	<TR>
		
		<TD colspan="2" ID="ContentPane" RUNAT="server"  CLASS="ContentPane" width="100%" HEIGHT="100%" VALIGN="TOP"></TD>
					
	</TR>
	<TR>
		<TD BGCOLOR="#63759C" COLSPAN="2" VALIGN="MIDDLE" ALIGN="CENTER" CLASS="OtherTabs"><dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="OtherTabs" />&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:HOSTNAME runat="server" id="dnnHOSTNAME" />&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;<dnn:TERMS runat="server" id="dnnTERMS" /></TD>
	</TR>
</TABLE>
