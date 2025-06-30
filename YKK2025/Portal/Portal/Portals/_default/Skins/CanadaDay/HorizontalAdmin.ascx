<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU2" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BREADCRUMB" Src="~/Admin/Skins/BreadCrumb.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="HOSTNAME" Src="~/Admin/Skins/HostName.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Privacy" Src="~/admin/Skins/Privacy.ascx"%>
<!---------------------------------Canada Day Admin Horizontal Skin-------------------------------------->
<!-----------------------------------------Created By Joie------------------------------------>
<!------------------------------------http://portals.my-asp.net--------------------------------------->
<table align="center" valign="middle" cellspacing="0" cellpadding="0" border="0" 
		class="skin-border-container"><tr>
	<td class="skin-border-top-left"><img src="spacer.gif" height="0" width="0"></td>
	<td class="skin-border-top-tile"><img src="spacer.gif" height="0" width="0"></td>
	<td class="skin-border-top-right"><img src="spacer.gif" height="0" width="0"></td></tr><tr>
	<td class="skin-border-left-tile"><img src="spacer.gif" height="0" width="0"></td>
	<td align="center" valign="middle" height="100%" width="100%">
<TABLE BGCOLOR="White" BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
	<TR>
		<TD CLASS="Canadaflag" WIDTH="100%" VALIGN="middle"><dnn:LOGO runat="server" id="dnnLOGO" /></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" ALIGN="center" VALIGN="middle"><img src="/Portals/_default/Skins/CanadaDay/flags.gif"></TD>
	</TR>
	<td align="left">
	<table CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
	<TR>
	<TD HEIGHT="16" ALIGN="left"><dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" cssclass="Date" /></td>
	<TD HEIGHT="16" ALIGN="left" valign="top"><dnn:MENU2 runat="server" id="dnnMENU2" cssclass="TopBarlinks" /></td>
	<TD CLASS="TopBarlinks" HEIGHT="16" ALIGN="left" valign="top"><img SRC="/Portals/_default/Skins/CanadaDay/spacer.gif" WIDTH="200" HEIGHT="1"></td>
	<TD HEIGHT="16" ALIGN="left" valign="top"><dnn:LOGIN runat="server" id="dnnLOGIN" cssclass="TopBarlinks" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<dnn:USER runat="server" id="dnnUSER" cssclass="TopBarlinks" /></td>
	<TD CLASS="TopBarlinks" HEIGHT="16" ALIGN="left" valign="top" col="2"><img SRC="/Portals/_default/Skins/CanadaDay/spacer.gif" WIDTH="20" HEIGHT="1"></td>
	</TR>
	</table></td>
	<TR>		
    	<TD ID="ContentPane" RUNAT="server"  CLASS="ContentPane" WIDTH="100%" HEIGHT="100%" VALIGN="TOP"></TD>
	</TR>
	<TR>
		<td>
	<table CLASS="Footergif" HEIGHT="18" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TD VALIGN="MIDDLE" ALIGN="left"><dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="Footer" /></td>
	<td VALIGN="MIDDLE" ALIGN="center"><dnn:TERMS runat="server" id="dnnTERMS" cssclass="Footer" />&nbsp;&nbsp;&nbsp;&nbsp;<dnn:PRIVACY runat="server" id="dnnPRIVACY" cssclass="Footer" /><img SRC="/Portals/_default/Skins/workingon/spacer.gif" WIDTH="220" HEIGHT="1"></td>
	<td VALIGN="MIDDLE" ALIGN="right"><dnn:HOSTNAME runat="server" id="dnnHOSTNAME" cssclass="Footer" /></TD>
	</table></td></tr>
</TABLE>
</td>
	<td class="skin-border-right-tile"><img src="spacer.gif" height="0" width="0"></td></tr><tr>
	<td class="skin-border-bottom-left"><img src="spacer.gif" height="0" width="0"></td>
	<td class="skin-border-bottom-tile"><img src="spacer.gif" height="0" width="0"></td>
	<td class="skin-border-bottom-right"><img src="spacer.gif" height="0" width="0"></td>
	</tr></table>
