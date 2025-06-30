<%@Control Inherits="DotNetNuke.Skin" AutoEventWireup="false" Explicit="True" %>
<%@ Register TagPrefix="dnn" TagName="Logo" Src="~/admin/Skins/Logo.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Banner" Src="~/admin/Skins/Banner.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Menu" Src="~/admin/Skins/Menu.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Currentdate" Src="~/admin/Skins/Currentdate.ascx"%>
<%@ Register TagPrefix="dnn" TagName="BreadCrumb" Src="~/admin/Skins/BreadCrumb.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Login" Src="~/admin/Skins/Login.ascx"%>
<%@ Register TagPrefix="dnn" TagName="User" Src="~/admin/Skins/User.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Copyright" Src="~/admin/Skins/Copyright.ascx"%>
<%@ Register TagPrefix="dnn" TagName="HostName" Src="~/admin/Skins/HostName.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Terms" Src="~/admin/Skins/Terms.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Privacy" Src="~/admin/Skins/Privacy.ascx"%>
<%@ Register TagPrefix="dnn" TagName="DotNetNuke" Src="~/admin/Skins/DotNetNuke.ascx"%>
<%@ Register TagPrefix="speerio" TagName="RssTicker" Src="~/controls/RssTicker/RssTicker.ascx"%>
<table border=0 cellpadding=0 cellspacing=0>
	<tr>
		<td class="Normal" rowspan="2" width="139" align="center" valign="middle"><img src="images/logo.gif" width="133" height="45" runat="server" ></td>
		<td width="60"><img src="images/Speerio-Connected-Home_02.gif" runat="server" width=60 height=51 /></td>
		<td width="58"><img src="images/Speerio-Connected-Home_03.gif" runat="server" width=58 height=51 /></td>
		<td class="Normal" colspan="3" rowspan="2" bgcolor="#9DBEC8" height="51" align="left" 
valign="middle" style="padding-left:10px;padding-right:10px">
		<speerio:RssTicker Url="http://www.wired.com/news_drop/netcenter/netcenter.rdf" Path="/rss/channel/item" runat="server" />
		<br><br>
		<dnn:BreadCrumb id="dnnBreadCrumb" runat="server" /></span></td>
		<td width="274" background="images/Speerio-Connected-Home_07.gif" runat="server">&nbsp;&nbsp;<dnn:User id="dnnUser" runat="server" /><br>
		&nbsp;&nbsp;<dnn:Login id="dnnLogin" runat="server" />
		</td>
	</tr>
	<tr>
		<td width="60" height="18"><img src="images/Speerio-Connected-Home_09.gif" runat="server" width=60 height=18 /></td>
		<td width="58" height="18"><img src="images/Speerio-Connected-Home_10.gif" runat="server" width=58 height=18 /></td>
		<td width="274"><img src="images/Speerio-Connected-Home_14.gif" runat="server" width=274 height=18 /></td>
	</tr>
	<tr>
		<td><img src="images/Speerio-Connected-Home_15.gif" runat="server" width=139 height=16 /></td>
		<td><img src="images/Speerio-Connected-Home_16.gif" runat="server" width=60 height=16 /></td>
		<td><img src="images/Speerio-Connected-Home_17.gif" runat="server" width=58 height=16 /></td>
		<td><img src="images/Speerio-Connected-Home_18.gif" runat="server" width=42 height=16 /></td>
		<td><img src="images/Speerio-Connected-Home_19.gif" runat="server" width=275 height=16 /></td>
		<td><img src="images/Speerio-Connected-Home_20.gif" runat="server" width="176" height=16 /></td>
		<td width="274" height=16><img src="images/Speerio-Connected-Home_21.gif" runat="server" width=274 height=16 /></td>
	</tr>
	<tr>
		<td bgcolor="#9DBEC8" valign="top" align="left" width=139 height=236>
		&nbsp;<br><br>
		<img src="images/Speerio-Connected-MenuHeader.gif" width="139" height="20" runat="server" /><dnn:Menu id="dnnMenu" runat="server" display="vertical" /><img src="images/Speerio-Connected-MenuFooter.gif" width="139" height="20" runat="server" />
		</td>
		<td><img src="images/Speerio-Connected-Home_23.gif" runat="server" width=60 height=236 /></td>
		<td><img src="images/Speerio-Connected-Home_24.gif" runat="server" width=58 height=236 /></td>
		<td><img src="images/Speerio-Connected-Home_25.gif" runat="server" width=42 height=236 /></td>
		<td><img src="images/Speerio-Connected-Home_26.gif" runat="server" width=275 height=236 /></td>
		<td class="Normal" width="176" height="236" bgcolor="white" align="left" valign="top"><img src="images/marketing.gif" runat="server" border="0" /></td>
		<td id="RightPane" rowspan="3" runat="server" bgcolor="#CAEBEB" width="274" height="700" align="left" valign="top" style="padding:10px">&nbsp;</td>
	</tr>
	<tr>
		<td><img src="images/Speerio-Connected-Home_29.gif" runat="server" width=139 height=55 /></td>
		<td><img src="images/Speerio-Connected-Home_30.gif" runat="server" width=60 height=55 /></td>
		<td><img src="images/Speerio-Connected-Home_31.gif" runat="server" width=58 height=55 /></td>
		<td><img src="images/Speerio-Connected-Home_32.gif" runat="server" width=42 height=55 /></td>
		<td><img src="images/Speerio-Connected-Home_33.gif" runat="server" width=275 height=55 /></td>
		<td><img src="images/Speerio-Connected-Home_34.gif" runat="server" width=176 height=55 /></td>
	</tr>
	<tr>
		<td id="LeftPane" runat="server" colspan="3" bgcolor="#CAEBEB" width="257" height="700" align="left" valign="top" style="padding:10px" />
		<td id="ContentPane" runat="server" colspan="3" bgcolor="white" width="493" height="700" align="left" valign="top" style="padding:10px"></td>
	</tr>
</table>

