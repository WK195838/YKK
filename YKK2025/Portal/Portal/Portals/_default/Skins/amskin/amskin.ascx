<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BANNER" Src="~/Admin/Skins/Banner.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SOLPARTMENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BREADCRUMB" Src="~/Admin/Skins/BreadCrumb.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LINKS1" Src="~/Admin/Skins/Links.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<html>
<head>
<link href="<%= SkinPath %>amskin.css" rel="stylesheet" type="text/css">
</head>
<body>
<script language="JavaScript">
window.moveTo(0,0);
if (document.all) {
top.window.resizeTo(screen.availWidth,screen.availHeight);
}
else if (document.layers||document.getElementById) {
if (top.window.outerHeight<screen.availHeight||top.window.outerWidth<screen.availWidth){
top.window.outerHeight = screen.availHeight;
top.window.outerWidth = screen.availWidth;
}
}
</script>
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="110" height="27" align="left" valign="top" background="<%= SkinPath %>amskin_r1_c1.gif" class="Footer"><dnn:LOGO runat="server" id="dnnLOGO" /></td>
    <td width="100%" height="27" align="center" valign="middle" background="<%= SkinPath %>amskin_r1_c2.gif" class="Footer"><dnn:BANNER runat="server" id="dnnBANNER" /></td>
    <td width="200" height="27" align="right" valign="middle" nowrap background="<%= SkinPath %>amskin_r1_c2.gif" class="Footer"><div align="right"><b><dnn:USER runat="server" id="dnnUSER" CssClass="SelectedTab" />&nbsp; |&nbsp; <dnn:LOGIN runat="server" id="dnnLOGIN" CssClass="SelectedTab" LogoffText="Sair" /></b></div></td>
  </tr>
  <tr>
    <td height="33" colspan="2" align="left" valign="middle" nowrap background="<%= SkinPath %>amskin_r2_c1.gif" class="MainMenu_MenuBar"><dnn:SOLPARTMENU runat="server" id="dnnSOLPARTMENU" CssClass="SelectedTab" menueffectsmenutransition="AlphaFade" menueffectsmenutransitionlength="0.6" /></td>
    <td height="33" align="right" valign="middle" nowrap background="<%= SkinPath %>amskin_r2_c2.gif" class="Footer"><dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="footer" /></td>
  </tr>
  <tr>
    <td width="111" height="111" align="left" valign="top" bgcolor="#314252"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="110" height="110" align="top">
        <param name="movie" value="<%= SkinPath %>White_clock.swf">
        <param name="quality" value="high">
        <param name="BGCOLOR" value="#314252">
        <param name="SCALE" value="exactfit">
        <embed src="<%= SkinPath %>White_clock.swf" width="110" height="110" align="top" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" bgcolor="#314252" scale="exactfit"></embed>
      </object></td>
    <td height="100%" colspan="2" rowspan="2" align="left" valign="top" nowrap class="ContentPane" id="ContentPane" runat="server" visible="false">&nbsp;</td>
  </tr>
  <tr>
    <td width="111" height="100%" align="left" valign="top" nowrap class="Normal" id="LeftPane" runat="server" visible="false"><p class="LeftPane">&nbsp;</p>
    </td>
  </tr>
  <tr>
    <td height="31" colspan="2" align="left" valign="middle" background="<%= SkinPath %>amskin_r4_c1.gif" class="MainMenu_MenuBar"><dnn:BREADCRUMB runat="server" id="dnnBREADCRUMB" /></td>
    <td height="31" background="<%= SkinPath %>amskin_r4_c1.gif" class="Footer">&nbsp;</td>
  </tr>
  <tr>
    <td height="47" colspan="2" align="left" valign="middle" background="<%= SkinPath %>amskin_r5_c1.gif" class="Footer">..:: <dnn:LINKS1 runat="server" id="dnnLINKS1" Separator="&nbsp;&nbsp;|&nbsp;&nbsp;" Level="Root" /> ::..</td>
    <td height="47" align="right" valign="middle" background="<%= SkinPath %>amskin_r5_c1.gif" class="Footer"><div align="right" class="Footer"><dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="OtherTabs" /></div></td>
  </tr>
</table>
</body>
</html>

